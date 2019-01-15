using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace EasyWord.UI
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        public class ReplaceData
        {
            public string From { get; set; } = "";
            public string To { get; set; } = "";
        }
        public MainWindow()
        {
            InitializeComponent();
            gridReplace.ItemsSource = ReplaceDatas;
        }

        private List<ReplaceData> ReplaceDatas { get; set; } = new List<ReplaceData>();

        private object Nothing = System.Reflection.Missing.Value;
        private string SelectFile()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Word97-2003文档|*.doc|Word文档|*.docx|Excel97-2003文档|*.xls|Excel文档|*.xlsx";
            dialog.Multiselect = false;
            if (!dialog.ShowDialog().GetValueOrDefault())
            {
                return null;
            }
            return dialog.FileName;
        }
        private void btnReplace_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(File_Replace))
            {
                MessageBox.Show("请选择文件!");
                return;
            }
            if (MessageBox.Show("确定替换" + File_Replace + "?", "确定", MessageBoxButton.OKCancel) != MessageBoxResult.OK)
            {
                return;
            }
            busyCtl.IsBusy = true;

            Dictionary<string, string> replaceDic = new Dictionary<string, string>();
            for (int i = 0; i < ReplaceDatas.Count; i++)
            {
                var data = ReplaceDatas[i];
                if (string.IsNullOrEmpty(data.From))
                {
                    continue;
                }
                if (replaceDic.Keys.Contains(data.From))
                {
                    MessageBox.Show("第" + (i + 1) + "行查找的内容与之前的行重复。");
                    continue;
                }
                replaceDic.Add(data.From.Replace(Environment.NewLine, "^p"), data.To.Replace(Environment.NewLine, "^p"));
            }
            //if (replaceDic.Count == 0)
            //{
            //    MessageBox.Show("没有有效的查找替换内容！");
            //    return;
            //}

            BackgroundWorker worker = new BackgroundWorker();
            worker.DoWork += Worker_DoWork;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            ReplacePara para = new ReplacePara()
            {
                ReplaceDatas = replaceDic,
                FilePath = File_Replace,
                All = chkSelectDir.IsChecked.GetValueOrDefault(),
                FileNameFrom = txtFileFrom.Text.Trim(),
                FileNameTo = txtFileTo.Text.Trim()
            };
            worker.RunWorkerAsync(para);
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            busyCtl.IsBusy = false;
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
                return;
            }
            if ((bool)e.Result)
            {
                MessageBox.Show("替换完毕!");
            }
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            Dictionary<string, string> replaceDatas;//被替换、替换成的文本
            bool all;//是否同目录全部替换
            string fileFrom, fileTo;
            ReplacePara para = e.Argument as ReplacePara;
            replaceDatas = para.ReplaceDatas;
            all = para.All;
            fileFrom = para.FileNameFrom;
            fileTo = para.FileNameTo;

            List<FileInfo> lstFileInfo = new List<FileInfo>();
            if (all)
            {
                FileInfo file = new FileInfo(para.FilePath);
                DirectoryInfo dir = file.Directory;
                FileInfo[] files = dir.GetFiles("*.*", SearchOption.AllDirectories);
                if (files != null && files.Length > 0)
                {
                    foreach (var f in files)
                    {
                        if (f.IsReadOnly || (f.Attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
                        {
                            continue;
                        }
                        if (f.Name.EndsWith(".doc") || f.Name.EndsWith(".docx") || f.Name.EndsWith(".xls") || f.Name.EndsWith(".xlsx"))
                        {
                            lstFileInfo.Add(f);
                        }
                    }
                }
            }
            else
            {
                lstFileInfo.Add(new FileInfo(para.FilePath));
            }

            Word.Application wordApp = new Word.ApplicationClass();
            Excel.Application excelApp = new Excel.ApplicationClass();
            //设置为不可见
            wordApp.Visible = false;
            excelApp.Visible = false;
            foreach (FileInfo file in lstFileInfo)
            {
                string newPath = null;
                if (!string.IsNullOrEmpty(fileFrom))
                {
                    newPath = System.IO.Path.Combine(file.DirectoryName, file.Name.Replace(fileFrom, fileTo));
                }
                try
                {
                    if (file.Name.EndsWith(".doc") || file.Name.EndsWith(".docx"))
                        WordReplace(replaceDatas, file.FullName, newPath, wordApp);
                    if (file.Name.EndsWith(".xls") || file.Name.EndsWith(".xlsx"))
                        ExcelReplace(replaceDatas, file.FullName, newPath, excelApp);
                }
                catch (Exception ex)
                {
                    Log.Error("替换失败！", ex);
                }
            }
            //Close App Component
            try { wordApp.Quit(ref Nothing, ref Nothing, ref Nothing); } catch { }
            try { excelApp.Quit(); } catch { }
            e.Result = true;
        }

        private void WordReplace(Dictionary<string, string> replaceDic, object objFile, object newFile, Word.Application app)
        {
            Word.Document doc = doc = app.Documents.Open(ref objFile,
                            ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                            ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                            ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            bool delete = false;//是否删除源文件
            try
            {
                object replace = Word.WdReplace.wdReplaceAll;
                string from, to;
                if (replaceDic.Count > 0)
                {
                    foreach (var item in replaceDic)
                    {
                        from = item.Key;
                        to = item.Value;
                        app.Selection.Find.Replacement.ClearFormatting();
                        app.Selection.Find.ClearFormatting();
                        app.Selection.Find.Text = from;//需要被替换的文本
                        app.Selection.Find.Replacement.Text = to;//替换文本 
                        app.Selection.Find.Execute(
                        ref Nothing, ref Nothing,
                        ref Nothing, ref Nothing,
                        ref Nothing, ref Nothing,
                        ref Nothing, ref Nothing, ref Nothing,
                        ref Nothing, ref replace,
                        ref Nothing, ref Nothing,
                        ref Nothing, ref Nothing);//执行替换操作
                    }
                }

                if (newFile == null || newFile.ToString() == objFile.ToString())
                {
                    doc.Save();
                }
                else
                {
                    doc.SaveAs2(newFile);
                    delete = true;

                }
            }
            catch (Exception ex)
            {
                Log.Error("保存失败！", ex);
            }
            finally
            {
                try { doc.Close(); } catch { }
                if (delete)
                {
                    try
                    {
                        File.Delete(objFile.ToString());
                    }
                    catch (Exception ex)
                    {
                        Log.Error("删除失败！", ex);
                    }
                }
            }
        }

        private void ExcelReplace(Dictionary<string, string> replaceDic, string objFile, string newFile, Excel.Application app)
        {
            Excel.Workbook ew = app.Workbooks.Open(objFile, Nothing, Nothing, Nothing, Nothing, Nothing,
                            Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);
            bool delete = false;//是否删除源文件
            try
            {
                Excel.Worksheet ews;
                int iEWSCnt = ew.Worksheets.Count;
                Excel.Range oRange;

                foreach (var item in replaceDic)
                {
                    for (int i = 1; i <= iEWSCnt; i++)
                    {
                        ews = (Excel.Worksheet)ew.Worksheets[i];
                        oRange = ews.UsedRange.Find(
                        item.Key, Nothing, Nothing,
                        Nothing, Nothing, Excel.XlSearchDirection.xlNext,
                        Nothing, Nothing, Nothing);
                        if (oRange != null && oRange.Cells.Rows.Count >= 1 && oRange.Cells.Columns.Count >= 1)
                        {
                            oRange.Replace(item.Key, item.Value, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);
                            ew.Save();
                        }
                    }
                }

                if (newFile == null || newFile.ToString() == objFile.ToString())
                {
                    ew.Save();
                }
                else
                {
                    ew.SaveAs(newFile);
                    delete = true;
                }
            }
            catch (Exception ex)
            {
                Log.Error("保存失败！", ex);
            }
            finally
            {
                try { ew.Close(); } catch { }
                if (delete)
                {
                    try
                    {
                        File.Delete(objFile.ToString());
                    }
                    catch (Exception ex)
                    {
                        Log.Error("删除失败！", ex);
                    }
                }
            }
        }

        private string File_Replace = "";

        private void btnSelectFile_Replace_Click(object sender, RoutedEventArgs e)
        {
            File_Replace = SelectFile();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            busyCtl.IsBusy = true;
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            ReplaceData data = new ReplaceData() { From = "1", To = "1" };
            ReplaceDatas.Add(data);
            gridReplace.ItemsSource = null;
            gridReplace.ItemsSource = ReplaceDatas;
        }
    }
}
