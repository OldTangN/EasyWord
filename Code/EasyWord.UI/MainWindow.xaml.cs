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
            dialog.Filter = "Word文档|*.docx|Word97-2003文档|*doc";
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
                FileInfo[] doc97 = dir.GetFiles(".doc", SearchOption.AllDirectories);
                FileInfo[] doc07 = dir.GetFiles(".docx", SearchOption.AllDirectories);
                if (doc97 != null && doc97.Length > 0)
                {
                    lstFileInfo.AddRange(doc97);
                }
                if (doc07 != null && doc07.Length > 0)
                {
                    lstFileInfo.AddRange(doc07);
                }
            }
            else
            {
                lstFileInfo.Add(new FileInfo(para.FilePath));
            }
            Word.Application app = new Word.ApplicationClass();
            //设置为不可见
            app.Visible = false;
            try
            {
                foreach (FileInfo file in lstFileInfo)
                {
                    if (string.IsNullOrEmpty(fileFrom))
                    {
                        WordReplace(replaceDatas, file.FullName, null, app);
                    }
                    else
                    {
                        string newPath = System.IO.Path.Combine(file.DirectoryName, file.Name.Replace(fileFrom, fileTo));
                        WordReplace(replaceDatas, file.FullName, newPath, app);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("替换失败！", ex);
                throw ex;
            }
            finally
            {
                //Close wordApp Component
                app.Quit(ref Nothing, ref Nothing, ref Nothing);
            }
            e.Result = true;
        }

        private void WordReplace(Dictionary<string, string> replaceDic, object objFile, object newFile, Word.Application app)
        {
            Word.Document doc = doc = app.Documents.Open(ref objFile,
                            ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                            ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                            ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
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
            catch (Exception ex)
            {
                Log.Error("保存失败！", ex);
            }
            finally
            {
                try { doc.Close(); } catch { }
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
