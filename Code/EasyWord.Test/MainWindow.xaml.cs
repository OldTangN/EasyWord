using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
using MSWord = Microsoft.Office.Interop.Word;
//using Word = Microsoft.Office.Tools.Word;
using EasyWord.Core;

namespace EasyWord.Test
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        string TemplateFile = "";

        private void btnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            TemplateFile = "";
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Word文档|*.docx|Word97-2003文档|*doc";
            dialog.Multiselect = false;
            if (!dialog.ShowDialog().GetValueOrDefault())
            {
                return;
            }
            TemplateFile = dialog.FileName;
            Load(TemplateFile);
        }

        object Nothing = System.Reflection.Missing.Value;

        void Save(string templateFile)
        {
            
            //创建一个Word应用程序实例 
            MSWord.Application app = new MSWord.ApplicationClass();
            //无法嵌入互操作类型“Microsoft.Office.Interop.Word.ApplicationClass”。请改用适用的接口
            //将dll属性中的“嵌入互操作类型”的值改为“false”即可
            try
            {
                FileInfo fileInfo = new FileInfo(templateFile);
                //设置为不可见
                app.Visible = false;
                //模板文件地址，这里假设在根目录  
                string templatepath = fileInfo.FullName;
                object path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, string.Format("{0:yyyy-MM-dd HH-mm-ss}{1}", DateTime.Now, fileInfo.Extension));
                System.IO.File.Copy(templatepath, path.ToString());
                //以模板为基础生成文档
                MSWord.Document doc = app.Documents.Add(ref path);
                //获取书签数组  
                foreach (MSWord.Bookmark item in doc.Bookmarks)
                {
                    BookMark mark = lstBookMarks.FirstOrDefault(p => p.Name == item.Name);
                    if (mark != null)
                    {
                        item.Range.Text = mark.Value;
                    }

                    #region example
                    /*  if (item.Name == "Name")
                    {
                        item.Range.Text = "Old.T";
                    }
                    else if (item.Name == "Birthday")
                    {
                        item.Range.Text = "2000.01.01";
                    }
                    else if (item.Name == "WorkYears")
                    {
                        item.Range.Text = "1";
                    }
                    else if (item.Name == "TelPhone")
                    {
                        item.Range.Text = "121345678912";
                    }
                    else if (item.Name == "Email")
                    {
                        item.Range.Text = "123@456.com";
                    }
                    else
                    {

                    }
                    */
                    #endregion
                }
                if (fileInfo.Extension == ".docx")
                {
                    doc.SaveAs(path, MSWord.WdSaveFormat.wdFormatDocumentDefault, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                }
                else
                {
                    doc.SaveAs(path, MSWord.WdSaveFormat.wdFormatDocument, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                }
               
                doc.Close();
                MessageBox.Show("1111");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //Close wordApp Component
                app.Quit(ref Nothing, ref Nothing, ref Nothing);
            }
            //WdSaveFormat is Word 2003 Format
            //object format = MSWord.WdSaveFormat.wdFormatDocument;
            //doc.Close(true, ref Nothing, ref Nothing);
            //object path2 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.doc");
            //doc.SaveAs(ref path2, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            //doc.Close(ref Nothing, ref Nothing, ref Nothing);
            //app.Quit(ref Nothing, ref Nothing, ref Nothing);
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            Save(TemplateFile);
        }

        List<BookMark> lstBookMarks = new List<BookMark>();

        void Load(string templateFile)
        {
            object objFile = templateFile;
           
            MSWord.Application app = new MSWord.ApplicationClass();
            try
            {
                MSWord.Document doc = app.Documents.Add(ref objFile);
                foreach (MSWord.Bookmark item in doc.Bookmarks)
                {
                    lstBookMarks.Add(new BookMark(item.Name));
                }
                doc.Close();
            }
            catch (Exception)
            {
            }
           finally
            {
                //Close wordApp Component
                app.Quit(ref Nothing, ref Nothing, ref Nothing);
            }
            dgBookMarks.ItemsSource = lstBookMarks;

        }
    }
}
