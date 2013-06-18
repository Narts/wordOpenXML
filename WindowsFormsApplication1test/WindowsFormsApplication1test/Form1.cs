using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Net;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using MSOT = Microsoft.Office.Tools;
using System.Reflection;
using MSOTW = Microsoft.Office.Tools.Word;


namespace WindowsFormsApplication1test
{
    public partial class Form1 : Form
    {
        String saved_document;
        String saved_commentar;
        List<String> saved_doc_list = new List<String>();
        List<String> saved_com_list = new List<String>();
        String saved_path;
        String created_folder;
        String opened_path;

        object strFileName;
        Object Nothing;
        Microsoft.Office.Interop.Word.Application myWordApp = new Word.Application();
        Word.Document myWordDoc;
        Word.Document anotherWordDoc;
        string strContent = " ";   

        public Form1()
        {
            InitializeComponent();
        }

        private void InsertButton_Click(object sender, EventArgs e)
        {
            String doc = saved_document;
            String com = saved_commentar;
            saved_doc_list.Add(doc);
            saved_com_list.Add(com);
            richTextBox1.Clear();
            richTextBox2.Clear();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            string str = "";
            foreach (string line in richTextBox1.Lines)
                str += line;
            get_document(str);
        }

        private void BuildButton_Click(object sender, EventArgs e)
        {
            String show_doc = "";
            String show_com = "";
            String show_together = "";

            for (int i = 0; i < saved_doc_list.Count(); i++)
            {
                show_doc = saved_doc_list[i];
                show_com = saved_com_list[i];
                show_together = "\n"+ show_doc + "\n" + show_com;
                MessageBox.Show(show_together);
            }
        }

        private void get_document(String str)
        {
            saved_document = str;
        }

        private void SaveFileButton_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = @"C:\";
            saveFileDialog1.Title = "Save text Files";
            saveFileDialog1.CheckFileExists = true;
            saveFileDialog1.CheckPathExists = true;
            saveFileDialog1.DefaultExt = "txt";
            saveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {

                saved_path = saveFileDialog1.FileName;
                MessageBox.Show(saved_path);

            } 

        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void CreateFolderButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            // Show the FolderBrowserDialog. 
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                created_folder = folderDlg.SelectedPath;
                MessageBox.Show(created_folder);
                Environment.SpecialFolder root = folderDlg.RootFolder;
            } 
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            string str = "";
            foreach (string line in richTextBox2.Lines)
                str += line;
            get_commentar(str);
        }

        private void get_commentar(String str)
        {
            saved_commentar = str;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void OpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Open text Files";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.DefaultExt = "txt";
            openFileDialog1.Filter = "All files (*.*)|*.*|txt files (*.txt)|*.txt";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                String opened_file = openFileDialog1.FileName;
                String[] names = opened_file.Split('\\');
                int len = names.Length;
                String doc_name = names[len - 1];
                int len_doc_name = doc_name.Length;
                opened_path = opened_file.Substring(0, opened_file.Length-len_doc_name);
                //MessageBox.Show(opened_path);
                System.Diagnostics.Process.Start("WINWORD",opened_file);
                copy_file_to_reporsitory(doc_name, opened_path, created_folder);
            }
        }

        private void copy_file_to_reporsitory(String file_name, String soure_path, String target_path)
        {
            // Use Path class to manipulate file and directory paths.
            string sourceFile = System.IO.Path.Combine(soure_path, file_name);
            string destFile = System.IO.Path.Combine(target_path, file_name);
            // To copy a folder's contents to a new location:
            // Create a new target folder, if necessary.
            if (!System.IO.Directory.Exists(target_path))
            {
                System.IO.Directory.CreateDirectory(target_path);
            }

            // To copy a file to another location and 
            // overwrite the destination file if it already exists.
            System.IO.File.Copy(sourceFile, destFile, true);

        }

        private void writeInDoc_Click(object sender, EventArgs e)
        {
            createWord();
            AddContent(strFileName.ToString());
        }

        private void createWord()
        {
            //strFileName = System.Windows.Forms.Application.StartupPath + "\\test.doc ";
            strFileName = created_folder + "\\test.docx";
            MessageBox.Show(strFileName.ToString());
            if (System.IO.File.Exists((string)strFileName))
                System.IO.File.Delete((string)strFileName);
            Object Nothing = System.Reflection.Missing.Value;
            myWordDoc = myWordApp.Documents.Add(ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing);

            #region   将数据写入到word文件中

            int len_doc = saved_doc_list.Count;
            for (int i = 0; i < len_doc; i++ )
            {
                strContent = "你好Zitat\n\n\r ";
                myWordDoc.Paragraphs.Last.Range.Text = strContent;

                strContent = saved_doc_list[i] + "\n\n\r";
                myWordDoc.Paragraphs.Last.Range.Text = strContent;

                strContent = "这是测试程序commentar\n\n\r ";
                myWordDoc.Paragraphs.Last.Range.Text = strContent;

                strContent = saved_com_list[i] + "\n\n\r";
                myWordDoc.Paragraphs.Last.Range.Text = strContent;
            }

            saved_doc_list.Clear();
            saved_com_list.Clear();
            #endregion

            //将WordDoc文档对象的内容保存为DOC文档  
            myWordDoc.SaveAs(ref   strFileName, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing);
            //关闭WordDoc文档对象  
            myWordDoc.Close(ref   Nothing, ref   Nothing, ref   Nothing);
            //关闭WordApp组件对象  
            myWordApp.Quit(ref   Nothing, ref   Nothing, ref   Nothing);

            MessageBox.Show(strFileName + "\r\n " + "创建成功 ");
            
        }

        public void AddContent(string filePath)
        {
            try
            {

                Object oMissing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word._Application WordApp = new Word.Application();
                WordApp.Visible = true;
                object filename = filePath;
                Microsoft.Office.Interop.Word._Document WordDoc = WordApp.Documents.Open(ref filename, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                //设置居左
                WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;

                //设置文档的行间距
                WordApp.Selection.ParagraphFormat.LineSpacing = 15f;
                //插入段落
                //WordApp.Selection.TypeParagraph();
                Microsoft.Office.Interop.Word.Paragraph para;
                para = WordDoc.Content.Paragraphs.Add(ref oMissing);
                //正常格式
                para.Range.Text = "This is paragraph 1";
                //para.Range.Font.Bold = 2;
                //para.Range.Font.Color = WdColor.wdColorRed;
                //para.Range.Font.Italic = 2;
                para.Range.InsertParagraphAfter();

                para.Range.Text = "This is paragraph 2";
                para.Range.InsertParagraphAfter();

                //插入Hyperlink
                Microsoft.Office.Interop.Word.Selection mySelection = WordApp.ActiveWindow.Selection;
                mySelection.Start = 9999;
                mySelection.End = 9999;
                Microsoft.Office.Interop.Word.Range myRange = mySelection.Range;

                Microsoft.Office.Interop.Word.Hyperlinks myLinks = WordDoc.Hyperlinks;
                //object linkAddr = @"http://www.cnblogs.com/lantionzy";
                object test_file_Path = created_folder +"\\test2.docx##" + "aora";
                object linkAddr = test_file_Path;

                Microsoft.Office.Interop.Word.Hyperlink myLink = myLinks.Add(myRange, ref linkAddr,
                    ref oMissing);
                WordApp.ActiveWindow.Selection.InsertAfter("\n");

                //落款
                WordDoc.Paragraphs.Last.Range.Text = "文档创建时间：" + DateTime.Now.ToString();
                WordDoc.Paragraphs.Last.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;

                //保存
                WordDoc.Save();
                WordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                WordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
                //return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                //return false;
            }
        }

        public void AddBookmarks(string targetDoc, int index)
        {
            //strFileName = System.Windows.Forms.Application.StartupPath + "\\test.doc ";
            strFileName = created_folder + "\\test2.docx";
            //MessageBox.Show(targetDoc.ToString());
            //if (System.IO.File.Exists((string)targetDoc))
                //System.IO.File.Delete((string)targetDoc);
            Object Nothing = System.Reflection.Missing.Value;
            anotherWordDoc = myWordApp.Documents.Open(ref strFileName,
                                                      ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                                                      ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                                                      ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            object findText = targetDoc;

            Word.Range rng = anotherWordDoc.Range();

            rng.Find.ClearFormatting();

            if (rng.Find.Execute(ref findText,
                ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                ref Nothing, ref Nothing))
            {
                
                // insert_one_bookmark(rng, index, anotherWordDoc);
                int start = rng.Start;
                int end = rng.End;
                object bookmark_rng = anotherWordDoc.Range(start, end);
                MessageBox.Show(index.ToString());
                MessageBox.Show(start.ToString(), end.ToString());
                string bookmark_name = index.ToString();
                anotherWordDoc.Bookmarks.Add(bookmark_name, ref bookmark_rng);
            }
            else
            {
                MessageBox.Show("Text not found.");
            }

            rng.Select(); 
        }

        private void manipulate_word_Click(object sender, EventArgs e)
        {
            String show_doc = "";

            for (int i = 0; i < saved_doc_list.Count(); i++)
            {
                show_doc = saved_doc_list[i];
                AddBookmarks(show_doc, i);
            }

            //string strFileName = created_folder + "\\test2.docx";
            //object objFileName = @strFileName;
            
        }

        public void insert_one_bookmark(object rng, int index, Word.Document anotherWordDoc)
        {
            //Microsoft.Office.Interop.Word.Application myWord = new Microsoft.Office.Interop.Word.Application();
            string bookmark_name =  index.ToString();
            anotherWordDoc.Bookmarks.Add(bookmark_name, ref rng);
            //myWord.ActiveDocument.Bookmarks.Add(bookmark_name, ref rng);

            //MSOT.Word.Bookmark bookmark;
            //bookmark = document.Bookmarks
            
            //document.Bookmarks.Add(bookmark_name, ref rng);
        }
    }
}
