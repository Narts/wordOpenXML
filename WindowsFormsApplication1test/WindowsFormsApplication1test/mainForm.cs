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
    public partial class mainForm : Form
    {
        string saved_document;
        string saved_commentar;
        string saved_category;
        public List<string> saved_doc_list = new List<string>();
        public List<string> saved_com_list = new List<string>();
        public List<string> saved_cat_list = new List<string>();
        //public List<string> bookmark_list = new List<string>();
        string saved_path;
        string created_folder;
        string opened_file_name;
        string opened_path;
        bool newSmry;
        Utilities utilities = Utilities.CreateInstance();


        //object strFileName;
        //Object Nothing;
        //Microsoft.Office.Interop.Word.Application myWordApp = new Word.Application();
        //Word.Document myWordDoc;
        //Word.Document anotherWordDoc;  

        public mainForm(string address, bool newSmry)
        {
            InitializeComponent();
            this.newSmry = newSmry;
            if (newSmry)
            {
                this.created_folder = address;
            }
            else 
            {
                this.saved_path = address;
            }
            
        }

        private void InsertButton_Click(object sender, EventArgs e)
        {
            string doc = saved_document;
            string com = saved_commentar;
            string cat = saved_category;
            if (doc == null || com == null || cat == null)
            {
                MessageBox.Show("please insert content, comment and category first.");
            }
            else 
            {
                Clipboard.SetText(richTextBox1.Rtf, TextDataFormat.Rtf);
                saved_doc_list.Add(doc);
                saved_com_list.Add(com);
                saved_cat_list.Add(cat);
                richTextBox1.Clear();
                richTextBox2.Clear();
                richTextBox3.Clear();
            }
           
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            string str = "";
            string strFrmt = "";
            foreach (string line in richTextBox1.Lines)
                str += line;
            foreach (char rtfText in richTextBox1.Rtf)
                strFrmt += rtfText;
            getDocument(str);
        }

        private void BuildSummaryBtn_Click(object sender, EventArgs e)
        {
            if (saved_doc_list.Count == 0 || saved_com_list.Count == 0 || saved_cat_list.Count == 0 || opened_file_name == null)
            {
                MessageBox.Show("Please\n 1. Open file via 'Open File' button;\n 2. Insert content, comment and category via 'Insert Content' button;\n before build a summary.");
            }
            else 
            {
                utilities.setSavedDocList(saved_doc_list);
                utilities.setSavedComList(saved_com_list);
                utilities.setSavedCatList(saved_cat_list);
                if (this.newSmry)
                {
                    MessageBox.Show(this.newSmry.ToString());
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.InitialDirectory = created_folder;
                    saveFileDialog1.Title = "Save Files";
                    //saveFileDialog1.CheckFileExists = true;
                    saveFileDialog1.CheckPathExists = true;
                    saveFileDialog1.DefaultExt = "txt";
                    saveFileDialog1.Filter = "Text files (*.txt)|*.txt|Microsoft Word (*.docx)|*.docx|All files (*.*)|*.*";
                    saveFileDialog1.FilterIndex = 2;
                    saveFileDialog1.RestoreDirectory = true;

                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        saved_path = saveFileDialog1.FileName;
                        MessageBox.Show(saved_path);
                        utilities.createWord(saved_path, this.newSmry);
                    }
                    OpenFileBtn.Enabled = false;
                    InsertBtn.Enabled = false;
                    BuildSummaryBtn.Enabled = false;
                }
                else 
                {
                    MessageBox.Show("newSmry == false");
                    utilities.processWord(saved_path, this.newSmry);
                    OpenFileBtn.Enabled = false;
                    InsertBtn.Enabled = false;
                    BuildSummaryBtn.Enabled = false;
                }

            }
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            string str = "";
            foreach (string line in richTextBox2.Lines)
                str += line;
            getComment(str);
        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {
            string str = "";
            foreach (string line in richTextBox3.Lines)
                str += line;
            getCategory(str);
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
            openFileDialog1.Filter = "All files (*.*)|*.*|Microsoft Word (*.docx)|*.docx|txt files (*.txt)|*.txt";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                String opened_file = openFileDialog1.FileName;
                String[] names = opened_file.Split('\\');
                int len = names.Length;
                String doc_name = names[len - 1];
                opened_file_name = doc_name;
                int len_doc_name = doc_name.Length;
                opened_path = opened_file.Substring(0, opened_file.Length - len_doc_name);
                MessageBox.Show(opened_file);
                utilities.openWordFile(opened_file);
                //System.Diagnostics.Process.Start("WINWORD", opened_file);
                string saved_old_folder = "";
                if (!this.newSmry)
                {
                    int end_path = this.saved_path.LastIndexOf('\\');
                    saved_old_folder = saved_path.Substring(0, end_path);
                }
                else 
                {
                    saved_old_folder = utilities.getCreatedFolder();
                }
                bool cp_success = utilities.copyFileToRepository(doc_name, opened_path, saved_old_folder);
                if (!cp_success) 
                {
                    MessageBox.Show("This file already exists");
                }
                    
                //copy_file_to_reporsitory(doc_name, opened_path, created_folder);
            }
            
        }

        private void getComment(string str)
        {
            saved_commentar = str;
        }

        private void getDocument(string str)
        {
            //MessageBox.Show(utilities.getActiveDocName());
            saved_document = str + "(" + utilities.getActiveDocName() + ")";
        }

        private void getCategory(string str)
        {
            saved_category = str;
        }

        private void NextSummaryBtn_Click(object sender, EventArgs e)
        {
            startForm sf = new startForm();
            this.Visible = false;
            sf.ShowDialog();
            sf.Visible = true;

        }


        //private void copy_file_to_reporsitory(String file_name, String soure_path, String target_path)
        //{
        //    // Use Path class to manipulate file and directory paths.
        //    string sourceFile = System.IO.Path.Combine(soure_path, file_name);
        //    string destFile = "";
        //    // To copy a folder's contents to a new location:
        //    // Create a new target folder, if necessary.
        //    if (!System.IO.Directory.Exists(target_path))
        //    {
        //        System.IO.Directory.CreateDirectory(target_path);
        //    }
           
        //    if (target_path != null) 
        //    {
        //        destFile = System.IO.Path.Combine(target_path, file_name);
        //    }
        //    // check, if the destFile already exists in the repository path
        //    if (!System.IO.File.Exists(destFile))
        //    {
        //        // To copy a file to another location and 
        //        // overwrite the destination file if it already exists.
        //        System.IO.File.Copy(sourceFile, destFile, true);
        //    }
        //    else 
        //    {
        //        MessageBox.Show("This file already exists");
        //    }
        //}

        //private void createWord(string saved_path)
        //{
        //    //strFileName = created_folder + "\\test.docx";
        //    strFileName = saved_path;
        //    MessageBox.Show(strFileName.ToString());
        //    if (System.IO.File.Exists((string)strFileName))
        //        System.IO.File.Delete((string)strFileName);
        //    Object Nothing = System.Reflection.Missing.Value;
        //    myWordDoc = myWordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            
        //    #region   将数据写入到word文件中

        //    int len_doc = saved_doc_list.Count;
        //    for (int i = 0; i < len_doc; i++)
        //    {
        //        strContent = "Zitat:\n\n\r ";
        //        myWordDoc.Paragraphs.Last.Range.Text = strContent;

        //        string doc_content = saved_doc_list[i];
        //        strContent = doc_content + "\n\n\r";
        //        myWordDoc.Paragraphs.Last.Range.Text = strContent;

        //        strContent = "Kommentar:\n\n\r ";
        //        myWordDoc.Paragraphs.Last.Range.Text = strContent;

        //        strContent = saved_com_list[i] + "\n\n\r";
        //        myWordDoc.Paragraphs.Last.Range.Text = strContent;

        //        strContent = "Ressource:\n\n\r ";
        //        myWordDoc.Paragraphs.Last.Range.Text = strContent;
             
        //        int len = doc_content.Length;
        //        int title_pos = doc_content.LastIndexOf("(");
        //        int start_index = title_pos + 1;
        //        string file_name = doc_content.Substring(start_index, len-1-start_index);

        //        AddContent(i, file_name);
        //        strContent = "\n\n\r";
        //        myWordDoc.Paragraphs.Last.Range.Text = strContent;
        //    }

        //    saved_doc_list.Clear();
        //    saved_com_list.Clear();
        //    #endregion

        //    //将WordDoc文档对象的内容保存为DOC文档  
        //    myWordDoc.SaveAs(ref strFileName, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
        //    //关闭WordDoc文档对象  
        //    myWordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
        //    //关闭WordApp组件对象  
        //    //myWordApp.Quit(ref Nothing, ref Nothing, ref Nothing);

        //    MessageBox.Show(strFileName + "\r\n " + "created successfully ");
            
        //}

        //public void AddContent(int index, string file_name)
        //{
        //    try
        //    {

        //        Object oMissing = System.Reflection.Missing.Value;
        //        // Word Interface
        //        //Microsoft.Office.Interop.Word._Application WordApp = new Word.Application();
        //        //WordApp.Visible = true;
        //        //object filename = filePath;
        //        //Microsoft.Office.Interop.Word._Document WordDoc = WordApp.Documents.Open(ref filename, ref oMissing,
        //        //    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
        //        //    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

        //        //设置居左
        //        //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;

        //        //设置文档的行间距
        //        //WordApp.Selection.ParagraphFormat.LineSpacing = 15f;
        //        //插入段落
        //        //WordApp.Selection.TypeParagraph();
        //        //Microsoft.Office.Interop.Word.Paragraph para;
        //        //para = myWordDoc.Content.Paragraphs.Add(ref oMissing);
        //        ////正常格式
        //        //para.Range.Text = "This is paragraph 1";
        //        ////para.Range.Font.Bold = 2;
        //        ////para.Range.Font.Color = WdColor.wdColorRed;
        //        ////para.Range.Font.Italic = 2;
        //        //para.Range.InsertParagraphAfter();

        //        //para.Range.Text = "This is paragraph 2";
        //        //para.Range.InsertParagraphAfter();

        //        //插入Hyperlink
        //        Microsoft.Office.Interop.Word.Selection mySelection = myWordApp.ActiveWindow.Selection;
        //        mySelection.Start = 9999;
        //        mySelection.End = 9999;
                
        //        Microsoft.Office.Interop.Word.Range myRange = mySelection.Range;

        //        Microsoft.Office.Interop.Word.Hyperlinks myLinks = myWordDoc.Hyperlinks;
        //        string file_Path = created_folder + "\\" + file_name;
        //        object linkAddr = file_Path;
        //        string single_bookmark = bookmark_list[index];
        //        object linkSubAddr = single_bookmark;
        //        // you may need more parameters here
        //        Microsoft.Office.Interop.Word.Hyperlink myLink = myLinks.Add(myRange, ref linkAddr, ref linkSubAddr);
        //        myWordApp.ActiveWindow.Selection.InsertAfter("\n");
        //        //object linkAddr = test_file_Path;
        //        //Microsoft.Office.Interop.Word.Hyperlink myLink = myLinks.Add(myRange, ref linkAddr, ref oMissing);
        //        //WordApp.ActiveWindow.Selection.InsertAfter("\n");

        //        //落款
        //        myWordDoc.Paragraphs.Last.Range.Text = "created in：" + DateTime.Now.ToString();
        //        //myWordDoc.Paragraphs.Last.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;

        //        //保存
        //        //myWordDoc.Save();
        //        //myWordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
        //        //myWordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
        //        //return true;
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.Message);
        //        Console.WriteLine(e.StackTrace);
        //        //return false;
        //    }
        //}

        //private void insert_bookmark()
        //{
        //    for (int i = 0; i < saved_doc_list.Count(); i++)
        //    {
        //        string doc_full = saved_doc_list[i];
        //        int len_doc = doc_full.Length;
        //        int title_pos = doc_full.LastIndexOf("(");
        //        string doc_content = doc_full.Substring(0, title_pos);
        //        int startIndex = title_pos+1;
        //        int strLength = len_doc-1-startIndex;
        //        string file_name = doc_full.Substring(startIndex, strLength);

        //            strFileName = created_folder + "\\" + file_name;
        //            //MessageBox.Show(targetDoc.ToString());
        //            //if (System.IO.File.Exists((string)targetDoc))
        //            //System.IO.File.Delete((string)targetDoc);
        //            Object Nothing = System.Reflection.Missing.Value;
        //            anotherWordDoc = myWordApp.Documents.Open(ref strFileName,
        //                                                      ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
        //                                                      ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
        //                                                      ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);

        //            AddBookmarks(doc_content, i, anotherWordDoc, strFileName);

        //            //关闭WordDoc文档对象  
        //            anotherWordDoc.Close(ref   Nothing, ref   Nothing, ref   Nothing);
        //            //关闭WordApp组件对象  
        //            //myWordApp.Quit(ref   Nothing, ref   Nothing, ref   Nothing);

        //            //string strFileName = created_folder + "\\test2.docx";
        //            //object objFileName = @strFileName;
                
        //    }
        //}

        //private void AddBookmarks(string targetDoc, int index, Word.Document anotherWordDoc, object strFileName)
        //{
        //    object findText = targetDoc;
        //    Object Nothing = System.Reflection.Missing.Value;
        //    Word.Range rng = anotherWordDoc.Range();

        //    rng.Find.ClearFormatting();

        //    if (rng.Find.Execute(ref findText,
        //        ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
        //        ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
        //        ref Nothing, ref Nothing))
        //    {
                
        //        // insert_one_bookmark(rng, index, anotherWordDoc);
        //        int start = rng.Start;
        //        int end = rng.End;
        //        object bookmark_rng = anotherWordDoc.Range(start, end);
        //        MessageBox.Show(index.ToString());
        //        MessageBox.Show(start.ToString(), end.ToString());
        //        string bookmark_name = "ST" + start.ToString(); //
        //        bookmark_list.Add(bookmark_name);
        //        anotherWordDoc.Bookmarks.Add(bookmark_name, ref bookmark_rng);
        //        //将WordDoc文档对象的内容保存为DOC文档  
        //        anotherWordDoc.SaveAs(ref   strFileName, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing);

        //    }
        //    else
        //    {
        //        MessageBox.Show("Text not found.");
        //    }
        //}
    }
}
