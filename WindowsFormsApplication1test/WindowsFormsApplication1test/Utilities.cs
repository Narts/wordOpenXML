using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApplication1test
{
    class Utilities
    {
        private static Microsoft.Office.Interop.Word.Application word_show = null; //new Word.Application();
        private static Microsoft.Office.Interop.Word.Application word_app = null;
        Word.Document word_wrt;
        List<string> saved_doc_list = new List<string>();
        List<string> saved_com_list = new List<string>();
        List<string> saved_cat_list = new List<string>();
        List<string> bookmark_list = new List<string>();
        string created_folder;

        private volatile static Utilities _instance = null;
        private static readonly object lockHelper = new object();
        private Utilities() { }
        public static Utilities CreateInstance()
        {
            if(_instance == null)
            {
                lock(lockHelper)
                {
                    if(_instance == null)
                        _instance = new Utilities();
                }
            }
            return _instance;
        }

        public static Word.Application createWordApp()
        {
            if (word_app == null)
            {
                word_app = new Word.Application();
            }
            return word_app;
        }

        public static Word.Application createWordShow()
        {
            if (word_show == null)
            {
                word_show = new Word.Application();
            }
            return word_show;
        }

        public bool copyFileToRepository(String file_name, String soure_path, String target_path)
        {
            // Use Path class to manipulate file and directory paths.
            string sourceFile = System.IO.Path.Combine(soure_path, file_name);
            string destFile = "";
            // To copy a folder's contents to a new location:
            // Create a new target folder, if necessary.
            if (!System.IO.Directory.Exists(target_path))
            {
                System.IO.Directory.CreateDirectory(target_path);
            }

            if (target_path != null)
            {
                destFile = System.IO.Path.Combine(target_path, file_name);
            }
            // check, if the destFile already exists in the repository path
            if (!System.IO.File.Exists(destFile))
            {
                // To copy a file to another location and 
                // overwrite the destination file if it already exists.
                System.IO.File.Copy(sourceFile, destFile, true);
                return true;
            }
            else
            {
                return false;
            }
        }

        public void createWord(string saved_path, bool newSmry)
        {
            word_app = createWordApp();
            this.insertBookmark(saved_doc_list);
            object strFileName = saved_path;
            //MessageBox.Show(strFileName.ToString());
            if (System.IO.File.Exists((string)strFileName))
                System.IO.File.Delete((string)strFileName);
            Object Nothing = System.Reflection.Missing.Value;

            word_wrt = word_app.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            object oStyleName = Word.WdBuiltinStyle.wdStyleBodyText;  //"Heading 1";
            word_wrt.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            word_wrt.Paragraphs.Last.Range.set_Style(ref oStyleName);
            string strContent = "Schlüsselwörter \r";
            word_wrt.Paragraphs.Last.Range.Font.Size = 14;
            word_wrt.Paragraphs.Last.Range.Font.Bold = 1;
            word_wrt.Paragraphs.Last.Range.Text = strContent;

            this.writeSammary();

            //将WordDoc文档对象的内容保存为DOC文档  
            word_wrt.SaveAs(ref strFileName, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            //关闭WordDoc文档对象  
            word_wrt.Close(ref Nothing, ref Nothing, ref Nothing);
            bookmark_list.Clear();


//Test code ##
            this.buildTOC(saved_path, newSmry);
//Test code ##


            try 
            {
                word_show.Quit(ref Nothing, ref Nothing, ref Nothing);
                word_app.Quit(ref Nothing, ref Nothing, ref Nothing);
            }
            catch(Exception)
            {
            }

            word_app = null;
            word_show = null;
        }

        public void processWord(string saved_path, bool newSmry)
        {
            word_app = createWordApp();
            this.insertBookmark(saved_doc_list);
            object strFileName = saved_path;
            Object Nothing = System.Reflection.Missing.Value;
            object readOnly = false;
            object isVisible = false;

            word_wrt = word_app.Documents.Open(ref strFileName, ref Nothing, ref readOnly,
                    ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                    ref Nothing, ref Nothing, ref Nothing, ref isVisible, ref Nothing,
                    ref Nothing, ref Nothing, ref Nothing);
            word_wrt.Activate();
            //word_wrt.Paragraphs.Last.Range.Text = "test text" + "\n";//加个结束符(增加一段),否则再次插入的时候就成了替换.
            this.writeSammary();
            //保存
            word_wrt.Save();

   // test code
            this.buildTOC(saved_path, newSmry);
   // test code

            try
            {
                word_show.Quit(ref Nothing, ref Nothing, ref Nothing);
                word_app.Quit(ref Nothing, ref Nothing, ref Nothing);
            }
            catch (Exception)
            {
            }

            word_app = null;
            word_show = null;
        }

        private void writeSammary()
        {
            //Object Nothing = System.Reflection.Missing.Value;
            #region   将数据写入到word文件中
            string strContent = " ";
            
            int len_doc = saved_doc_list.Count;
            for (int i = 0; i < len_doc; i++)
            {
                //word_wrt.Paragraphs.Last.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
        //test code
                //strContent = "test: \r";
                strContent = saved_cat_list[i] + "\r";
                object oStyleName = Word.WdBuiltinStyle.wdStyleHeading1;  //"Heading 1";
                word_wrt.Paragraphs.Last.Range.set_Style(ref oStyleName);
                word_wrt.Paragraphs.Last.Range.Text = strContent;
        //test code

                oStyleName = Word.WdBuiltinStyle.wdStyleBodyText;  //"Heading 1";
                word_wrt.Paragraphs.Last.Range.set_Style(ref oStyleName);
                strContent = "Zitat:\r ";
                word_wrt.Paragraphs.Last.Range.Font.Size = 12;
                word_wrt.Paragraphs.Last.Range.Font.Bold = 1;
                word_wrt.Paragraphs.Last.Range.Text = strContent;

                string doc_content = saved_doc_list[i];
                int len = doc_content.Length;
                int title_pos = doc_content.LastIndexOf("(");

                strContent = doc_content.Substring(0, title_pos) + "\r";
                word_wrt.Paragraphs.Last.Range.Font.Size = 11;
                word_wrt.Paragraphs.Last.Range.Font.Bold = 0;
                word_wrt.Paragraphs.Last.Range.Text = strContent;

                strContent = "Kommentar:\r ";
                word_wrt.Paragraphs.Last.Range.Font.Size = 12;
                word_wrt.Paragraphs.Last.Range.Font.Bold = 1;
                word_wrt.Paragraphs.Last.Range.Text = strContent;

                strContent = saved_com_list[i] + "\r";
                word_wrt.Paragraphs.Last.Range.Font.Size = 11;
                word_wrt.Paragraphs.Last.Range.Font.Bold = 0;
                word_wrt.Paragraphs.Last.Range.Text = strContent;

                strContent = "Quelle:\r ";
                word_wrt.Paragraphs.Last.Range.Font.Size = 12;
                word_wrt.Paragraphs.Last.Range.Font.Bold = 1;
                word_wrt.Paragraphs.Last.Range.Text = strContent;


                int start_index = title_pos + 1;
                string file_name = doc_content.Substring(start_index, len - 1 - start_index);

                addLink(i, file_name);
                strContent = "\r";
                word_wrt.Paragraphs.Last.Range.Text = strContent;
            }

            saved_doc_list.Clear();
            saved_com_list.Clear();
            saved_cat_list.Clear();
            #endregion

            ////将WordDoc文档对象的内容保存为DOC文档  
            //word_wrt.SaveAs(ref strFileName, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            ////关闭WordDoc文档对象  
            //word_wrt.Close(ref Nothing, ref Nothing, ref Nothing);
            //bookmark_list.Clear();
        }

        public void addLink(int index, string file_name)
        {
            try
            {
                Microsoft.Office.Interop.Word.Selection linkSelection = word_app.ActiveWindow.Selection;
                linkSelection.Start = 9999;
                linkSelection.End = 9999;

                Microsoft.Office.Interop.Word.Range linkRange = linkSelection.Range;

                Microsoft.Office.Interop.Word.Hyperlinks bookmarksLinks = word_wrt.Hyperlinks;
                string file_Path = this.created_folder + "\\" + file_name;
                object linkAddr = file_Path;
                string single_bookmark = bookmark_list[index];
                object linkSubAddr = single_bookmark;
                // you may need more parameters here
                Microsoft.Office.Interop.Word.Hyperlink bookmarkLink = bookmarksLinks.Add(linkRange, ref linkAddr, ref linkSubAddr);
                bookmarkLink.Range.Font.Size = 8;
                word_app.ActiveWindow.Selection.InsertAfter("\n");

                //落款
                word_wrt.Paragraphs.Last.Range.Text = "created in：" + DateTime.Now.ToString();
                //myWordDoc.Paragraphs.Last.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;

                //保存
                //myWordDoc.Save();
                //myWordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                //myWordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
                //return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                //return false;
            }
        }

        public void insertBookmark(List<string> doc_list)
        {
            Word.Document word_bck;
            for (int i = 0; i < doc_list.Count(); i++)
            {
                string doc_full = doc_list[i];
                int len_doc = doc_full.Length;
                int title_pos = doc_full.LastIndexOf("(");
                string doc_content = doc_full.Substring(0, title_pos);
                int startIndex = title_pos + 1;
                int strLength = len_doc - 1 - startIndex;
                string file_name = doc_full.Substring(startIndex, strLength);
                object str_File_Name = this.created_folder + "\\" + file_name;
          
                //MessageBox.Show(targetDoc.ToString());
                //if (System.IO.File.Exists((string)targetDoc))
                //System.IO.File.Delete((string)targetDoc);
                Object Nothing = System.Reflection.Missing.Value;
                word_bck = word_app.Documents.Open(ref str_File_Name,
                                                          ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                                                          ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                                                          ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);

                createBookmarks(doc_content, i, word_bck, str_File_Name);

                //关闭WordDoc文档对象  
                word_bck.Close(ref   Nothing, ref   Nothing, ref   Nothing);
                //关闭WordApp组件对象  
                //myWordApp.Quit(ref   Nothing, ref   Nothing, ref   Nothing);

                //string strFileName = created_folder + "\\test2.docx";
                //object objFileName = @strFileName;

            }
        }

        private bool createBookmarks(string targetDoc, int index, Word.Document wordDocPar, object strFileName)
        {
            int docLen = targetDoc.Length;
            string startText ="";
            if (docLen >= 250)
            {
                startText = targetDoc.Substring(0, 250);
            }
            else 
            {
                startText = targetDoc;
            }

            Object Nothing = System.Reflection.Missing.Value;

            Word.Range rng = wordDocPar.Range();
            rng.Find.ClearFormatting();
            object findText = startText;

            if (rng.Find.Execute(ref findText,
                ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                ref Nothing, ref Nothing))
            {
                // insert_one_bookmark(rng, index, anotherWordDoc);
                int start = rng.Start;
                //end = rng.End;
                object bookmark_rng = wordDocPar.Range(start, start + docLen);
                //MessageBox.Show(index.ToString());
                //MessageBox.Show(start.ToString(), end.ToString());
                string bookmark_name = "ST" + start.ToString(); //
                bookmark_list.Add(bookmark_name);
                wordDocPar.Bookmarks.Add(bookmark_name, ref bookmark_rng);
                //将WordDoc文档对象的内容保存为DOC文档  
                wordDocPar.SaveAs(ref   strFileName, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing, ref   Nothing);
                return true;
            }
            else
            {
                return false;
                //MessageBox.Show("Text not found.");
            }
        }

        public string getActiveDocName()
        {
            word_show = createWordShow();
            return word_show.ActiveWindow.Document.Name;
            //return word_app.ActiveDocument.Name;
        }

        public void openWordFile(string str_File_Name)
        {
            word_show = createWordShow();
            object file_Name = str_File_Name;
            Object Nothing = System.Reflection.Missing.Value;
            try
            {
                Word.Document word_op = word_show.Documents.Open(ref file_Name,
                                                              ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                                                              ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                                                              ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            }
            catch (Exception)
            {
                word_show = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application") as Word.Application;
                //word_show = null;
                //createWordShow();
                Word.Document word_op = word_show.Documents.Open(ref file_Name,
                                                              ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                                                              ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                                                              ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            }
            word_show.Visible = true;
        }

        public void buildTOC(string string_File_Name, bool newSmry)
        {
//TODO: try if process old TOC or build new TOC
            word_app = createWordApp();
            object str_File_Name = string_File_Name;  //this.created_folder + "\\" + "summary.docx";
            Object Nothing = System.Reflection.Missing.Value;
            Word.Document word_toc = word_app.Documents.Open(ref str_File_Name,
                                          ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                                          ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                                          ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            if (newSmry)
            {
                Object oTrue = true;
                //SETTING THE OUTLINE LEVEL
                //SELECT THE CONTENTS WHOSE OUTLINE LEVEL NEEDS TO BE CHANGED AND
                //SET THE VALUE
                word_app.Selection.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel2;
                word_app.Selection.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel3;
                word_app.Selection.Paragraphs.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
                // NAME OF THE BOOKMARK IN THE DOCUMENT (.dot Template) WHERE TABLE OF
                // CONTENTS NEEDS TO BE ADDED
                Word.Selection toc_pos = word_app.ActiveWindow.Selection;
                toc_pos.Start = 16;
                toc_pos.End = 16;

                Word.Range rngTOC = toc_pos.Range;

                // SELECTING THE SET RANGE
                rngTOC.Select();
                // INCLUDING THE TABLE OF CONTENTS
                Object oUpperHeadingLevel = "1";
                Object oLowerHeadingLevel = "3";
                Object oTOCTableID = "TableOfContents";
                word_toc.TablesOfContents.Add(rngTOC, ref oTrue, ref oUpperHeadingLevel,
                    ref oLowerHeadingLevel, ref Nothing, ref oTOCTableID, ref oTrue,
                    ref oTrue, ref Nothing, ref oTrue, ref oTrue, ref oTrue);
            }
            else 
            {
                word_toc.TablesOfContents[1].Update();
            }


            word_toc.Save();
        }

        public void setCreatedFolder(string created_folder)
        {
            this.created_folder = created_folder;
        }

        public string getCreatedFolder()
        {
            return this.created_folder;
        }

        public void setSavedDocList(List<string> saved_doc_list)
        {
            this.saved_doc_list = saved_doc_list;
        }

        public List<string> getSavedDocList()
        {
            return this.saved_doc_list;
        }

        public void setSavedComList(List<string> saved_com_list)
        {
            this.saved_com_list = saved_com_list;
        }

        public List<string> getSavedComList()
        {
            return this.saved_com_list;
        }

        public void setSavedCatList(List<string> saved_cat_list)
        {
            this.saved_cat_list = saved_cat_list;
        }

        public List<string> getSavedCatList()
        {
            return this.saved_cat_list;
        }

        public void setBookmarkList(List<string> bookmark_list)
        {
            this.bookmark_list = bookmark_list;
        }

        public List<string> getBookmarkList()
        {
            return this.bookmark_list;
        }
    }
}
