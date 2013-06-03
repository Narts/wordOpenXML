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

namespace WindowsFormsApplication1test
{
    public partial class Form1 : Form
    {
        String saved_document;
        String saved_commentar;
        List<String> saved_doc_list = new List<String>();
        List<String> saved_com_list = new List<String>();
        String saved_path;
        String opened_folder;

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
            saved_doc_list.Clear();
            saved_com_list.Clear();
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
                opened_folder = folderDlg.SelectedPath;
                MessageBox.Show(opened_folder);
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

        private void IEButton_Click(object sender, EventArgs e)
        {
            Process[] current_Process = Process.GetProcessesByName("NOTEPAD");
            foreach(var item in current_Process)
            {
                String full_path = item.Modules[0].FileName;
                String full_path_test = Path.GetFullPath(item.ToString());
                //String str = System.IO.Path.GetDirectoryName(item.ProcessName);
                MessageBox.Show(full_path_test);
            }
        }
    }
}
