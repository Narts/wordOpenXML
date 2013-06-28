using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1test
{
    public partial class startForm : Form
    {
        Utilities utilities = Utilities.CreateInstance();
        string address;
        bool newSmry;

        public startForm()
        {
            InitializeComponent();
        }

        private void NewSummary_Click(object sender, EventArgs e)
        {
            this.newSmry = true;
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            // Show the FolderBrowserDialog. 
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                address = folderDlg.SelectedPath;
                MessageBox.Show(address);
                utilities.setCreatedFolder(address);
                Environment.SpecialFolder root = folderDlg.RootFolder;
            }
            this.DialogResult = DialogResult.OK; 
        }

        private void ProcessSummary_Click(object sender, EventArgs e)
        {
            this.newSmry = false;
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
                string address = openFileDialog1.FileName;
                this.DialogResult = DialogResult.OK; 
                //String[] names = opened_file.Split('\\');
                //int len = names.Length;
                //String doc_name = names[len - 1];
                //string opened_file_name = doc_name;
                //int len_doc_name = doc_name.Length;
                //string opened_path = opened_file.Substring(0, opened_file.Length - len_doc_name);
                ////MessageBox.Show(opened_path);
                //utilities.openWordFile(opened_file);
                ////System.Diagnostics.Process.Start("WINWORD", opened_file);
                //bool cp_success = utilities.copyFileToRepository(doc_name, opened_path, created_folder);
                //if (!cp_success)
                //{
                //    MessageBox.Show("This file already exists");
                //}

                //copy_file_to_reporsitory(doc_name, opened_path, created_folder);
            }
            
        }

        public string getAddress()
        {
            return this.address;
        }

        public bool getNewSmry()
        {
            return this.newSmry;
        }
    }
}
