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
            //this.DialogResult = DialogResult.OK;
            if (this.getAddress() != null)
            {
                mainForm mf = new mainForm(this.getAddress(), this.getNewSmry());
                this.Visible = false;
                mf.ShowDialog();
                mf.Visible = true;
                this.Close();
            }

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
                this.address = openFileDialog1.FileName;
                int end_path = this.address.LastIndexOf('\\');
                string created_folder = address.Substring(0, end_path);
                utilities.setCreatedFolder(created_folder);
                //this.DialogResult = DialogResult.OK;
                if (this.getAddress() != null)
                {
                    mainForm mf = new mainForm(this.getAddress(), this.getNewSmry());
                    this.Visible = false;
                    mf.ShowDialog();
                    mf.Visible = true;
                    this.Close();
                }
                
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
