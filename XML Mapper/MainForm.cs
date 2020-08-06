using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace XML_Mapper
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            String sourcePath = sourceFolderLabel.Text;
            String targetPath = targetFolderLabel.Text;
            if (sourcePath == String.Empty || targetPath == String.Empty)
            {
                String[] lines = { };
                if (File.Exists("Path.txt"))
                {
                    lines = File.ReadAllLines("Path.txt");
                    if (lines.Length >= 1)
                    {
                        sourcePath = lines[0];
                    }if (lines.Length >= 2)
                    {
                        targetPath = lines[1];
                    }
                }
                if (!Directory.Exists(sourcePath) || !Directory.Exists(targetPath))
                {
                    MessageBox.Show("Default paths don't exist");
                }
                else
                {
                    sourceFolderLabel.Text = sourcePath;
                    targetFolderLabel.Text = targetPath;
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //sourceFolderSelection

            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = false;
            folderDlg.SelectedPath = Environment.CurrentDirectory;
            String[] lines = { };
            if (File.Exists("Path.txt"))
            {
                lines = File.ReadAllLines("Path.txt");
                if (!Directory.Exists(lines[0]))
                {
                    folderDlg.SelectedPath = lines[0];
                }
            }
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                String Text = folderDlg.SelectedPath;
                sourceFolderLabel.Text = Text;
                Environment.SpecialFolder root = folderDlg.RootFolder;
                String exePath = Environment.CurrentDirectory;
                String temp = "";
                int i = 0;
                using (StreamWriter writer = new StreamWriter(Path.Combine(exePath, "Path.txt")))
                {
                    writer.WriteLine(Text);
                    //writer.Write(lines[1]);
                }
                //Form1 form1 = new Form1(Text);
                //this.Hide();
                //form1.Show();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            String sourcePath = sourceFolderLabel.Text;
            String targetPath = targetFolderLabel.Text;
            if (sourcePath == String.Empty || targetPath == String.Empty)
            {
                String[] lines = { };
                if (File.Exists("Path.txt"))
                {
                    lines = File.ReadAllLines("Path.txt");
                    if (lines.Length == 2)
                    {
                        sourcePath = lines[0];
                        targetPath = lines[1];
                        if (!Directory.Exists(sourcePath) || !Directory.Exists(targetPath))
                        {
                            MessageBox.Show("The paths are invalid, make sure that correct paths are selected");
                            sourcePath = String.Empty;
                            targetPath = String.Empty;
                        }
                    }
                    else
                    {
                        MessageBox.Show("The paths are invalid, make sure that correct paths are selected");
                        sourcePath = String.Empty;
                        targetPath = String.Empty;
                    }

                }
            }
            if(sourcePath != String.Empty && targetPath != String.Empty){
                Form1 form1 = new Form1(sourcePath, targetPath);
                this.Hide();
                form1.Show();
            }
            

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //targetFolderSelection
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = false;
            folderDlg.SelectedPath = Environment.CurrentDirectory;
            String[] lines = { };
            if (File.Exists("Path.txt"))
            {
                lines = File.ReadAllLines("Path.txt");
                if (lines.Length > 1)
                {
                    if (!Directory.Exists(lines[1]))
                    {
                        folderDlg.SelectedPath = lines[1];
                    }
                }
            }
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                String Text = folderDlg.SelectedPath;
                targetFolderLabel.Text = Text;
                Environment.SpecialFolder root = folderDlg.RootFolder;
                String exePath = Environment.CurrentDirectory;
                int i = 0;

                using (StreamWriter writer =File.AppendText(Path.Combine(exePath, "Path.txt")))
                {
                    //writer.WriteLine(lines[0]);
                    writer.Write(Text);
                }
                //Form1 form1 = new Form1(Text);
                //this.Hide();
                //form1.Show();
            }

        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();

        }
    }
}
