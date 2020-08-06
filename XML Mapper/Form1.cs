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
using System.Xml;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace XML_Mapper
{
    public partial class Form1 : Form
    {
        List<parametersClass> parametersList;
        String sourceFolderPath;
        String targetFolderPath;
        int fileNumber = 0;
        int totalFiles = 0;
        int currentFile = 0;
        int totalRows = 0;
        int dataRows = 0;
        String unwantedCharacters = String.Empty;
        String[] filePaths;
        bool next = false;
        public Form1()
        {
            InitializeComponent();
            //folderPath = "C:\Users\Salman Ahmad\source\repos\XML Mapper\XML Mapper\bin\Debug";
        }
        public Form1(String sourcePath, string targetPath)
        {
            InitializeComponent();
            sourceFolderPath = sourcePath;
            targetFolderPath = targetPath;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //create a instance for the Excel object  
            Excel.Application oExcel = new Excel.Application();

            //specify the file name where its actually exist  
            string filepath = "param.xlsx";

            String exePath = Environment.CurrentDirectory;
            //pass that to workbook object  
            Excel.Workbook WB = oExcel.Workbooks.Open(Path.Combine(exePath, filepath));

            Excel.Worksheet wks = (Excel.Worksheet)WB.Worksheets[1];
            //Get the used Range
            Excel.Range usedRange = wks.UsedRange;

            int rowsNumbers = 0;
            for (int i = 2; i < 10000; i++)
            {
                String temp=wks.Rows[i].Cells[1].Text;
                if (temp==null||temp==String.Empty)
                {
                    rowsNumbers = i;
                    break;
                }
            }
            totalRows = rowsNumbers - 1;
            dataRows = rowsNumbers - 2;
            unwantedCharacters = wks.Rows[2].Cells[5].Text;
            parametersList = new List<parametersClass>();
            for (int i = 2; i <= totalRows; i++)
            {
                String originalName = wks.Rows[i].Cells[1].Text;
                String updatedName= wks.Rows[i].Cells[2].Text;
                String mandatory = wks.Rows[i].Cells[3].Text;

                parametersClass paramObj = new parametersClass();
                paramObj.originalName = originalName;
                paramObj.updatedName = updatedName;
                if (mandatory == String.Empty)
                {
                    paramObj.mandatory = false;
                }
                else
                {
                    paramObj.mandatory = true;
                }
                parametersList.Add(paramObj);
            }
            WB.Close();
            KillExcel(oExcel);
            System.Threading.Thread.Sleep(100);
            fileNumber = 0;
            totalFiles = Directory.EnumerateFiles(sourceFolderPath, "*.xml").Count();
            string userName = Environment.UserName;
            userNameLabel.Text = userName;
            filePaths = Directory.GetFiles(sourceFolderPath, "*.xml");
            dataGridView1.Rows.Clear();
            if (filePaths.Length > 0)
            {
                String filePath = filePaths[currentFile];
                tempRunner(filePath);
            }
        }

        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);
        private static void KillExcel(Excel.Application theApp)
        {
            int id = 0;
            IntPtr intptr = new IntPtr(theApp.Hwnd);
            System.Diagnostics.Process p = null;
            try
            {
                GetWindowThreadProcessId(intptr, out id);
                p = System.Diagnostics.Process.GetProcessById(id);
                if (p != null)
                {
                    p.Kill();
                    p.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("KillExcel:" + ex.Message);
            }
        }

        private void tempRunner(String filePath)
        {
            dataGridView1.Rows.Clear();
            try
            {

                using (XmlReader reader = XmlReader.Create(filePath))
                {
                    fileNumber++;
                    xmlFilesLabel.Text = Convert.ToString(fileNumber) + " of " + Convert.ToString(totalFiles);
                    reader.MoveToContent();
                    String fieldName = "";
                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.Element)
                        {
                            if (reader.IsStartElement())
                            {
                                //return only when you have START tag  
                                switch (reader.Name.ToString())
                                {
                                    case "field_name":
                                        fieldName = reader.ReadString();
                                        break;
                                    case "field_value":
                                        String fieldValue = reader.ReadString();
                                        mapper(ref fieldName, ref fieldValue);
                                        fieldName = "";
                                        break;
                                }
                            }
                        }
                    }
                    if (!Directory.Exists(sourceFolderPath + "\\original"))
                    {
                        Directory.CreateDirectory(sourceFolderPath + "\\original");
                    }

                }
            }
            catch (XmlException exc)
            {
                MessageBox.Show("Make sure that the XML File is Correct");
            }
            catch (Exception exc)
            {
                MessageBox.Show("Unexpected error occured" + exc.Message);
            }
            dataGridView1.CurrentCell = dataGridView1[1, 0];
        }

        private void mapper(ref string fieldName, ref string fieldValue)
        {
            if (fieldName == "ReceiveDate")
            {
                receiveDateLabel.Text = fieldValue;
            }
            else if (fieldName == "Form_Type")
            {
                formTypeLabel.Text = fieldValue;
            }
            else
            {
                if(fieldName== "Task_ID")
                {
                    return;
                }
                String finalFieldValue = String.Empty;
                for(int i = 0; i < parametersList.Count; i++)
                {
                    parametersClass obj = parametersList.ElementAt(i);
                    if (obj.originalName == fieldName)
                    {
                        if (obj.updatedName != String.Empty)
                        {
                            fieldName = obj.updatedName;
                            finalFieldValue = fieldValue;
                        }
                        else
                        {
                            finalFieldValue = fieldValue;
                        }
                    }
                }
                if (finalFieldValue == String.Empty)
                {
                    finalFieldValue = fieldValue;
                }
                bool containsSpecialCharacters = false;
                for(int i = 0; i < unwantedCharacters.Length; i++)
                {
                    if (fieldValue.Contains(unwantedCharacters[i]))
                    {
                        containsSpecialCharacters = true;
                        break;
                    }
                }

                var index = dataGridView1.Rows.Add();
                dataGridView1.Rows[index].Cells["UpdatedName"].Value = finalFieldValue;
                dataGridView1.Rows[index].Cells["OriginalName"].Value = fieldName;
                if (containsSpecialCharacters)
                {
                    dataGridView1.Rows[index].Cells["UpdatedName"].Style.BackColor = Color.Red;
                }
            }
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                int columnIndex = dataGridView1.CurrentCell.ColumnIndex;
                int rowIndex = dataGridView1.CurrentCell.RowIndex;
                int i = rowIndex + 1;
                if(i < dataGridView1.RowCount)
                {
                    dataGridView1.CurrentCell = dataGridView1[1, i];
                }
                //for (int i = rowIndex+1; i < dataGridView1.RowCount; i++)
                //{
                //    if (dataGridView1.Rows[i].Cells["UpdatedName"].Style.BackColor == Color.Red)
                //    {
                //        dataGridView1.CurrentCell = dataGridView1[1, i];
                //        break;
                //    }
                //}
            }
            if (e.KeyData == Keys.Down)
            {
                int columnIndex = dataGridView1.CurrentCell.ColumnIndex;
                int rowIndex = dataGridView1.CurrentCell.RowIndex;
                int i = rowIndex + 1;
                if (i < dataGridView1.RowCount)
                {
                    dataGridView1.CurrentCell = dataGridView1[1, i];
                }
            }
            if (e.KeyData == Keys.Up)
            {
                int columnIndex = dataGridView1.CurrentCell.ColumnIndex;
                int rowIndex = dataGridView1.CurrentCell.RowIndex;
                int i = rowIndex - 1;
                if (i < dataGridView1.RowCount)
                {
                    dataGridView1.CurrentCell = dataGridView1[1, i];
                }
            }
            e.Handled = true;

        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void SkipToNextButton_Click(object sender, EventArgs e)
        {
            String filePath = filePaths[currentFile];
            dataGridView1.Rows.Clear();
            currentFile++;
            if (currentFile < totalFiles)
            {
                filePath = filePaths[currentFile];
                tempRunner(filePath);
            }
            else
            {
                MessageBox.Show("No more XML files in the source folder");
                                MainForm mainForm = new MainForm();
                this.Hide();
                mainForm.Show();

            }


        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            String filePath = filePaths[currentFile];
            String xmlFileName = Path.GetFileName(filePath);
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(filePath);

            List<parametersClass> newTempList = new List<parametersClass>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                parametersClass obj = new parametersClass();
                if (row.Cells["UpdatedName"].Value != null)
                {
                    obj.originalName = row.Cells["OriginalName"].Value.ToString();
                    obj.updatedName = row.Cells["UpdatedName"].Value.ToString();//value

                }else if(row.Cells["UpdatedName"].Value == null)
                {
                    continue;
                }
                else
                {
                    obj.originalName = row.Cells["OriginalName"].Value.ToString();
                    obj.updatedName = "";//value
                }
                newTempList.Add(obj);
            }

            XmlNodeList dataNodeslist = xdoc.SelectNodes("/names/OCR_properties/data");
            int i = 0;
            foreach(XmlNode xmlNode in dataNodeslist)
            {
                XmlNodeList childNodes = xmlNode.ChildNodes;
                childNodes[0].InnerText = newTempList.ElementAt(i).originalName;
                childNodes[1].InnerText = newTempList.ElementAt(i).updatedName;
                i++;
            }
            dataNodeslist = xdoc.SelectNodes("/names/OCR_properties/data");
            bool userExists = false;
            foreach (XmlNode xmlNode in dataNodeslist)
            {
                XmlNodeList childNodes = xmlNode.ChildNodes;
                if (childNodes[0].InnerText == "XMLUser")
                {
                    string userName = Environment.UserName;
                    childNodes[1].InnerText = userName;
                    userExists = true;
                    break;
                }
            }
            if (!userExists)
            {
                //Writing username in that xml file
                XmlNode OCRPROPERTIES = xdoc.SelectSingleNode("/names/OCR_properties");
                string userName = Environment.UserName;
                XmlNode newElem = xdoc.CreateElement("data");
                XmlNode childNode = xdoc.CreateElement("field_name");
                childNode.InnerText = "XMLUser";
                newElem.AppendChild(childNode);
                XmlNode childNode1 = xdoc.CreateElement("field_value");
                childNode1.InnerText = userName;
                newElem.AppendChild(childNode1);
                OCRPROPERTIES.AppendChild(newElem);

            }

            xdoc.Save(targetFolderPath + "\\" + xmlFileName);

            File.Copy(Path.Combine(sourceFolderPath, xmlFileName), Path.Combine(sourceFolderPath + "\\original", xmlFileName), true);
            File.Delete(filePath);
            dataGridView1.Rows.Clear();
            currentFile++;
            if (currentFile < totalFiles)
            {
                filePath = filePaths[currentFile];
                tempRunner(filePath);
            }
            else
            {
                MessageBox.Show("No more XML files in the source folder");
                MainForm mainForm = new MainForm();
                this.Hide();
                mainForm.Show();
            }

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();

        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            if (totalFiles == 0)
            {
                MessageBox.Show("No XML files present in source folder");
                MainForm mainForm = new MainForm();
                this.Hide();
                mainForm.Show();
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.BeginEdit(true);
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.BeginEdit(true);

        }
    }



}
