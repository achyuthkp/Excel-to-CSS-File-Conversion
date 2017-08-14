using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel;

namespace ExcelConvertertoCSV
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataSet result = new DataSet();

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = "";
            fileName = textBox3.Text;

            if (fileName == "")
            {
                MessageBox.Show("Enter Valid file name");
                return;
            }

            convertToCSV(comboBox1.SelectedIndex);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string Chosen_File = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Chosen_File = openFileDialog1.FileName;
            }
            if (Chosen_File == String.Empty)
            {
                return;
            }
            textBox1.Text = Chosen_File;

            getExcelData(textBox1.Text);
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result = this.folderBrowserDialog1.ShowDialog();
            string foldername = "";
            if (result == DialogResult.OK)
            {
                foldername = this.folderBrowserDialog1.SelectedPath;
            }
            
            textBox2.Text = foldername;
        }

        private void getExcelData(string file)
        {
            FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = file.Contains(".xlsx")
                      ? ExcelReaderFactory.CreateOpenXmlReader(stream)
                      : ExcelReaderFactory.CreateBinaryReader(stream);
            excelReader.IsFirstRowAsColumnNames = false;
            result = excelReader.AsDataSet();
            excelReader.Close();
            List<string> items = new List<string>();
            for (int i = 0; i < result.Tables.Count; i++)
                items.Add(result.Tables[i].TableName.ToString());
            comboBox1.DataSource = items;

        }

        private void convertToCSV(int ind)
        {
            string First_Table = "";
            string Second_Table = "";
            int row_Location = 0;
            if (checkBox1.Checked)
            {
                while (row_Location < result.Tables[ind].Rows.Count)
                {
                    string temp_First = "";
                    string temp_Second = "";
                    Boolean csvFlag = true;
                    for (int i = 0; i < result.Tables[ind].Columns.Count; i++)
                    {
                        // to switch appending the 2 strings
                        if (string.IsNullOrEmpty(result.Tables[ind].Rows[row_Location][i].ToString()))
                        {
                            csvFlag = false;
                            continue;
                        }

                        if (csvFlag)
                        {
                            temp_First += "\"" + result.Tables[ind].Rows[row_Location][i].ToString() + "\"" + ",";
                        }

                        else
                        {
                            temp_Second += "\"" + result.Tables[ind].Rows[row_Location][i].ToString() + "\"" + ",";
                        }
                       
                    }
                    // to avoid appending empty rows into the strings
                    if(!temp_First.Equals(";;;"))
                    {
                        First_Table += temp_First;
                    }
                    if (!temp_Second.Equals(";;;"))
                    {
                        Second_Table += temp_Second;
                    }

                    row_Location++;
                    First_Table += "\n";
                    Second_Table += "\n";
                }
                string output1 = textBox2.Text + "\\" + textBox3.Text + ".csv";
                string output2 = textBox2.Text + "\\" + textBox4.Text + ".csv";
                StreamWriter csv1 = new StreamWriter(@output1, false);
                StreamWriter csv2 = new StreamWriter(@output2, false);
                csv1.Write(First_Table);
                csv1.Close();
                csv2.Write(Second_Table);
                csv2.Close();
            }
            else
            {
                while (row_Location < result.Tables[ind].Rows.Count)
                {
                    for (int i = 0; i < result.Tables[ind].Columns.Count; i++)
                    {
                        First_Table += "\"" + result.Tables[ind].Rows[row_Location][i].ToString() + "\"" + ",";
                    }
                    row_Location++;
                    First_Table += "\n";
                }
                string output1 = textBox2.Text + "\\" + textBox3.Text + ".csv";
                StreamWriter csv = new StreamWriter(@output1, false);
                csv.Write(First_Table);
                csv.Close();
            }

            MessageBox.Show("File converted successfully");
            
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            label6.Visible = false;
            textBox4.Visible = false;
            comboBox1.DataSource = null;
            checkBox1.Checked = false;
            return;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                label6.Visible = true;
                textBox4.Visible = true;
            }
            else
            {
                label6.Visible = false;
                textBox4.Visible = false;
            }
        }
    }
}