using System;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

namespace ExcelEditor
{
    public partial class Form1 : Form
    {
        string path = "", TitleName = "ExcelEditor";
        ExcelPackage excel;//Install-Package EPPlus

        public Form1()
        {
            InitializeComponent();
            Text = TitleName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string coord = textBox1.Text;
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[listBox1.SelectedItem.ToString()];
                textBox2.Text = worksheet?.Cells[coord]?.Value?.ToString() ?? "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            excel?.Dispose();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string coord = textBox1.Text;
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[listBox1.SelectedItem.ToString()];
                worksheet.Cells[coord].Value = (object)textBox2.Text;
                excel.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                path = openFileDialog1.FileName;
                Text = openFileDialog1.SafeFileName + " - " + TitleName;
                excel = new ExcelPackage(new FileInfo(path));
                listBox1.Items.Clear();
                for (int i = 0; i < excel.Workbook.Worksheets.Count; i++)
                    listBox1.Items.Add(excel.Workbook.Worksheets[i + 1].Name);
                listBox1.SelectedIndex = 0;
            }
        }
    }
}
