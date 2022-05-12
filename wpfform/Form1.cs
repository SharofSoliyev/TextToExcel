using ClosedXML.Excel;
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

namespace wpfform
{
    public partial class Form1 : Form
    {
        public List<Words> Suzlar = new List<Words>();
        public Form1()
        {
            InitializeComponent();
            GetAllData();
        }
        public void GetAllData()
        {

            string line;
            // List<Words> Words = new List<Words>();
            List<string> words = new List<string>();


            string contents = File.ReadAllText(@"C:\Users\cyber\Desktop\a.txt");
            string[] stringSeparators = new string[] { "\r\n" };
            string[] lines = contents.Split(stringSeparators, StringSplitOptions.None);
            foreach (string item in lines)
            {
                string[] wordes = item.Split("-", StringSplitOptions.None);

                Words words1 = new Words();
                words1.english = wordes[0];
                words1.uzbek = wordes[1];
                Suzlar.Add(words1);
            }
            dataGridView1.DataSource = Suzlar;


        }
        public class Words
        {
            public string uzbek { get; set; }
            public string english { get; set; }
        }
        public class ExportToExcel
        {
            public void ExportToExcelData(DataGridView dataGrid)
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                foreach (DataGridViewColumn item in dataGrid.Columns)
                {
                    dt.Columns.Add(item.Name);
                }
                foreach (DataGridViewRow row in dataGrid.Rows)
                {
                    DataRow dataRow = dt.NewRow();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        dataRow[cell.ColumnIndex] = cell.Value;
                    }
                    dt.Rows.Add(dataRow);
                }

                using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel workbook|*.xlsx" })
                {

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            using (XLWorkbook workbook = new XLWorkbook())
                            {
                                workbook.Worksheets.Add(dt, "Statistics");
                                workbook.SaveAs(sfd.FileName);
                            }
                            MessageBox.Show("You have successfully exported your data to an axcel file.",
                                            "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExportToExcel exportToExcel = new ExportToExcel();
            exportToExcel.ExportToExcelData(dataGridView1);
        }
    }
}
