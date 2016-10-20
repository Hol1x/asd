using LinqToExcel;
using MetroFramework.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;


namespace designBIB
{
    public partial class frmService : MetroForm
    {
        public frmService()
        {
            InitializeComponent();
        }

        public class Row
        {
            public string Servicenummer { get; set; }
            public string Servicestalle { get; set; }
            public string Kontaktinformation { get; set; }
            public string Serienummer { get; set; }
            public string Anmalningsdatum { get; set; }
            public string Leveransdatum { get; set; }
            public string User { get; set; }
            public string Felbestrivning { get; set; }
            public string Atgard { get; set; }
            public string Skickad { get; set; }
            public string Fardig { get; set; }
        }

        private void Print(Row obj)
        {
            Console.WriteLine(obj.ToString());
        }

        private void Print()
        {
            Console.WriteLine();
        }

        private void frmService_Load(object sender, EventArgs e)
        {

            using (FileStream fs = new FileStream(@"service.xml",
               FileMode.Open, FileAccess.ReadWrite, FileShare.Read)) {
                XDocument xDoc = XDocument.Load(fs);

                List<Row> items = (from r in xDoc.Elements("DocumentElement").Elements("Row")
                                   select new Row
                                   {
                                       Servicenummer = (string)r.Element("Servicenummer") + "",
                                       Servicestalle = (string)r.Element("Servicestalle") + "",
                                       Kontaktinformation = (string)r.Element("Kontaktinformation"),
                                       Serienummer = (string)r.Element("Serienummer"),
                                       Anmalningsdatum = (string)r.Element("Anmalningsdatum"),
                                       Leveransdatum = (string)r.Element("Leveransdatum"),
                                       User = (string)r.Element("User"),
                                       Felbestrivning = (string)r.Element("Felbeskrivning"),
                                       Atgard = (string)r.Element("Atgard"),
                                       Skickad = (string)r.Element("Skickad"),
                                       Fardig = (string)r.Element("Fardig")

                                   }).ToList();

                fs.SetLength(0);
                xDoc.Save(fs);
                items.ForEach(Print);
                var list = new BindingList<Row>(items);
                ListtoDataTableConverter converter = new ListtoDataTableConverter();
                DataTable dt = converter.ToDataTable(list);
                dataGridView1.DataSource = dt;
                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("Fardig LIKE '%{0}%'", "Unchecked");
            }
        }


        private void Save_button_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void Save()

        {
            string path = @"service.xml";
            DataTable ds = dataGridView1.DataSource as DataTable;
            ds.WriteXml(path);
        }
        private DataTable WorksheetToDataTable(ExcelWorksheet ws, bool hasHeader = true)
        {
            DataTable dt = new DataTable(ws.Name);
            int totalCols = ws.Dimension.End.Column;
            int totalRows = ws.Dimension.End.Row;
            int startRow = hasHeader ? 2 : 1;
            ExcelRange wsRow;
            DataRow dr;
            foreach (var firstRowCell in ws.Cells[1, 1, 1, totalCols]) {
                dt.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }

            for (int rowNum = startRow; rowNum <= totalRows; rowNum++) {
                wsRow = ws.Cells[rowNum, 1, rowNum, totalCols];
                dr = dt.NewRow();
                foreach (var cell in wsRow) {
                    dr[cell.Start.Column - 1] = cell.Text;
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }

        private void btnCheckKlasser_Click(object sender, EventArgs e)
        {

            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel File (*.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK) {
                try {
                    // Create an ExcelPackage from file
                    using (var pck = new ExcelPackage(new FileInfo(openFileDialog.FileName))) {
                        // Get the first worksheet
                        ExcelWorksheet ws = pck.Workbook.Worksheets.First();
                        // Convert the worksheet to a DataTable and set it as data source of a DataGridView
                        dataGridView1.DataSource = WorksheetToDataTable(ws, chkHasHeader.Checked);
                    }
                }
                catch (Exception ex) {
                    MessageBox.Show("Importing data from Excel file failed. Exception: " + ex.Message, "Error");
                }
            }
        }

        private void chkHasHeader_CheckedChanged(object sender, EventArgs e)
        {

        }

        

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Console.WriteLine("pressed");
            var results = dataGridView1.SelectedRows
                           .Cast<DataGridViewRow>()
                           .Select(x => Convert.ToString(x.Cells[0].Value));
            var result = results.ToArray();
            foreach (string value in result) {
                Console.WriteLine(value);
            }
            List<string> SelectedRows = new List<string>();
            Console.WriteLine(SelectedRows.Count);
            foreach (DataGridViewRow r in dataGridView1.SelectedRows) {
                SelectedRows.Add(r.Cells[0].Value.ToString());
            }
            foreach (string value in SelectedRows) {
                Console.WriteLine(value);
            }
        }

        void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            
        }

        void dataGridView1_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            
        }

        private void metroToggle1_CheckedChanged(object sender, EventArgs e)
        {
            if (metroToggle1.Checked) {
                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("Fardig LIKE '%{0}%'", "Unchecked");
            }
            else (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("Fardig LIKE '%{0}%'", "Checked");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("Serienummer LIKE '%{0}%'", textBox1.Text);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("User LIKE '%{0}%'", textBox2.Text);
        }
    }
}
