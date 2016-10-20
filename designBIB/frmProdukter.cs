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
    public partial class frmProdukter : MetroForm
    {
        public frmProdukter()
        {
            InitializeComponent();
        }

        public class Row
        {
            public string Serienummer { get; set; }
            public string Modell { get; set; }
            public string Marke { get; set; }
            public string Owner { get; set; }
        }

        private void Print(Row obj)
        {
            Console.WriteLine(obj.ToString());
        }

        private void Print()
        {
            Console.WriteLine();
        }
        private DataTable Loader()
        {
            using (FileStream fs = new FileStream(@"produkter.xml",
               FileMode.Open, FileAccess.ReadWrite, FileShare.Read)) {
                XDocument xDoc = XDocument.Load(fs);

                List<Row> items = (from r in xDoc.Elements("DocumentElement").Elements("Row")
                                   select new Row
                                   {
                                       Serienummer = (string)r.Element("Serienummer") + "",
                                       Modell = (string)r.Element("Modell") + "",
                                       Marke = (string)r.Element("Marke"),
                                       Owner = (string)r.Element("Owner")

                                   }).ToList();

                fs.SetLength(0);
                xDoc.Save(fs);
                items.ForEach(Print);
                var list = new BindingList<Row>(items);
                ListtoDataTableConverter converter = new ListtoDataTableConverter();
                DataTable dt = converter.ToDataTable(list);
                return dt;
                }
            }
        async void Activatee() {
            dataGridView1.DataSource = await Task.Run(() => Loader());
        }

        async private void frmProdukter_Load(object sender, EventArgs e)
        {


                Activatee();
            /*    var xmlfile = @"‪produkter.xml";
                var doc = XDocument.Load(xmlfile);
                //var node = doc.Descendants().Where(n => n.Value == metroTextBox1.Text);
                //metroComboBox2.DataSource = node.ToList();
                var itemType = doc.Root.Elements("Row")
                       .Where(i => string.IsNullOrWhiteSpace((string)i.Element("Owner")))
                       ;
                int counter = 0;
                foreach (var ex in itemType) {
                    counter++;
                }
                txtLediaDatorer.Text = counter.ToString();

    */
            
        }


        private void Save_button_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void Save()

        {
            string path = @"produkter.xml";
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
        bool ButtonPressed = false;
        private void checker_Click(object sender, EventArgs e)
        {
            ButtonPressed = !ButtonPressed;
            if(ButtonPressed)
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("Owner Is Null");
            else (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Empty;
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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("Serienummer LIKE '%{0}%'", textBox1.Text);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("Owner LIKE '%{0}%'", textBox2.Text);
        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            string headerText =
            dataGridView1.Columns[e.ColumnIndex].HeaderText;

            // Abort validation if cell is not in the CompanyName column.
            if (!headerText.Equals("Owner")) return;

            // Confirm that the cell is not empty.
            var xmlfile3 = @"Elever.xml";
            var doc2 = XDocument.Load(xmlfile3);
            var ID = doc2.Root.Elements("Row").Where(i => (string)i.Element("ID") == e.FormattedValue.ToString()).Select(i => (string)i.Element("ID")).FirstOrDefault();
            if (string.IsNullOrEmpty(e.FormattedValue.ToString())) {
                return;
            }
            
            else if (string.IsNullOrEmpty(doc2.Root.Elements("Row").Where(i => (string)i.Element("ID") == e.FormattedValue.ToString()).Select(i => (string)i.Element("ID")).FirstOrDefault())) {
                dataGridView1.Rows[e.RowIndex].ErrorText =
                    "Stämmer inte överens med något användar ID, kolla stavningen eller lägg till användare";
                MessageBox.Show("Stämmer inte överens med något användar ID, kolla stavningen eller lägg till användare", "Error",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
            }
        }

        private void metroLabel2_Click(object sender, EventArgs e)
        {

        }

        private void metroLabel3_Click(object sender, EventArgs e)
        {

        }
    }
}
