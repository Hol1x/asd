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
using System.DirectoryServices.AccountManagement;

namespace designBIB
{
    public partial class frmElever : MetroForm
    {
        public frmElever()
        {
            InitializeComponent();
        }

        public class Row
        {
            public string Fornamn { get; set; }
            public string Efternamn { get; set; }
            public string Klass { get; set; }
            public string ID { get; set; }
        }

        private void Print(Row obj)
        {
            Console.WriteLine(obj.ToString());
        }

        private void Print()
        {
            Console.WriteLine();
        }
        
       
        async void Activatee()
        {
            metroProgressBar1.Value = 20;
            dataGridView1.DataSource = await Task.Run(() => Loader());
            metroProgressBar1.Value = 100;
        }
        private DataTable Loader() {
            using (FileStream fs = new FileStream(@"elever.xml",
               FileMode.Open, FileAccess.ReadWrite, FileShare.Read)) {
                XDocument xDoc = XDocument.Load(fs);

                List<Row> items = (from r in xDoc.Elements("DocumentElement").Elements("Row")
                                   select new Row
                                   {
                                       Fornamn = (string)r.Element("Fornamn") + "",
                                       Efternamn = (string)r.Element("Efternamn") + "",
                                       Klass = (string)r.Element("Klass"),
                                       ID = (string)r.Element("ID")

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

        async private void frmElever_Load(object sender, EventArgs e)
        {
            //await. Loader();
            Activatee();
        }


        private void Save_button_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void Save()

        {
            string path = @"elever.xml";
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
                        ((DataTable)dataGridView1.DataSource).Merge(WorksheetToDataTable(ws, chkHasHeader.Checked));
                        //dataGridView1.DataSource = WorksheetToDataTable(ws, chkHasHeader.Checked);
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

        private void checker_Click(object sender, EventArgs e)
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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("ID LIKE '%{0}%'", textBox1.Text);
        }
        public static bool CheckUserinAD(string domain, string username)
        {
            using (var domainContext = new PrincipalContext(ContextType.Domain, domain)) {
                using (var user = new UserPrincipal(domainContext)) {
                    user.SamAccountName = username;

                    using (var pS = new PrincipalSearcher()) {
                        pS.QueryFilter = user;

                        using (PrincipalSearchResult<Principal> results = pS.FindAll()) {
                            if (results != null && results.Count() > 0) {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }
        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            string headerText =
    dataGridView1.Columns[e.ColumnIndex].HeaderText;

            // Abort validation if cell is not in the CompanyName column.
            if (!headerText.Equals("ID")) return;

            //CheckUserinAD("learnet.se",e.FormattedValue.ToString());
            // Confirm that the cell is not empty.
            if (string.IsNullOrEmpty(e.FormattedValue.ToString())) {
                dataGridView1.Rows[e.RowIndex].ErrorText =
                    "Company Name must not be empty";
                MessageBox.Show("Det går inte att lägga till en användare ut en unik identifierare (Schoolsoft användarnamn)", "Error",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
            }
        }
    }
}
