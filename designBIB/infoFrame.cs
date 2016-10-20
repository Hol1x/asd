using Library.Forms;
using LinqToExcel;
using MetroFramework.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace designBIB
{
    public partial class infoFrame : MetroForm
    {
       

        public infoFrame()
        {
            InitializeComponent();
        }

        public class Row
        {
            public string ID { get; set; }

            public string Titel { get; set; }

            public string Namn { get; set; }

            public string Klass { get; set; }

            public string TimeStamp { get; set; }
        }

        private void infoFrame_Load(object sender, EventArgs e)
        {
            EditMode(false);
            Reload();
            metroToggle1.Enabled = false;
        }

        private void Reload()
        {
            using (FileStream fs = new FileStream(@"dataset.xml",
                   FileMode.Open, FileAccess.ReadWrite, FileShare.Read)) {
                XDocument xDoc = XDocument.Load(fs);

                List<Row> items = (from r in xDoc.Elements("DocumentElement").Elements("Row")
                                   select new Row
                                   {
                                       ID = (string)r.Element("ID") + "",
                                       Titel = (string)r.Element("Titel") + "",
                                       Namn = (string)r.Element("Namn") + "",
                                       Klass = (string)r.Element("Klass") + "",
                                       TimeStamp = (string)r.Element("TimeStamp") + ""
                                   }).ToList();

                fs.SetLength(0);
                xDoc.Save(fs);
                items.ForEach(Print);
                var list = new BindingList<Row>(items);
                ListtoDataTableConverter converter = new ListtoDataTableConverter();
                DataTable dt = converter.ToDataTable(list);
                dataGridView1.DataSource = dt;
            }
        }

        private void Print(Row obj)
        {
            Console.WriteLine(obj.ToString());
        }

        private void Print()
        {
            Console.WriteLine();
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

        private void metroButton1_Click(object sender, EventArgs e)
        {
            Console.WriteLine("pressed");
            var results = dataGridView1.SelectedRows
                           .Cast<DataGridViewRow>()
                           .Select(x => Convert.ToString(x.Cells[0].Value));
            var result = results.ToArray();
            foreach (string value in result) {
                Console.WriteLine(value);
                using (FileStream fs = new FileStream(@"dataset.xml",
                   FileMode.Open, FileAccess.ReadWrite, FileShare.Read)) {
                    XmlDocument xDoc = new XmlDocument();
                    xDoc.Load(fs);
                    foreach (XmlNode xNode in xDoc.SelectNodes("DocumentElement/Row"))
                        if (xNode.SelectSingleNode("ID").InnerText == value.Trim()) {
                            xNode.ParentNode.RemoveChild(xNode);
                        }
                    fs.SetLength(0);
                    xDoc.Save(fs);
                }
            }
            foreach (DataGridViewRow item in this.dataGridView1.SelectedRows) {
                dataGridView1.Rows.RemoveAt(item.Index);
            }

            // Check if the Starting Balance row is included in the selected rows

            //  using (FileStream fs = new FileStream(@"tester.xml",
            //      FileMode.Open, FileAccess.ReadWrite, FileShare.Read)) {
            //  XmlDocument xDoc = new XmlDocument();
            // xDoc.Load(fs);
            //foreach (XmlNode xNode in xDoc.SelectNodes("borrowed/Borrowed_info"))
            //  if (xNode.SelectSingleNode("bok_id").InnerText == metroTextBox1.Text) {
            //    xNode.ParentNode.RemoveChild(xNode);
            //  metroLabel1.Text = metroLabel1.Text + " Är nu återlämnad!";
            // }
            //fs.SetLength(0);
            // xDoc.Save(fs);
            //   }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("Namn LIKE '%{0}%'", textBox1.Text);
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            // Check which column is selected, otherwise set NewColumn to null.
            DataGridViewColumn newColumn =
                dataGridView1.Columns.GetColumnCount(
                DataGridViewElementStates.Selected) == 1 ?
                dataGridView1.SelectedColumns[0] : null;

            DataGridViewColumn oldColumn = dataGridView1.SortedColumn;
            ListSortDirection direction;

            // If oldColumn is null, then the DataGridView is not currently sorted.
            if (oldColumn != null) {
                // Sort the same column again, reversing the SortOrder.
                if (oldColumn == newColumn &&
                    dataGridView1.SortOrder == SortOrder.Ascending) {
                    direction = ListSortDirection.Descending;
                }
                else {
                    // Sort a new column and remove the old SortGlyph.
                    direction = ListSortDirection.Ascending;
                    oldColumn.HeaderCell.SortGlyphDirection = SortOrder.None;
                }
            }
            else {
                direction = ListSortDirection.Ascending;
            }

            // If no column has been selected, display an error dialog  box.
            if (newColumn == null) {
                MessageBox.Show("Select a single column and try again.",
                    "Error: Invalid Selection", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else {
                dataGridView1.Sort(newColumn, direction);
                newColumn.HeaderCell.SortGlyphDirection =
                    direction == ListSortDirection.Ascending ?
                    SortOrder.Ascending : SortOrder.Descending;
            }
        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridViewColumn newColumn = dataGridView1.Columns[e.ColumnIndex];
            DataGridViewColumn oldColumn = dataGridView1.SortedColumn;
            ListSortDirection direction;

            // If oldColumn is null, then the DataGridView is not sorted.
            if (oldColumn != null) {
                // Sort the same column again, reversing the SortOrder.
                if (oldColumn == newColumn &&
                    dataGridView1.SortOrder == SortOrder.Ascending) {
                    direction = ListSortDirection.Descending;
                }
                else {
                    // Sort a new column and remove the old SortGlyph.
                    direction = ListSortDirection.Ascending;
                    oldColumn.HeaderCell.SortGlyphDirection = SortOrder.None;
                }
            }
            else {
                direction = ListSortDirection.Ascending;
            }

            // Sort the selected column.
            dataGridView1.Sort(newColumn, direction);
            newColumn.HeaderCell.SortGlyphDirection =
                direction == ListSortDirection.Ascending ?
                SortOrder.Ascending : SortOrder.Descending;
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            // Put each of the columns into programmatic sort mode.
            foreach (DataGridViewColumn column in dataGridView1.Columns) {
                column.SortMode = DataGridViewColumnSortMode.Programmatic;
            }
        }

        private void Save_button_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void Save()

        {
            string path = @"dataset.xml";
            DataTable ds = dataGridView1.DataSource as DataTable;
            ds.WriteXml(path);
        }

        private void metroToggle1_CheckedChanged(object sender, EventArgs e)
        {
            EditMode(metroToggle1.Checked);
        }

        private void EditMode(bool enabled)
        {
            Console.WriteLine(enabled);
            if (!enabled) {
                ManualAdd_ID.Enabled = false;
                ManualAdd_Namn.Enabled = false;
                ManualAdd_Klass.Enabled = false;
                ManualAdd_Tid.Enabled = false;
                ManualAdd_Titel.Enabled = false;
                ManualAdd_AddButton.Enabled = false;
            }
            else if (enabled) {
                ManualAdd_ID.Enabled = true;
                ManualAdd_Namn.Enabled = true;
                ManualAdd_Klass.Enabled = true;
                ManualAdd_Tid.Enabled = true;
                ManualAdd_Titel.Enabled = true;
                ManualAdd_AddButton.Enabled = true;
            }
        }

        private void ManualAdd_Tid_Click(object sender, EventArgs e)
        {
            DateTime date = DateTime.Now;
            ManualAdd_Tid.Text = GetTimestamp(date);
        }

        private String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyy/MM/dd/ HH:mm:ss:fff");
        }

        private void ManualAdd_AddButton_Click(object sender, EventArgs e)
        {
            this.dataGridView1.Rows.Add(ManualAdd_ID.Text, ManualAdd_Titel.Text, ManualAdd_Namn.Text, ManualAdd_Klass.Text, ManualAdd_Tid.Text);
        }

        private void Paste_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            //bindingsource.DataSource = null;
            //bindingsource.Clear();
            try {
                string s = Clipboard.GetText();
                string[] lines = s.Split('\n');
                int iFail = 0, iRow = dataGridView1.CurrentCell.RowIndex;
                int iCol = dataGridView1.CurrentCell.ColumnIndex;
                DataGridViewCell oCell;
                if (dataGridView1.Rows.Count < lines.Length)
                    dataGridView1.Rows.Add(lines.Length - 1);
                foreach (string line in lines) {
                    if (iRow < dataGridView1.RowCount && line.Length > 0) {
                        string[] sCells = line.Split('\t');
                        for (int i = 0; i < sCells.GetLength(0); ++i) {
                            if (iCol + i < this.dataGridView1.ColumnCount) {
                                oCell = dataGridView1[iCol + i, iRow];
                                if (!oCell.ReadOnly) {
                                    if (oCell.Value == null || oCell.Value.ToString() != sCells[i]) {
                                        oCell.Value = Convert.ChangeType(sCells[i],
                                                              oCell.ValueType);
                                        //  oCell.Style.BackColor = Color.Tomato;
                                    }
                                    else
                                        iFail++;
                                    //only traps a fail if the data has changed
                                    //and you are pasting into a read only cell
                                }
                            }
                            else { break; }
                        }
                        iRow++;
                    }
                    else { break; }
                    if (iFail > 0)
                        MessageBox.Show(string.Format("{0} updates failed due" +
                                        " to read only column setting", iFail));
                }
            }
            catch (FormatException) {
                MessageBox.Show("The data you pasted is in the wrong format for the cell");
                return;
            }
        }

        private void metroLabel5_Click(object sender, EventArgs e)
        {
        }

        private void metroLabel1_Click(object sender, EventArgs e)
        {

        }
    }
}