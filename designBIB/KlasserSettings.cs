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
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace designBIB
{
    public partial class KlasserSettings : MetroForm
    {
        public KlasserSettings()
        {
            InitializeComponent();
        }

        public class Row
        {
            public string Klasser { get; set; }
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

        private void KlasserSettings_Load(object sender, EventArgs e)
        {
           
                using (FileStream fs = new FileStream(@"klasser.xml",
                   FileMode.Open, FileAccess.ReadWrite, FileShare.Read)) {
                    XDocument xDoc = XDocument.Load(fs);

                    List<Row> items = (from r in xDoc.Elements("DocumentElement").Elements("Row")
                                       select new Row
                                       {
                                           Klasser = (string)r.Element("Klasser") + "",
                                           ID = (string)r.Element("ID")+""
                                           
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
        

        private void Save_button_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void Save()

        {
            string path = @"klasser.xml";
            DataTable ds = dataGridView1.DataSource as DataTable;
            ds.WriteXml(path);
        }

        private void btnCheckKlasser_Click(object sender, EventArgs e)
        {
            var xmlfile = @"‪produkter.xml";
            var doc = XDocument.Load(xmlfile);
            //var node = doc.Descendants().Where(n => n.Value == metroTextBox1.Text);
            //metroComboBox2.DataSource = node.ToList();
            var itemType = doc.Root.Elements("Row")
                   .Where(i => string.IsNullOrWhiteSpace((string)i.Element("Owner")))
                   ;

            foreach (var ex in itemType) {
                Console.WriteLine(ex.ToString());
            }
                
            
        }
    }
}
