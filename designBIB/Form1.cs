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
using System.Xml;
using System.Xml.Linq;

namespace designBIB
{
    public partial class Form1 : MetroForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Text = Text + " " +typeof(Form1).Assembly.GetName().Version;
            var xmlfile = @"‪elever.xml";
            var doc = XDocument.Load(xmlfile);

            var xmlfile2 = @"‪klasser.xml";
            var doc2 = XDocument.Load(xmlfile2);
            //var node = doc.Descendants().Where(n => n.Value == metroTextBox1.Text);
            //metroComboBox2.DataSource = node.ToList();
            var itemType = from key in doc2.Descendants("Row").Descendants("Klasser")
                           select key.Value;
            var namn = from key in doc.Descendants("record").Descendants("Name")
                       select key.Value + " ";

            metroComboBox1.DataSource = itemType.ToList();
            metroComboBox2.DataSource = namn.ToList();
        }

        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var xmlfile = "‪elever.xml";
            metroProgressBar1.Value = 20;
            var doc = XDocument.Load(xmlfile);

            metroProgressBar1.Value = 40;
            var loadklass = metroComboBox1.SelectedItem.ToString();

            if (doc.Root != null)
            {
                var selectedKlass = doc.Root.Elements("Row")
                    .Where(i => (string)i.Element("Klass") == loadklass).Select(i => (string)i.Element("Fornamn")+" "+(string)i.Element("Efternamn"));

                metroProgressBar1.Value = 80;
                metroComboBox2.DataSource = selectedKlass.ToList();
            }
            metroProgressBar1.Value = 100;
            //Console.Write(loadklass);
            reset_progressbar();
        }

        private void reset_progressbar()
        {
            metroProgressBar1.Value = 0;
        }

        private string getBookByID(string bok)
        {
            var xmlfile = "‪/Bok2.xml";
            var doc = XDocument.Load(xmlfile);
            //var node = doc.Descendants().Where(n => n.Value == metroTextBox1.Text);
            //metroComboBox2.DataSource = node.ToList();
            var itemType = doc.Root.Elements("bok").Elements("recorde")
                   .Where(i => (string)i.Element("nummer") == bok)
                   .Where(i => (string)i.Element("InUse") == "no")
                   .Select(i => (string)i.Element("title"))

                   .FirstOrDefault();
            return itemType;
        }

        private string[] getBookByID(string bok, string nothing)
        {
            var xmlfile = "‪/Bok2.xml";
            var doc = XDocument.Load(xmlfile);
            var node = doc.Descendants().Where(n => n.Value == metroTextBox1.Text);
            //metroComboBox2.DataSource = node.ToList();
            var itemList = node.ToArray();
            var itemType = doc.Root.Elements("bok").Elements("recorde")
                   .Where(i => (string)i.Element("nummer") == bok)
                   .Where(i => (string)i.Element("InUse") == "no")
                   .Select(i => (string)i.Element("title")).ToArray();
            foreach (string i in itemList) {
                Console.Write("{0} ", i);
            }

            return itemType;
        }

        private string getBookByTitle(string bok)
        {
            var xmlfile = "‪/Bok2.xml";
            var doc = XDocument.Load(xmlfile);
            var node = doc.Descendants().Where(n => n.Value == metroTextBox1.Text).ToArray();
            var itemList = node.ToArray();
            //metroComboBox2.DataSource = node.ToList();
            var itemType = doc.Root.Elements("bok").Elements("recorde")
                   .Where(i => (string)i.Element("title") == bok)
                   .Where(i => (string)i.Element("InUse") == "no")
                   .Select(i => (string)i.Element("nummer"))
                   .FirstOrDefault();

            return itemType;
        }

        private bool InUse()
        {
            return false;
        }

        private String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyy/MM/dd/ HH:mm");
        }

        

        private void metroButton1_Click(object sender, EventArgs e)
        {
            DateTime date = DateTime.Now;
            updateCheck();
            
                using (FileStream fs = new FileStream(@"service.xml",
                   FileMode.Open, FileAccess.ReadWrite, FileShare.Read)) {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(fs);
                    XmlNode lanad = doc.CreateElement("Row");
                    XmlNode Servicenummer = doc.CreateElement("Servicenummer");
                    Servicenummer.InnerText = txtServicenummer.Text;
                    XmlNode Servicestalle = doc.CreateElement("Servicestalle");
                    Servicestalle.InnerText = txtServicestalle.Text;
                    XmlNode Kontaktinformation = doc.CreateElement("Kontaktinformation");
                    Kontaktinformation.InnerText = txtKontaktinformation.Text;
                    XmlNode Serienummer = doc.CreateElement("Serienummer");
                    Serienummer.InnerText = txtSerienummer.Text;
                    XmlNode Anmalningsdatum = doc.CreateElement("Anmalningsdatum");
                    Anmalningsdatum.InnerText = txtAnmalingsdatum.Text;
                    XmlNode Leveransdatum = doc.CreateElement("Leveransdatum");
                    Leveransdatum.InnerText = txtLeveransdatum.Text;
                    XmlNode User = doc.CreateElement("User");
                    User.InnerText =  metroComboBox2.SelectedItem.ToString() +" "+ metroComboBox1.SelectedItem.ToString();
                    XmlNode Felbeskrvningxml = doc.CreateElement("Felbeskrivning");
                    Felbeskrvningxml.InnerText = Felbeskrivning.Text;
                    XmlNode Atgardxml = doc.CreateElement("Atgard");
                    Atgardxml.InnerText = Atgard.Text;
                    XmlNode Skickad = doc.CreateElement("Skickad");
                    Skickad.InnerText = chkSkickad.CheckState.ToString();
                    XmlNode Fardig = doc.CreateElement("Fardig");
                    Fardig.InnerText = chkFardig.CheckState.ToString();
                    lanad.AppendChild(Servicenummer);
                    lanad.AppendChild(Servicestalle);
                    lanad.AppendChild(Kontaktinformation);
                    lanad.AppendChild(Serienummer);
                    lanad.AppendChild(Anmalningsdatum);
                    lanad.AppendChild(Leveransdatum);
                    lanad.AppendChild(User);
                    lanad.AppendChild(Felbeskrvningxml);
                    lanad.AppendChild(Atgardxml);
                    lanad.AppendChild(Skickad);
                    lanad.AppendChild(Fardig);
                    doc.DocumentElement.AppendChild(lanad);
                    fs.SetLength(0);
                    doc.Save(fs);
                    doc = null;
                    metroLabel1.Text ="Ärendet är upplagdt";
                
            }
        }

     
        private void metroTextBox1_TextChanged(object sender, EventArgs e)
        {
            //metroLabel1.Text = SearchResults;
            metroID.Text = metroTextBox1.Text;
            //metroTextBox2.Text = SearchResults;
        }

        private string returnedID(string keyword)
        {
            var xmlfile = "‪/Bok2.xml";
            var doc = XDocument.Load(xmlfile);
            //var node = doc.Descendants().Where(n => n.Value == metroTextBox1.Text);
            //metroComboBox2.DataSource = node.ToList();
            var query = doc.Descendants("bok").Descendants("recorde").Descendants("title")
    .Where(x => !x.HasElements &&
                x.Value.IndexOf(keyword, StringComparison.InvariantCultureIgnoreCase) >= 0);
            //foreach (var element in query)
            //Console.WriteLine(query.FirstOrDefault().Value);
            //foreach (var element in books)
            // Console.WriteLine(matches.First().Value);
            if (query != null) {
                if (query.FirstOrDefault() != null) {
                    if (query.FirstOrDefault().Value != null)

                        return query.FirstOrDefault().Value;
                    else
                        return "";
                }
                else return "";
            }
            else
                return "";
        }

        private bool borrowed_check(string bok)
        {
            var xmlfile = "‪dataset.xml";
            var doc = XDocument.Load(xmlfile);
            //var node = doc.Descendants().Where(n => n.Value == metroTextBox1.Text);
            //metroComboBox2.DataSource = node.ToList();
            var itemType = doc.Root.Elements("Row")
                   .Where(i => (string)i.Element("ID") == bok)
                   .Select(i => (string)i.Element("ID"))
                   .FirstOrDefault();
            if (string.IsNullOrWhiteSpace(itemType))
                return false;
            else
                return true;
        }

        private void metroTextBox2_TextChanged(object sender, EventArgs e)
        {
            var SeachResults = returnedID(metroTextBox2.Text);
            metroLabel1.Text = SeachResults;

            metroTextBox1.Text = getBookByTitle(SeachResults);
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            //infoFrame frame = new infoFrame();
            //KlasserSettings frame = new KlasserSettings();
            frmService frame = new frmService();
            frame.Show();
        }
        private void updateCheck() {
            using (FileStream fs = new FileStream(@"service.xml",
                   FileMode.Open, FileAccess.ReadWrite, FileShare.Read)) {
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(fs);
                foreach (XmlNode xNode in xDoc.SelectNodes("DocumentElement/Row"))
                    if (xNode.SelectSingleNode("Serienummer").InnerText == txtSerienummer.Text && xNode.SelectSingleNode("Fardig").InnerText == "Unchecked") {
                        xNode.ParentNode.RemoveChild(xNode);
                        metroLabel1.Text ="update was executed ";
                    }
                fs.SetLength(0);
                xDoc.Save(fs);
            }
        }
        private void metroButton3_Click(object sender, EventArgs e)
        {
            using (FileStream fs = new FileStream(@"service.xml",
                   FileMode.Open, FileAccess.ReadWrite, FileShare.Read)) {
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(fs);
                foreach (XmlNode xNode in xDoc.SelectNodes("DocumentElement/Row"))
                    if (xNode.SelectSingleNode("Serienummer").InnerText == txtSerienummer.Text) {
                        xNode.ParentNode.RemoveChild(xNode);
                        metroLabel1.Text = metroLabel1.Text + " Är nu återlämnad!";
                    }
                fs.SetLength(0);
                xDoc.Save(fs);
            }
        }

        private void metroLabel1_Click(object sender, EventArgs e)
        {
        }

        private void metroTextBox1_Click(object sender, EventArgs e)
        {
        }

        private void metroTextBox2_Click(object sender, EventArgs e)
        {
        }

        private void metroComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

            char[] delimiterChars = { ' ' };

            
            var xmlfile = @"‪service.xml";
            var xmlfile2 = @"produkter.xml";
            var xmlfile3 = @"Elever.xml";
            metroProgressBar1.Value = 20;
            var doc = XDocument.Load(xmlfile);
            var doc1 = XDocument.Load(xmlfile2);
            var doc2 = XDocument.Load(xmlfile3);

            txtServicenummer.Text = "";
            txtServicestalle.Text = "";
            txtKontaktinformation.Text = "";

            txtAnmalingsdatum.Text = "";
            txtLeveransdatum.Text = "";

            Felbeskrivning.Text = "";
            Atgard.Text = "";
            chkSkickad.CheckState = CheckState.Unchecked;
            chkFardig.CheckState = CheckState.Unchecked;

            metroProgressBar1.Value = 40;
            var loadklass = metroComboBox2.SelectedItem.ToString() +" "+ metroComboBox1.SelectedItem.ToString();
            var elevNamn = metroComboBox2.SelectedItem.ToString();

            var Servicenummer = doc.Root.Elements("Row").Where(i => (string)i.Element("User") == loadklass).Where(i => (string)i.Element("Fardig")== "Unchecked").Select(i => (string)i.Element("Servicenummer")).FirstOrDefault();
            var Servicestalle = doc.Root.Elements("Row").Where(i => (string)i.Element("User") == loadklass).Where(i => (string)i.Element("Fardig") == "Unchecked").Select(i => (string)i.Element("Servicestalle")).FirstOrDefault();
            var Kontaktinformation = doc.Root.Elements("Row").Where(i => (string)i.Element("User") == loadklass).Where(i => (string)i.Element("Fardig") == "Unchecked").Select(i => (string)i.Element("Kontaktinformation")).FirstOrDefault();
            var Serienummer = doc.Root.Elements("Row").Where(i => (string)i.Element("User") == loadklass).Where(i => (string)i.Element("Fardig") == "Unchecked").Select(i => (string)i.Element("Serienummer")).FirstOrDefault();
            if (string.IsNullOrWhiteSpace(Serienummer)) {
                string[] words = elevNamn.Split(delimiterChars);
                var ID = "";
                System.Diagnostics.Debug.WriteLine(words.Length);
                if (words.Length > 2) {
                
                    ID = doc2.Root.Elements("Row").Where(i => (string)i.Element("Fornamn") == words[0]).Where(i => (string)i.Element("Efternamn") == words[1] + " " + words[2].Trim()).Where(i => (string)i.Element("Klass") == metroComboBox1.SelectedItem.ToString()).Select(i => (string)i.Element("ID")).FirstOrDefault();
                }
                else {
                    ID = doc2.Root.Elements("Row").Where(i => (string)i.Element("Fornamn") == words[0]).Where(i => (string)i.Element("Efternamn") == words[1]).Where(i => (string)i.Element("Klass") == metroComboBox1.SelectedItem.ToString()).Select(i => (string)i.Element("ID")).FirstOrDefault();
                }
                    Serienummer = doc1.Root.Elements("Row").Where(i => (string)i.Element("Owner") == ID).Select(i => (string)i.Element("Serienummer")).FirstOrDefault(); 
            }
            var Anmalningsdatum = doc.Root.Elements("Row").Where(i => (string)i.Element("User") == loadklass).Where(i => (string)i.Element("Fardig") == "Unchecked").Select(i => (string)i.Element("Anmalningsdatum")).FirstOrDefault();
            var Leveransdatum = doc.Root.Elements("Row").Where(i => (string)i.Element("User") == loadklass).Where(i => (string)i.Element("Fardig") == "Unchecked").Select(i => (string)i.Element("Leveransdatum")).FirstOrDefault();
            var User = doc.Root.Elements("Row").Where(i => (string)i.Element("User") == loadklass).Where(i => (string)i.Element("Fardig") == "Unchecked").Select(i => (string)i.Element("User")).FirstOrDefault();
            var FelbeskrivningValue = doc.Root.Elements("Row").Where(i => (string)i.Element("User") == loadklass).Where(i => (string)i.Element("Fardig") == "Unchecked").Select(i => (string)i.Element("Felbeskrivning")).FirstOrDefault();
            var AtgardValue = doc.Root.Elements("Row").Where(i => (string)i.Element("User") == loadklass).Where(i => (string)i.Element("Fardig") == "Unchecked").Select(i => (string)i.Element("Atgard")).FirstOrDefault();
            var Skickad = doc.Root.Elements("Row").Where(i => (string)i.Element("User") == loadklass).Where(i => (string)i.Element("Fardig") == "Unchecked").Select(i => (string)i.Element("Skickad")).FirstOrDefault();
            var Fardig = doc.Root.Elements("Row").Where(i => (string)i.Element("User") == loadklass).Where(i => (string)i.Element("Fardig") == "Unchecked").Select(i => (string)i.Element("Fardig")).FirstOrDefault();
            metroProgressBar1.Value = 80;
            txtSerienummer.Text = Serienummer;
            if (Fardig == "Unchecked") {
                txtServicenummer.Text = Servicenummer;
                txtServicestalle.Text = Servicestalle;
                txtKontaktinformation.Text = Kontaktinformation;
                
                txtAnmalingsdatum.Text = Anmalningsdatum;
                txtLeveransdatum.Text = Leveransdatum;
                Felbeskrivning.Text = FelbeskrivningValue;
                Atgard.Text = AtgardValue;
                if (Skickad == "Checked")
                    chkSkickad.CheckState = CheckState.Checked;
                else chkSkickad.CheckState = CheckState.Unchecked;
                //if (Fardig == "Checked")
                //    chkFardig.CheckState = CheckState.Checked;
                //else chkFardig.CheckState = CheckState.Unchecked;

                metroProgressBar1.Value = 100;
            }//Console.Write(loadklass);
                reset_progressbar();
            
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            //infoFrame frame = new infoFrame();
            //KlasserSettings frame = new KlasserSettings();
            frmProdukter frame = new frmProdukter();
            frame.Show();
        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            //infoFrame frame = new infoFrame();
            KlasserSettings frame = new KlasserSettings();
            //frmProdukter frame = new frmProdukter();
            frame.Show();
        }

        private void metroButton6_Click(object sender, EventArgs e)
        {
            frmElever form = new frmElever();
            form.Show();
        }

        private void txtAnmalingsdatum_Click(object sender, EventArgs e)
        {

            txtAnmalingsdatum.Text = GetTimestamp(DateTime.Now);
        }

        private void metroLabel15_Click(object sender, EventArgs e)
        {

        }
    }
}