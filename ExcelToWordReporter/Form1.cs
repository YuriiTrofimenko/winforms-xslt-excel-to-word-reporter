using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.IO;
using System.IO.Compression;

using System.Threading;
using System.Diagnostics;

namespace MyCsharpExam
{
    public partial class Form1 : Form
    {
        private XmlNodeList sharedStrings; //все строки которые excel хранит в отдельном документе
        internal BrowserForHelp helpWindow; 

        private List<string> parameters = new List<string>();
        private string xlsxTablePath;
        private string templatePath;

        private string stylesheetPath = @"stylesheet.xsl";

        private string outDir = "out";
        private string outName = "output_file_";
        private string outExtension = ".xml";


        private XmlDocument GetXmlInXslx()
        {
            string extractPath = @".\extract";

            try
            {
                ZipFile.ExtractToDirectory(xlsxTablePath, extractPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка чтения", MessageBoxButtons.OK);
                throw;
            }

            string xmlDataPath = @".\extract\xl\worksheets\sheet1.xml";

            var xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlDataPath);         

            //get params
            var param = new XmlDocument();
            param.Load(@".\extract\xl\sharedStrings.xml");
            //загрузка всех текстовых ячеек
            sharedStrings = param.GetElementsByTagName("t");

            foreach (XmlNode item in param.GetElementsByTagName("t"))
                if (item.InnerText.First() == '{' && item.InnerText.Last() == '}')
                    parameters.Add(item.InnerText.Substring(1, item.InnerText.Length - 2));

            DirectoryInfo di = new DirectoryInfo(extractPath);

            foreach (FileInfo file in di.GetFiles())
                file.Delete();

            foreach (DirectoryInfo dir in di.GetDirectories())
                dir.Delete(true);

            di.Delete();
            return xmlDoc;
        }

        public Form1()
        {
            InitializeComponent();         
        }
        private void Form_Load(object sender, EventArgs e)
        {
            if (!Directory.Exists("out"))
                Directory.CreateDirectory("out");

            progressBar.Minimum = 0;
            progressBar.Step = 1;
        }

        public void Process()
        {
            var data = GetXmlInXslx();
            var template = new XmlDocument();

            try
            {
                template.Load(templatePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка чтения", MessageBoxButtons.OK);
                throw;
            }           

            XslCompiledTransform transform = new XslCompiledTransform();
            transform.Load(stylesheetPath);
            transform.Transform(templatePath, "res.xsl");

            var style = new XslCompiledTransform();
            style.Load("res.xsl");
            File.Delete("res.xsl");

            XsltArgumentList args = new XsltArgumentList();

            int filesQuantity = data.GetElementsByTagName("row").Count;
            int counter = 1;
            progressBar.Maximum = filesQuantity - 1;
            progressBar.Visible = true;
            progressLabel.Visible = true;
            progressBar.Update();
           

            DateTime startTime = DateTime.Now;
            foreach (XmlNode node in data.GetElementsByTagName("row"))
            {
                if (int.Parse(node.Attributes[0].InnerText) < 2)
                    continue;

                //if (int.Parse(node.Attributes[0].InnerText) > 250)
                //{
                //    progressBar.Value = progressBar.Maximum;
                //    break;
                //}
                var stamp = progressBar.Value;
                var children = node.ChildNodes;

                for (int i = 0; i < parameters.Count; ++i)
                {
                    var it = children.Item(i);

                    if (it.Attributes.Count > 1 && it.Attributes.Item(1).InnerText == "s")
                        args.AddParam(parameters[i], "", sharedStrings.Item(int.Parse(it.InnerText)).InnerText);
                    else
                        args.AddParam(parameters[i], "", it.InnerText);                  
                }

                using (StreamWriter writer = new StreamWriter(Path.Combine(outDir, outName + stamp + outExtension)))
                {
                    style.Transform(template, args, writer);
                    writer.Flush();
                }
                args.Clear();
                progressBar.PerformStep();
                ++counter;
                progressLabel.Text = ((int)((counter * 100.0 / filesQuantity)*100))/100.0 +"%";                
                progressLabel.Update();
            }

            lTime.Visible = true;
            lTime.Text = ((int)((DateTime.Now - startTime).TotalSeconds *100)) / 100.0 + " c";
            progressLabel.Text = "100%";
            progressLabel.ForeColor = Color.Green;            
        }

        private void Start_Click(object sender, EventArgs e)
        {
            //Thread process = new Thread(new ThreadStart(Process));
            //process.Start();
            if (templatePath != null && xlsxTablePath != null)
                Process();
            else
                MessageBox.Show("Выберите файл", "Файл не выбран", MessageBoxButtons.OK);
        }

        private void bXlsx_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog { Filter = "Document |*.xlsx;" })
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    xlsxTablePath = dlg.FileName;
                    bXlsx.BackColor = Color.LightGreen;
                }
            }                   
        }

        private void bXml_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog { Filter = "Document |*.xml;" })
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    templatePath = dlg.FileName;
                    bXml.BackColor = Color.LightGreen;
                }
            }
        }

        private void Help_Click(object sender, LinkLabelLinkClickedEventArgs e)
        {
            
            if (helpWindow == null)
            {
                if (sender == linkLabel1)
                    helpWindow = new BrowserForHelp(@"help/xlsx.html");
                else if (sender == linkLabel2)
                    helpWindow = new BrowserForHelp(@"help/xml.html");
                else 
                    helpWindow = new BrowserForHelp(@"help/about.html");

                helpWindow.Owner = this;
                helpWindow.Show();
            }
            else helpWindow.Activate();
        }

        private void About_Click(object sender, EventArgs e) => Help_Click(sender, null);
   
    }
}
