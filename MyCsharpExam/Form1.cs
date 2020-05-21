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
        private XmlDocument template;
        private XslCompiledTransform style;
        private XmlDocument data;

        private XmlNodeList sharedStrings; //все строки которые excel хранит в отдельном документе
        internal BrowserForHelp helpWindow;

        private CancellationTokenSource cancelToken;
        private bool Cancel;

        private List<string> parameters = new List<string>();
        private string xlsxTablePath;
        private string templatePath;

        private string stylesheetPath = @"stylesheet.xsl";

        private string outDir = "out";
        private string outName = "output_file_";
        private string outExtension = ".xml";


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

        private bool IsFileLocked(string filePath)
        {
            FileStream stream = null;
            try
            {
                stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message, "Ошибка чтения", MessageBoxButtons.OK);
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            return false;
        }

        private XmlDocument GetXmlInXslx()
        {
            string extractPath = @".\extract";

            ZipFile.ExtractToDirectory(xlsxTablePath, extractPath);

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

        public void StartJob()
        {
            data = GetXmlInXslx();
            template = new XmlDocument();
            template.Load(templatePath);

            XslCompiledTransform transform = new XslCompiledTransform();
            transform.Load(stylesheetPath);
            transform.Transform(templatePath, "res.xsl");

            style = new XslCompiledTransform();
            style.Load("res.xsl");
            File.Delete("res.xsl");

            progressBar.Maximum = data.GetElementsByTagName("row").Count - 1;
            progressBar.Value = 0;
            progressBar.Visible = true;
            progressLabel.Visible = true;
            progressLabel.ForeColor = Color.Green;
            progressBar.Update();

            Task.Factory.StartNew(() => ProcessFilesMulty());

            //ProcessFilesSingle();                               
        }

        private void  ProcessFilesMulty()
        {
            try
            {
                int filesQuantity = data.GetElementsByTagName("row").Count;
                int count = -1;

                DateTime startTime = DateTime.Now;
                ParallelOptions parOpts = new ParallelOptions { CancellationToken = cancelToken.Token };

                Parallel.ForEach((from XmlNode n in data.GetElementsByTagName("row") select n).ToList(), parOpts, (node) =>
                {
                    XsltArgumentList args = new XsltArgumentList();
                    int stamp = Interlocked.Increment(ref count);

                    if (stamp < 2)
                        return;

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

                    if (stamp % 20 == 0)
                    {
                        Invoke((Action)delegate
                        {
                            progressBar.Value = stamp;
                            progressLabel.Text = ((int)((stamp * 100.0 / filesQuantity) * 100)) / 100.0 + "%";
                            progressLabel.Update();
                        });
                    }
                });

                Invoke((Action)delegate
                {
                    lTime.Visible = true;
                    lTime.Text = ((int)((DateTime.Now - startTime).TotalSeconds * 100)) / 100.0 + " c";
                    progressLabel.Text = "100%";
                    progressLabel.ForeColor = Color.Green;                   
                });
            }
            catch (OperationCanceledException ex)
            {
                Invoke((Action)delegate
                {                  
                    progressBar.Visible = false;
                    progressLabel.Text = "Calceled";
                    progressLabel.ForeColor = Color.Red;
                    cancelToken.Dispose();
                });
            }
            finally
            {
                parameters.Clear();
            }
                
        }

        private void ProcessFilesSingle()
        {         
            int filesQuantity = data.GetElementsByTagName("row").Count;
            int counter = 1;
            DateTime startTime = DateTime.Now;

            XsltArgumentList args = new XsltArgumentList();
            foreach (XmlNode node in data.GetElementsByTagName("row"))
            {
                if (int.Parse(node.Attributes[0].InnerText) < 2)
                    continue;

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
                progressLabel.Text = ((int)((counter * 100.0 / filesQuantity) * 100)) / 100.0 + "%";
                progressLabel.Update();
            }

            lTime.Visible = true;
            lTime.Text = ((int)((DateTime.Now - startTime).TotalSeconds * 100)) / 100.0 + " c";
            progressLabel.Text = "100%";
            progressLabel.ForeColor = Color.Green;
            parameters.Clear();
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

        private void Cancel_Click(object sender, EventArgs e)
        {
            if (!Cancel)
            {
                cancelToken.Cancel();
                Cancel = true;
            }
        }

        private void Start_Click(object sender, EventArgs e)
        {
            if (templatePath == null || xlsxTablePath == null)
                MessageBox.Show("Выберите файл", "Файл не выбран", MessageBoxButtons.OK);
            else if (!IsFileLocked(xlsxTablePath) && !IsFileLocked(templatePath))
            {
                cancelToken = new CancellationTokenSource();
                Cancel = false;
                StartJob();
            }
        }
    }
}
