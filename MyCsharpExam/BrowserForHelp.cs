using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace MyCsharpExam
{
    public partial class BrowserForHelp : Form
    {
        private string uri;
        public BrowserForHelp(string uri)
        {
            InitializeComponent();
            this.uri = uri;
        }

        private void BrowserForHelp_Load(object sender, EventArgs e) => webBrowser1.Navigate(Path.GetFullPath(uri));

        private void BrowserForHelp_FormClosing(object sender, FormClosingEventArgs e)=> ((Form1)Owner).helpWindow = null; 
    }
}
