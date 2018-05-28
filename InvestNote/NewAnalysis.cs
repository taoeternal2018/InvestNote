using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InvestNote
{
    public partial class NewAnalysis : Form
    {
        int hkid;
        public NewAnalysis()
        {
            InitializeComponent();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void NewAnalysis_Load(object sender, EventArgs e)
        {
            hkid = HotKeyHelper.RegisterHotKey(Keys.V, KeyModifiers.Control);
            HotKeyHelper.HotKeyPressed += HotKeyHelper_HotKeyPressed;
        }

        private void HotKeyHelper_HotKeyPressed(object sender, HotKeyEventArgs e)
        {
            pictureBox1.Image = Clipboard.GetImage();
        }

        private void NewAnalysis_FormClosed(object sender, FormClosedEventArgs e)
        {
            HotKeyHelper.UnregisterHotKey(hkid);
        }
    }
}
