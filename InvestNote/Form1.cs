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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void EditAnalysis_Click(object sender, EventArgs e)
        {
            
        }

        private void NewAnalysis_Click(object sender, EventArgs e)
        {
            NewAnalysis newForm = new NewAnalysis();
            newForm.Show();
        }

        private void NewDirectory_Click(object sender, EventArgs e)
        {

        }

        private void Remove_Click(object sender, EventArgs e)
        {

        }
    }
}
