using System;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace InvestNote
{
    public partial class NewAnalysis : Form
    {
        int hkid;
        bool isDirty = false;
        string currentFile = string.Empty;
        public NewAnalysis()
        {
            InitializeComponent();
        }

        private void SaveAndExit_Click(object sender, EventArgs e)
        {
            if (!isDirty)
            {
                System.Environment.Exit(0);
            }
            else if(MessageBox.Show("Are you sure?", "Save before exit", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                if (string.IsNullOrEmpty(tbTarget.Text))
                {
                    MessageBox.Show("Please input target name");
                    return;
                }
                Save_Click(sender, e);
                //this.Close();
                System.Environment.Exit(0);
            }
        }

        private void CancelAndExit_Click(object sender, EventArgs e)
        {
            if (!isDirty)
            {
                System.Environment.Exit(0);
            }
            else if(MessageBox.Show("Are you sure?","Exit without saving", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                this.Close();
            }
        }
        private void NewAnalysis_Load(object sender, EventArgs e)
        {
            hkid = HotKeyHelper.RegisterHotKey(Keys.OemSemicolon, KeyModifiers.Control);
            HotKeyHelper.HotKeyPressed += HotKeyHelper_HotKeyPressed;
        }

        private void HotKeyHelper_HotKeyPressed(object sender, HotKeyEventArgs e)
        {
            Clipboard.SetData("Text", DateTime.Now.ToString());
            richTextBox1.Paste();
        }

        private void NewAnalysis_FormClosed(object sender, FormClosedEventArgs e)
        {
            HotKeyHelper.UnregisterHotKey(hkid);
        }

        private void Save_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbTarget.Text))
            {
                //MessageBox.Show("Please input target name");
                tbTarget.Text = Interaction.InputBox("Name", "Please input name for the analysis", "Default");
                if (string.IsNullOrEmpty(tbTarget.Text))
                {
                    return;
                }
            }
            try
            {

                if (string.IsNullOrEmpty(currentFile))
                {
                    string path = tbTarget.Text + DateTime.Now.ToString("_yyyyMMddHHmmss") + ".rtf";
                    currentFile = GetOutputFilePath(path);
                }
                richTextBox1.SaveFile(currentFile);
                isDirty = false;
                this.Text = this.Text.Trim('*');
                toolStripStatusLabel1.Text = "File saved to "+currentFile;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private string GetOutputFilePath(string str) {
            string dir = ConfigurationManager.AppSettings["DestDirectory"];
            try
            {
                return Path.Combine(dir, str);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return str;
            }
        }

        private void PeriodButton_Click(object sender, EventArgs e)
        {
            string str = string.Format("----{0}-{0}-{0}-{0}-{0}-{0}-{0}-{0}-{0}----\n",sender.ToString());
            richTextBox1.AppendText(str);
        }

        private void Config_Click(object sender, EventArgs e)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
            FolderBrowserDialog dilog = new FolderBrowserDialog();
            dilog.Description = "Select a folder";
            try
            {
                if (dilog.ShowDialog() == DialogResult.OK || dilog.ShowDialog() == DialogResult.Yes)
                {
                    config.AppSettings.Settings["DestDirectory"].Value= dilog.SelectedPath;
                    config.Save(ConfigurationSaveMode.Modified);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void NewAnalysis_Click(object sender, EventArgs e)
        {
            if (isDirty)
            {
                Save_Click(sender, e);
                NewAnalysis_Click(sender, e);
                currentFile = string.Empty;
            }
            else
            {
                richTextBox1.Clear();
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            isDirty = true;
            if (!this.Text.EndsWith("*"))
            {
                this.Text += "*";
            }
        }

        private void Open_Click(object sender, EventArgs e)
        {
            if (isDirty)
            {
                Save_Click(sender, e);
                Open_Click(sender, e);
            }
            else
            {
                OpenFileDialog ofd = new OpenFileDialog
                {
                    Filter = @"Rtf File|*.rtf",
                    Title = @"Please choose an rtf file",
                    RestoreDirectory = true
                };

                ofd.Multiselect = false;

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    richTextBox1.LoadFile(ofd.FileName);
                    currentFile = ofd.FileName;
                    isDirty = false;
                    this.Text = this.Text.Trim('*');
                }
            }
        }

        private void ClearAll_Click(object sender, EventArgs e)
        {
            if(MessageBox.Show("Empty the content without saving?","Are you sure", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                richTextBox1.Clear();
            }
        }

        private void tbTarget_TextChanged(object sender, EventArgs e)
        {
            string path = tbTarget.Text + DateTime.Now.ToString("_yyyyMMddHHmmss") + ".rtf";
            currentFile = GetOutputFilePath(path);
        }
    }

    //enum Period {
    //    M5,
    //    M15,
    //    M30,
    //    H1,
    //    H2,
    //    H3,
    //    H4,
    //    H8,
    //    D1,
    //    W1,
    //    MN,
    //    SSN
    //}
}
