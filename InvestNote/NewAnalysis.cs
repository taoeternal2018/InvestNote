using System;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Spire.Doc;
using System.Text;
using Spire.Doc.Documents;

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
            else if(MessageBox.Show("Save before exit?","Attention",  MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                tbTarget.Text = Interaction.InputBox("Please input target name", "Enter a name", "Default");
                if (string.IsNullOrEmpty(tbTarget.Text))
                {
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
            else if(MessageBox.Show("Exit without saving?","Attention", MessageBoxButtons.OKCancel) == DialogResult.OK)
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
                tbTarget.Text = Interaction.InputBox("Please input target name", "Enter a name", "Default");
                if (string.IsNullOrEmpty(tbTarget.Text))
                {
                    return;
                }
            }
            try
            {

                if (string.IsNullOrEmpty(currentFile))
                {
                    string path = tbTarget.Text + DateTime.Now.ToString("_yyyyMMddHHmmss") + ".doc";
                    currentFile = GetOutputFilePath(path);
                }
                var stream = new MemoryStream(Encoding.Unicode.GetBytes(richTextBox1.Rtf));
                Document doc = new Document();

                doc.LoadRtf(stream, Encoding.Unicode);
                foreach(Section v in doc.Sections)
                {
                    v.PageSetup.PageSize = PageSize.A3;
                    v.PageSetup.Orientation = PageOrientation.Landscape;
                }
                doc.SaveToFile(currentFile, FileFormat.Doc);
                isDirty = false;
                this.Text = this.Text.Trim('*');
                toolStripStatusLabel1.Text = "File saved to "+currentFile;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
                    Filter = @"Rtf File|*.doc",
                    Title = @"Please choose an doc file",
                    RestoreDirectory = true
                };

                ofd.Multiselect = false;
                ofd.InitialDirectory = ConfigurationManager.AppSettings["DestDirectory"].ToString();

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    tbTarget.Text = Path.GetFileName(ofd.FileName).Replace(".doc","");
                    Document doc = new Document();
                    doc.LoadFromFile(ofd.FileName, FileFormat.Doc);
                    string tmp = string.Format("{0}.rtf", Guid.NewGuid().ToString());
                    doc.SaveToFile(tmp, FileFormat.Rtf);
                    richTextBox1.LoadFile(tmp);
                    File.Delete(tmp);

                    currentFile = ofd.FileName;
                    isDirty = false;
                    this.Text = this.Text.Trim('*');
                }
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
            richTextBox1.SelectedText = string.Format("------Analysis for period {0}------\n", sender.ToString());
        }

        private void Config_Click(object sender, EventArgs e)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (!string.IsNullOrEmpty(config.AppSettings.Settings["DestDirectory"].Value)) {
                dialog.SelectedPath = config.AppSettings.Settings["DestDirectory"].Value;
            }
            dialog.Description = "Select a folder";
            try
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    config.AppSettings.Settings["DestDirectory"].Value= dialog.SelectedPath;
                    config.Save(ConfigurationSaveMode.Modified);
                    if (!string.IsNullOrEmpty(currentFile))
                    {
                        currentFile = Path.Combine(dialog.SelectedPath, Path.GetFileName(currentFile));
                    }
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
                tbTarget.Text = "";
                isDirty = false;
                this.Text = this.Text.Trim('*');
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

        
        private void ClearAll_Click(object sender, EventArgs e)
        {
            if (isDirty) {
                return;
            }
            else
            {
                if (MessageBox.Show("Empty the content without saving?", "Attention", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    richTextBox1.Clear();
                }
            }
        }

        private void tbTarget_TextChanged(object sender, EventArgs e)
        {
            string path = tbTarget.Text + DateTime.Now.ToString("_yyyyMMddHHmmss") + ".doc";
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
