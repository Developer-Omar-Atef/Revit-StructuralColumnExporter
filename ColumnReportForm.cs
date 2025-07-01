using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace KaitechColumnsReportAddin
{
    public partial class ColumnReportForm : Form
    {
        public string SelectedPath { get; private set; }
        public ColumnReportForm()
        {
            InitializeComponent();
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select a folder to save the report";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    textBoxReportLocation.Text = dialog.SelectedPath;
                    SelectedPath = dialog.SelectedPath;
                }
            }
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBoxReportLocation.Text) || !Directory.Exists(textBoxReportLocation.Text))
            {
                MessageBox.Show("Please select a valid report location.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SelectedPath = textBoxReportLocation.Text;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void pictureBoxLinkedIn_Click(object sender, EventArgs e)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "https://www.linkedin.com/in/omar-atef-60217b263",
                    UseShellExecute = true // This is key for URLs and registered file types
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not open LinkedIn profile: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
