using Pri.LongPath;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DeliveryApplication.Forms
{
    public partial class frmConfig : Form
    {
        public frmConfig()
        {
            InitializeComponent();
        }

        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(tbReleasePath.Text) && !string.IsNullOrEmpty(tbTransmittalPath.Text) && !string.IsNullOrEmpty(tbInterval.Text) && Directory.Exists(tbReleasePath.Text) && Directory.Exists(tbTransmittalPath.Text))
            {
                Properties.Settings.Default.Save();
                MessageBox.Show("Config has been saved !!!. Program will be restart for accept this config.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Application.Restart();
            }
            else
            {
                tbReleasePath.Text = string.Empty;
                tbTransmittalPath.Text = string.Empty;
                tbInterval.Text = string.Empty;
            }
        }

        private void btnBrowse1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Select release path";
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                tbReleasePath.Text = fbd.SelectedPath;
            }
        }

        private void btnBrowse2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Select transmittal path";
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                tbTransmittalPath.Text = fbd.SelectedPath;
                Properties.Settings.Default.TrasmittalPath = fbd.SelectedPath;
            }
        }

        private void frmConfig_Shown(object sender, EventArgs e)
        {
            tbReleasePath.Text = Properties.Settings.Default.ReleasePath;
            tbTransmittalPath.Text = Properties.Settings.Default.TrasmittalPath;
            tbInterval.Text = Properties.Settings.Default.Interval.ToString();
        }

        private void tbInterval_TextChanged(object sender, EventArgs e)
        {
            int interval;
            if (!string.IsNullOrEmpty(tbInterval.Text))
            {
                if (Int32.TryParse(tbInterval.Text, out interval))
                {
                    Properties.Settings.Default.Interval = interval;
                }
                else
                {
                    MessageBox.Show("Please enter the correct number", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void tbReleasePath_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(tbReleasePath.Text) && Directory.Exists(tbReleasePath.Text))
            {
                Properties.Settings.Default.ReleasePath = tbReleasePath.Text;
            }
            else
            {
                tbReleasePath.Text = string.Empty;
            }
        }

        private void tbTransmittalPath_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(tbTransmittalPath.Text) && Directory.Exists(tbTransmittalPath.Text))
            {
                Properties.Settings.Default.TrasmittalPath = tbTransmittalPath.Text;
            }
            else
            {
                tbTransmittalPath.Text = string.Empty;
            }
        }
    }
}
