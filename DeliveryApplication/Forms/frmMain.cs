using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using DeliveryApplication.Class;
using Pri.LongPath;

namespace DeliveryApplication.Forms
{
    public partial class frmMain : Form
    {
        private static clsVariable.ProjectSet project;
        public int rowIndex;

        static List<string> listEnum = new List<string>()
        {
            "Caselaw",
            "NonVirgo",
            "Virgo",
            "CaseRelated",
            "StateNet",
            "SecuritiesMosaic"
        };

        public frmMain()
        {
            InitializeComponent();
        }

        private void RunDelivery()
        {
            tsslblStatus.Text = "Cleaning up Input folder ...";
            clsFunction.CleanupInput();
            tsslblStatus.Text = "Delivery console app is running. Please wait ...";
            Program.ShowWindow(Program.handle, 5);
            if (clsDelivery.TransmittalInitialize())
            {
                try
                {
                    listEnum.ForEach(x =>
                    {
                        try
                        {
                            project = clsFunction.StringToEnum(x);

                            clsDelivery.Delivery(project);

                            if (!string.IsNullOrEmpty(clsVariable.JobnamesPath))
                            {
                                dataGridView1.Rows.Add(new string[]
                                {
                                    x,
                                    clsDelivery.RemoteSFtpPart,
                                    clsVariable.projectPath,
                                    clsVariable.nextPartPath,
                                    clsVariable.JobnamesPath
                                });
                            }
                        }
                        finally
                        {
                            clsDelivery.RemoteSFtpPart = string.Empty;
                            clsVariable.projectPath = string.Empty;
                            clsVariable.nextPartPath = string.Empty;
                            clsVariable.JobnamesPath = string.Empty;
                            clsVariable.listPartFolders = null;
                            clsVariable.listFileInfosCaselaw = null;
                            clsVariable.listFileInfosNonVirgo = null;
                            clsVariable.listFileInfosVirgo = null;
                            clsVariable.listFileInfosStateNet = null;
                            clsVariable.listDirectoryInfosNonVirgo = null;
                            clsVariable.listDirectoryInfosCaseRelated = null;
                            clsVariable.listDirectoryInfosStateNet = null;
                            clsVariable.listDirectoryInfosSecuritiesMosaic = null;
                            clsVariable.TransmittalFileInfos = null;
                        }
                    });
                }
                finally
                {
                    clsDelivery.SaveAsLNSchedule();
                    clsDelivery.LNScheduleBackup();
                    Program.ShowWindow(Program.handle, 0);
                    clsFunction.ReleaseMemory();

                    if (dataGridView1.Rows.Count > 0)
                    {
                        if (dataGridView1.Rows.Count > 1)
                        {
                            if (dataGridView1.Rows[0].Cells[0].Value != null)
                            {
                                var msg = MessageBox.Show("Upload all files ?", "Infomation", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                                if (msg == DialogResult.Yes)
                                {
                                    RunUpload();
                                }
                                else
                                {
                                    clsFunction.ReleaseMemory();
                                }
                            }
                        }
                    }
                    tsslblStatus.Text = "Done.";
                }
            }
            else
            {
                if (clsVariable.isAutoRun)
                {
                    clsConsole.WriteLine("\r\n LN Schedule in used by another process. Program will restart after 5 minutes.", ConsoleColor.Red);
                    System.Threading.Thread.Sleep(5000);
                    clsFunction.RestartApp(5);
                }
                else
                {
                    var msg = MessageBox.Show("LN Schedule in used by another process. Please close the excel files then press retry, if not please cancel.", "Warning", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning);
                    if (msg == DialogResult.Retry)
                    {
                        RunDelivery();
                    }
                    else
                    {
                        Program.ShowWindow(Program.handle, 0);
                    }
                }

            }

        }

        private void RunUpload()
        {
            if (dataGridView1.Rows.Count > 0)
            {
                Program.ShowWindow(Program.handle, 5);
                tsslblStatus.Text = "Upload console app is running. Please wait ...";
                try
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.Cells[0].Value != null)
                        {
                            clsSFTP.sFtpUploadFiles(clsFunction.StringToEnum(row.Cells[0].Value.ToString()), row.Cells[2].Value.ToString(), row.Cells[3].Value.ToString());
                        }
                    }
                }
                finally
                {
                    clsFunction.ReleaseMemory();
                    Program.ShowWindow(Program.handle, 0);
                    tsslblStatus.Text = "Done.";
                }
            }
        }

        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                dataGridView1.Rows[e.RowIndex].Selected = true;
                rowIndex = e.RowIndex;
                dataGridView1.CurrentCell = dataGridView1.Rows[e.RowIndex].Cells[1];
                contextMenuStrip1.Show(dataGridView1, e.Location);
                contextMenuStrip1.Show(Cursor.Position);
            }
        }

        private void uploadToLagunaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string tmpProjectPath = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                string tmpNextPath = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                Program.ShowWindow(Program.handle, 5);
                tsslblStatus.Text = string.Format("Uploading {0}", tmpNextPath);
                clsFunction.ClipboardSetText(dataGridView1.CurrentRow.Cells[1].Value.ToString());
                clsSFTP.sFtpUploadFiles(clsFunction.StringToEnum(dataGridView1.CurrentRow.Cells[0].Value.ToString()), tmpProjectPath, tmpNextPath);
            }
            finally
            {
                tsslblStatus.Text = "Done.";
                dataGridView1.Rows.RemoveAt(dataGridView1.CurrentCell.RowIndex);
                dataGridView1.ClearSelection();
                Program.ShowWindow(Program.handle, 0);
            }
        }

        private void deleteCurrentRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!dataGridView1.Rows[rowIndex].IsNewRow)
            {
                string Jobname = dataGridView1.Rows[rowIndex].Cells[4].Value.ToString();
                if (File.Exists(Jobname))
                {
                    File.Delete(Jobname);
                }
                dataGridView1.Rows.RemoveAt(rowIndex);
            }
        }

        private void btnConfig_Click(object sender, EventArgs e)
        {
            clsFunction.OpenFormSTA(FormConfig);
        }

        private static void FormConfig()
        {
            frmConfig form = new frmConfig();
            form.ShowDialog();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 1)
            {
                var msg = MessageBox.Show("Are you sure to delete grid data and jobnames ?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (msg == DialogResult.Yes)
                {
                    dataGridView1.Rows.Clear();
                    clsFunction.ClearJobnames();
                }
            }
            else
            {
                MessageBox.Show("Grid is empty ???", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnAbout_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Delivery Application v1.0 x E2322\r\n\r\n =====Thanks for using my casual program=====", "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void frmMain_Shown(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            RunDelivery();
        }

        private void btnRestart_Click(object sender, EventArgs e)
        {
            if (new DirectoryInfo(clsVariable.documentPath).GetFiles("*.xls").Where(p => Regex.IsMatch(p.Name, "(Jobnames|Copy of Jobnames)_(Caselaw|CaseRelated|NonVirgo|Virgo)_Part [A-Z0-9]+.xls", RegexOptions.IgnoreCase)).ToList().Count > 0)
            {
                var msg = MessageBox.Show("Are you sure to delete all old Jobname files ?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (msg == DialogResult.Yes)
                {
                    clsFunction.ClearJobnames();
                }
            }

            clsFunction.ReleaseMemory();
            RunDelivery();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex.Equals(1) && e.RowIndex != -1)
            {
                if (dataGridView1.CurrentCell != null && dataGridView1.CurrentCell.Value != null)
                {
                    clsFunction.ClipboardSetText(dataGridView1.CurrentCell.Value.ToString());
                }
            }
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            RunUpload();
        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (dataGridView1.Rows.Count > 1)
            {
                btnClear.Enabled = true;
            }
        }
    }
}
