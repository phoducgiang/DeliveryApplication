
namespace DeliveryApplication.Forms
{
    partial class frmConfig
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmConfig));
            this.tbReleasePath = new System.Windows.Forms.TextBox();
            this.btnBrowse1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnBrowse2 = new System.Windows.Forms.Button();
            this.tbTransmittalPath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnSaveConfig = new System.Windows.Forms.Button();
            this.tbInterval = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // tbReleasePath
            // 
            this.tbReleasePath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbReleasePath.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbReleasePath.Location = new System.Drawing.Point(38, 50);
            this.tbReleasePath.Name = "tbReleasePath";
            this.tbReleasePath.Size = new System.Drawing.Size(408, 25);
            this.tbReleasePath.TabIndex = 2;
            this.tbReleasePath.TextChanged += new System.EventHandler(this.tbReleasePath_TextChanged);
            // 
            // btnBrowse1
            // 
            this.btnBrowse1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnBrowse1.BackgroundImage")));
            this.btnBrowse1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnBrowse1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBrowse1.Location = new System.Drawing.Point(452, 50);
            this.btnBrowse1.Name = "btnBrowse1";
            this.btnBrowse1.Size = new System.Drawing.Size(27, 25);
            this.btnBrowse1.TabIndex = 1;
            this.btnBrowse1.UseVisualStyleBackColor = true;
            this.btnBrowse1.Click += new System.EventHandler(this.btnBrowse1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(38, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 17);
            this.label1.TabIndex = 6;
            this.label1.Text = "Release Path";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(38, 83);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(108, 17);
            this.label2.TabIndex = 7;
            this.label2.Text = "Transmittal Path";
            // 
            // btnBrowse2
            // 
            this.btnBrowse2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnBrowse2.BackgroundImage")));
            this.btnBrowse2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnBrowse2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBrowse2.Location = new System.Drawing.Point(452, 111);
            this.btnBrowse2.Name = "btnBrowse2";
            this.btnBrowse2.Size = new System.Drawing.Size(27, 25);
            this.btnBrowse2.TabIndex = 3;
            this.btnBrowse2.UseVisualStyleBackColor = true;
            this.btnBrowse2.Click += new System.EventHandler(this.btnBrowse2_Click);
            // 
            // tbTransmittalPath
            // 
            this.tbTransmittalPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbTransmittalPath.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbTransmittalPath.Location = new System.Drawing.Point(38, 111);
            this.tbTransmittalPath.Name = "tbTransmittalPath";
            this.tbTransmittalPath.Size = new System.Drawing.Size(408, 25);
            this.tbTransmittalPath.TabIndex = 4;
            this.tbTransmittalPath.TextChanged += new System.EventHandler(this.tbTransmittalPath_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(38, 143);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 17);
            this.label3.TabIndex = 8;
            this.label3.Text = "Interval";
            // 
            // btnSaveConfig
            // 
            this.btnSaveConfig.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnSaveConfig.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSaveConfig.Font = new System.Drawing.Font("Tahoma", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveConfig.Location = new System.Drawing.Point(315, 143);
            this.btnSaveConfig.Name = "btnSaveConfig";
            this.btnSaveConfig.Size = new System.Drawing.Size(164, 53);
            this.btnSaveConfig.TabIndex = 0;
            this.btnSaveConfig.Text = "Save";
            this.btnSaveConfig.UseVisualStyleBackColor = true;
            this.btnSaveConfig.Click += new System.EventHandler(this.btnSaveConfig_Click);
            // 
            // tbInterval
            // 
            this.tbInterval.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbInterval.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbInterval.Location = new System.Drawing.Point(38, 171);
            this.tbInterval.Name = "tbInterval";
            this.tbInterval.Size = new System.Drawing.Size(166, 25);
            this.tbInterval.TabIndex = 5;
            this.tbInterval.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.tbInterval.TextChanged += new System.EventHandler(this.tbInterval_TextChanged);
            // 
            // frmConfig
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(515, 226);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnSaveConfig);
            this.Controls.Add(this.tbInterval);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnBrowse2);
            this.Controls.Add(this.tbTransmittalPath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnBrowse1);
            this.Controls.Add(this.tbReleasePath);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmConfig";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Config";
            this.Shown += new System.EventHandler(this.frmConfig_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbReleasePath;
        private System.Windows.Forms.Button btnBrowse1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnBrowse2;
        private System.Windows.Forms.TextBox tbTransmittalPath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnSaveConfig;
        private System.Windows.Forms.TextBox tbInterval;
    }
}