namespace KaitechColumnsReportAddin
{
    partial class ColumnReportForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxReportLocation = new System.Windows.Forms.TextBox();
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.buttonExport = new System.Windows.Forms.Button();
            this.buttonClose = new System.Windows.Forms.Button();
            this.pictureBoxLinkedIn = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLinkedIn)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 84);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(105, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Report Location:";
            // 
            // textBoxReportLocation
            // 
            this.textBoxReportLocation.Location = new System.Drawing.Point(165, 78);
            this.textBoxReportLocation.Name = "textBoxReportLocation";
            this.textBoxReportLocation.ReadOnly = true;
            this.textBoxReportLocation.Size = new System.Drawing.Size(441, 22);
            this.textBoxReportLocation.TabIndex = 1;
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.Location = new System.Drawing.Point(647, 77);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(85, 23);
            this.buttonBrowse.TabIndex = 2;
            this.buttonBrowse.Text = "Browse";
            this.buttonBrowse.UseVisualStyleBackColor = true;
            this.buttonBrowse.Click += new System.EventHandler(this.buttonBrowse_Click);
            // 
            // buttonExport
            // 
            this.buttonExport.Location = new System.Drawing.Point(219, 177);
            this.buttonExport.Name = "buttonExport";
            this.buttonExport.Size = new System.Drawing.Size(96, 23);
            this.buttonExport.TabIndex = 3;
            this.buttonExport.Text = "Export";
            this.buttonExport.UseVisualStyleBackColor = true;
            this.buttonExport.Click += new System.EventHandler(this.buttonExport_Click);
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(463, 177);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(96, 23);
            this.buttonClose.TabIndex = 4;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // pictureBoxLinkedIn
            // 
            this.pictureBoxLinkedIn.Image = global::KaitechColumnsReportAddin.Properties.Resources.linkedin_icon;
            this.pictureBoxLinkedIn.Location = new System.Drawing.Point(683, 150);
            this.pictureBoxLinkedIn.Name = "pictureBoxLinkedIn";
            this.pictureBoxLinkedIn.Size = new System.Drawing.Size(49, 50);
            this.pictureBoxLinkedIn.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBoxLinkedIn.TabIndex = 5;
            this.pictureBoxLinkedIn.TabStop = false;
            this.pictureBoxLinkedIn.Click += new System.EventHandler(this.pictureBoxLinkedIn_Click);
            // 
            // ColumnReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(778, 232);
            this.Controls.Add(this.pictureBoxLinkedIn);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.buttonExport);
            this.Controls.Add(this.buttonBrowse);
            this.Controls.Add(this.textBoxReportLocation);
            this.Controls.Add(this.label1);
            this.Name = "ColumnReportForm";
            this.Text = "Columns Report";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLinkedIn)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxReportLocation;
        private System.Windows.Forms.Button buttonBrowse;
        private System.Windows.Forms.Button buttonExport;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.PictureBox pictureBoxLinkedIn;
    }
}