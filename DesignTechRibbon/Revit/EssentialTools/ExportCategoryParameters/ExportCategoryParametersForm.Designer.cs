namespace DesignTechRibbonPaid.Revit.EssentialTools.ExportCategoryParametersForm
{
    partial class ExportCategoryParametersForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExportCategoryParametersForm));
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.cancelButton = new System.Windows.Forms.Button();
            this.ExportToExcel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.CloseButton = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.designtechLogo = new System.Windows.Forms.PictureBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.designtechLogo)).BeginInit();
            this.SuspendLayout();
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(9, 322);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(272, 43);
            this.progressBar1.TabIndex = 9;
            // 
            // cancelButton
            // 
            this.cancelButton.Location = new System.Drawing.Point(287, 322);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(138, 43);
            this.cancelButton.TabIndex = 10;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // ExportToExcel
            // 
            this.ExportToExcel.Location = new System.Drawing.Point(287, 45);
            this.ExportToExcel.Name = "ExportToExcel";
            this.ExportToExcel.Size = new System.Drawing.Size(141, 43);
            this.ExportToExcel.TabIndex = 11;
            this.ExportToExcel.Text = "Export to Excel";
            this.ExportToExcel.UseVisualStyleBackColor = true;
            this.ExportToExcel.Click += new System.EventHandler(this.ExportToExcel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(274, 17);
            this.label1.TabIndex = 12;
            this.label1.Text = "Please select a Category in the box below:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.CloseButton);
            this.groupBox1.Controls.Add(this.listBox1);
            this.groupBox1.Controls.Add(this.designtechLogo);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.ExportToExcel);
            this.groupBox1.Controls.Add(this.progressBar1);
            this.groupBox1.Controls.Add(this.cancelButton);
            this.groupBox1.Location = new System.Drawing.Point(12, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(436, 452);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            // 
            // CloseButton
            // 
            this.CloseButton.Location = new System.Drawing.Point(287, 403);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new System.Drawing.Size(138, 43);
            this.CloseButton.TabIndex = 15;
            this.CloseButton.Text = "Close";
            this.CloseButton.UseVisualStyleBackColor = true;
            this.CloseButton.Click += new System.EventHandler(this.CloseButton_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 16;
            this.listBox1.Location = new System.Drawing.Point(7, 45);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(274, 260);
            this.listBox1.TabIndex = 14;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            // 
            // designtechLogo
            // 
            this.designtechLogo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.designtechLogo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.designtechLogo.Image = ((System.Drawing.Image)(resources.GetObject("designtechLogo.Image")));
            this.designtechLogo.Location = new System.Drawing.Point(6, 372);
            this.designtechLogo.Margin = new System.Windows.Forms.Padding(4);
            this.designtechLogo.Name = "designtechLogo";
            this.designtechLogo.Size = new System.Drawing.Size(275, 73);
            this.designtechLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.designtechLogo.TabIndex = 13;
            this.designtechLogo.TabStop = false;
            // 
            // ExportCategoryParametersForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(456, 465);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ExportCategoryParametersForm";
            this.Text = "Export Category Parameters";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ExportCategoryParametersForm_FormClosing);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.designtechLogo)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Button ExportToExcel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.PictureBox designtechLogo;
        private System.Windows.Forms.Button CloseButton;
    }
}