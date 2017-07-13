namespace Module_Report
{
    partial class Report
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
            this.lstMapping = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnLoad = new System.Windows.Forms.Button();
            this.lstBrowser = new System.Windows.Forms.ListView();
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lstMapping
            // 
            this.lstMapping.AllowDrop = true;
            this.lstMapping.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2});
            this.lstMapping.FullRowSelect = true;
            this.lstMapping.GridLines = true;
            this.lstMapping.Location = new System.Drawing.Point(6, 62);
            this.lstMapping.Name = "lstMapping";
            this.lstMapping.Size = new System.Drawing.Size(333, 262);
            this.lstMapping.TabIndex = 3;
            this.lstMapping.UseCompatibleStateImageBehavior = false;
            this.lstMapping.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Template Bookmarks";
            this.columnHeader1.Width = 136;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Mapping column";
            this.columnHeader2.Width = 184;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Location = new System.Drawing.Point(12, 1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(360, 589);
            this.panel1.TabIndex = 4;
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(253, 19);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(69, 23);
            this.btnGenerate.TabIndex = 66;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lstBrowser);
            this.groupBox1.Controls.Add(this.btnGenerate);
            this.groupBox1.Controls.Add(this.btnLoad);
            this.groupBox1.Controls.Add(this.lstMapping);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(9, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(345, 556);
            this.groupBox1.TabIndex = 65;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Load Directory";
            // 
            // btnLoad
            // 
            this.btnLoad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLoad.Location = new System.Drawing.Point(20, 19);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(63, 23);
            this.btnLoad.TabIndex = 72;
            this.btnLoad.Text = "&Load";
            this.btnLoad.UseVisualStyleBackColor = true;
            // 
            // lstBrowser
            // 
            this.lstBrowser.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader3});
            this.lstBrowser.FullRowSelect = true;
            this.lstBrowser.GridLines = true;
            this.lstBrowser.Location = new System.Drawing.Point(6, 330);
            this.lstBrowser.Name = "lstBrowser";
            this.lstBrowser.Size = new System.Drawing.Size(333, 220);
            this.lstBrowser.TabIndex = 4;
            this.lstBrowser.UseCompatibleStateImageBehavior = false;
            this.lstBrowser.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Available Figures";
            this.columnHeader3.Width = 320;
            // 
            // Report
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(392, 602);
            this.Controls.Add(this.panel1);
            this.Name = "Report";
            this.Text = "Report";
            this.panel1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListView lstMapping;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        public System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.ListView lstBrowser;
        private System.Windows.Forms.ColumnHeader columnHeader3;
    }
}