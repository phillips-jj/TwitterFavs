namespace TwitterFavs
{
    partial class frmMain
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
            this.btnGo = new System.Windows.Forms.Button();
            this.strip = new System.Windows.Forms.StatusStrip();
            this.stripLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.btnFile = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtFile = new System.Windows.Forms.TextBox();
            this.strip.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnGo
            // 
            this.btnGo.Location = new System.Drawing.Point(152, 110);
            this.btnGo.Name = "btnGo";
            this.btnGo.Size = new System.Drawing.Size(75, 23);
            this.btnGo.TabIndex = 0;
            this.btnGo.Text = "Go";
            this.btnGo.UseVisualStyleBackColor = true;
            this.btnGo.Click += new System.EventHandler(this.btnGo_Click);
            // 
            // strip
            // 
            this.strip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.stripLabel});
            this.strip.Location = new System.Drawing.Point(0, 168);
            this.strip.Name = "strip";
            this.strip.Size = new System.Drawing.Size(406, 22);
            this.strip.TabIndex = 2;
            this.strip.Text = "statusStrip1";
            // 
            // stripLabel
            // 
            this.stripLabel.Name = "stripLabel";
            this.stripLabel.Size = new System.Drawing.Size(118, 17);
            this.stripLabel.Text = "toolStripStatusLabel1";
            // 
            // btnFile
            // 
            this.btnFile.Location = new System.Drawing.Point(373, 59);
            this.btnFile.Name = "btnFile";
            this.btnFile.Size = new System.Drawing.Size(28, 23);
            this.btnFile.TabIndex = 8;
            this.btnFile.Text = "...";
            this.btnFile.UseVisualStyleBackColor = true;
            this.btnFile.Click += new System.EventHandler(this.btnFile_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(26, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "File:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtFile
            // 
            this.txtFile.Location = new System.Drawing.Point(44, 62);
            this.txtFile.Name = "txtFile";
            this.txtFile.Size = new System.Drawing.Size(323, 20);
            this.txtFile.TabIndex = 6;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(406, 190);
            this.Controls.Add(this.btnFile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtFile);
            this.Controls.Add(this.strip);
            this.Controls.Add(this.btnGo);
            this.Name = "frmMain";
            this.Text = "TwitterFavs2Excel";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.strip.ResumeLayout(false);
            this.strip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGo;
        private System.Windows.Forms.StatusStrip strip;
        private System.Windows.Forms.ToolStripStatusLabel stripLabel;
        private System.Windows.Forms.Button btnFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtFile;
    }
}

