namespace SCBFILEUpdate_
{
    partial class Form1
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
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.lblInsert = new System.Windows.Forms.Label();
            this.btnExport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(100, 70);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(166, 52);
            this.button1.TabIndex = 0;
            this.button1.Text = "Select Folder to Insert Data";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lblInsert
            // 
            this.lblInsert.BackColor = System.Drawing.Color.Transparent;
            this.lblInsert.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lblInsert.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInsert.ForeColor = System.Drawing.Color.MediumBlue;
            this.lblInsert.Location = new System.Drawing.Point(0, 137);
            this.lblInsert.Name = "lblInsert";
            this.lblInsert.Size = new System.Drawing.Size(381, 39);
            this.lblInsert.TabIndex = 1;
            this.lblInsert.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(100, 12);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(166, 52);
            this.btnExport.TabIndex = 2;
            this.btnExport.Text = "Export Bank Data to CSV file";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.BackgroundImage = global::SCBFILEUpdate_.Properties.Resources.snb;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(381, 176);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.lblInsert);
            this.Controls.Add(this.button1);
            this.DoubleBuffered = true;
            this.Name = "Form1";
            this.Text = "StandardChartered Bank_UpdateFiles";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label lblInsert;
        private System.Windows.Forms.Button btnExport;
    }
}

