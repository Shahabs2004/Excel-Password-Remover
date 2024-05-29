namespace ExcelPasswordRemover
{
    partial class MainForm
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
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.LogBox = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Enabled = false;
            this.btnOpenFile.Location = new System.Drawing.Point(86, 12);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(120, 23);
            this.btnOpenFile.TabIndex = 0;
            this.btnOpenFile.Text = "Open Protected File";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Visible = false;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // LogBox
            // 
            this.LogBox.FormattingEnabled = true;
            this.LogBox.Location = new System.Drawing.Point(6, 42);
            this.LogBox.Name = "LogBox";
            this.LogBox.Size = new System.Drawing.Size(282, 186);
            this.LogBox.TabIndex = 1;
            this.LogBox.SelectedIndexChanged += new System.EventHandler(this.LogBox_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 231);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(131, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Code by: Shahab Sadeghi";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(293, 250);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.LogBox);
            this.Controls.Add(this.btnOpenFile);
            this.Name = "MainForm";
            this.Text = "Excel Password Remover";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.ListBox LogBox;
        private System.Windows.Forms.Label label1;
    }
}

