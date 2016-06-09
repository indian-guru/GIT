namespace ExcelImageImporter
{
    partial class Importer
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
            this.buttonExtract = new System.Windows.Forms.Button();
            this.textBoxFilePath = new System.Windows.Forms.TextBox();
            this.textBoxImagePath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.buttonExit = new System.Windows.Forms.Button();
            this.buttonFilePath = new System.Windows.Forms.Button();
            this.buttonImagePath = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonExtract
            // 
            this.buttonExtract.Location = new System.Drawing.Point(50, 93);
            this.buttonExtract.Name = "buttonExtract";
            this.buttonExtract.Size = new System.Drawing.Size(88, 30);
            this.buttonExtract.TabIndex = 1;
            this.buttonExtract.Text = "Extract";
            this.buttonExtract.UseVisualStyleBackColor = true;
            this.buttonExtract.Click += new System.EventHandler(this.buttonExtract_Click);
            // 
            // textBoxFilePath
            // 
            this.textBoxFilePath.Location = new System.Drawing.Point(190, 27);
            this.textBoxFilePath.Name = "textBoxFilePath";
            this.textBoxFilePath.Size = new System.Drawing.Size(331, 20);
            this.textBoxFilePath.TabIndex = 2;
            // 
            // textBoxImagePath
            // 
            this.textBoxImagePath.Location = new System.Drawing.Point(190, 53);
            this.textBoxImagePath.Name = "textBoxImagePath";
            this.textBoxImagePath.Size = new System.Drawing.Size(331, 20);
            this.textBoxImagePath.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Excel File Path";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(121, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Folder to Extract Images";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(321, 93);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(351, 23);
            this.progressBar.TabIndex = 4;
            // 
            // buttonExit
            // 
            this.buttonExit.Location = new System.Drawing.Point(144, 93);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Size = new System.Drawing.Size(81, 30);
            this.buttonExit.TabIndex = 1;
            this.buttonExit.Text = "Exit";
            this.buttonExit.UseVisualStyleBackColor = true;
            this.buttonExit.Click += new System.EventHandler(this.buttonExit_Click);
            // 
            // buttonFilePath
            // 
            this.buttonFilePath.Location = new System.Drawing.Point(527, 27);
            this.buttonFilePath.Name = "buttonFilePath";
            this.buttonFilePath.Size = new System.Drawing.Size(33, 24);
            this.buttonFilePath.TabIndex = 1;
            this.buttonFilePath.Text = "...";
            this.buttonFilePath.UseVisualStyleBackColor = true;
            this.buttonFilePath.Click += new System.EventHandler(this.buttonFilePath_Click);
            // 
            // buttonImagePath
            // 
            this.buttonImagePath.Location = new System.Drawing.Point(527, 53);
            this.buttonImagePath.Name = "buttonImagePath";
            this.buttonImagePath.Size = new System.Drawing.Size(33, 24);
            this.buttonImagePath.TabIndex = 1;
            this.buttonImagePath.Text = "...";
            this.buttonImagePath.UseVisualStyleBackColor = true;
            this.buttonImagePath.Click += new System.EventHandler(this.buttonImagePath_Click);
            // 
            // Importer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(681, 155);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxImagePath);
            this.Controls.Add(this.textBoxFilePath);
            this.Controls.Add(this.buttonImagePath);
            this.Controls.Add(this.buttonFilePath);
            this.Controls.Add(this.buttonExit);
            this.Controls.Add(this.buttonExtract);
            this.Name = "Importer";
            this.Text = "Export Images From Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button buttonExtract;
        private System.Windows.Forms.TextBox textBoxFilePath;
        private System.Windows.Forms.TextBox textBoxImagePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button buttonExit;
        private System.Windows.Forms.Button buttonFilePath;
        private System.Windows.Forms.Button buttonImagePath;
    }
}

