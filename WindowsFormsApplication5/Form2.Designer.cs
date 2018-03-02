namespace WindowsFormsApplication5
{
    partial class Form2
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
            this.newFile = new System.Windows.Forms.Button();
            this.start = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // newFile
            // 
            this.newFile.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.newFile.Font = new System.Drawing.Font("Microsoft JhengHei", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.newFile.ForeColor = System.Drawing.Color.White;
            this.newFile.Location = new System.Drawing.Point(129, 12);
            this.newFile.Name = "newFile";
            this.newFile.Size = new System.Drawing.Size(120, 50);
            this.newFile.TabIndex = 1;
            this.newFile.Text = "Creat";
            this.newFile.UseVisualStyleBackColor = false;
            this.newFile.Click += new System.EventHandler(this.newFile_Click);
            // 
            // start
            // 
            
            this.start.BackColor = System.Drawing.Color.DarkGreen;
            this.start.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.start.Font = new System.Drawing.Font("Microsoft JhengHei", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.start.ForeColor = System.Drawing.Color.White;
            this.start.Location = new System.Drawing.Point(3, 12);
            this.start.Name = "start";
            this.start.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.start.Size = new System.Drawing.Size(120, 50);
            this.start.TabIndex = 0;
            this.start.Text = "Start";
            this.start.UseVisualStyleBackColor = false;
            this.start.Click += new System.EventHandler(this.start_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.BackColor = System.Drawing.Color.LightGray;
            this.ClientSize = new System.Drawing.Size(1020, 741);
            this.Controls.Add(this.newFile);
            this.Controls.Add(this.start);
            this.Name = "Form2";
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button start;
        private System.Windows.Forms.Button newFile;
    }
}