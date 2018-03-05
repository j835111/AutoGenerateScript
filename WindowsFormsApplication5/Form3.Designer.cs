namespace WindowsFormsApplication5
{
    partial class Form3
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form3));
            this.button1 = new System.Windows.Forms.Button();
            this.chap_choise = new System.Windows.Forms.ComboBox();
            this.button2 = new System.Windows.Forms.Button();
            this.TestCatagory = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.White;
            this.button1.Font = new System.Drawing.Font("標楷體", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button1.ForeColor = System.Drawing.Color.Black;
            this.button1.Location = new System.Drawing.Point(174, 188);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(259, 50);
            this.button1.TabIndex = 0;
            this.button1.Text = "開啟excel資料表";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // chap_choise
            // 
            this.chap_choise.BackColor = System.Drawing.Color.White;
            this.chap_choise.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.chap_choise.FormattingEnabled = true;
            this.chap_choise.Location = new System.Drawing.Point(202, 278);
            this.chap_choise.Name = "chap_choise";
            this.chap_choise.Size = new System.Drawing.Size(208, 28);
            this.chap_choise.TabIndex = 2;
            this.chap_choise.SelectedIndexChanged += new System.EventHandler(this.Chap_choise_SelectedIndexChanged);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.White;
            this.button2.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button2.Location = new System.Drawing.Point(202, 439);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(208, 35);
            this.button2.TabIndex = 3;
            this.button2.Text = "下一步";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // TestCatagory
            // 
            this.TestCatagory.BackColor = System.Drawing.Color.White;
            this.TestCatagory.Font = new System.Drawing.Font("標楷體", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.TestCatagory.FormattingEnabled = true;
            this.TestCatagory.Items.AddRange(new object[] {
            "選擇測試種類"});
            this.TestCatagory.Location = new System.Drawing.Point(202, 354);
            this.TestCatagory.Name = "TestCatagory";
            this.TestCatagory.Size = new System.Drawing.Size(208, 26);
            this.TestCatagory.TabIndex = 6;
            this.TestCatagory.Text = "選擇測試種類";
            this.TestCatagory.Visible = false;
            this.TestCatagory.SelectedIndexChanged += new System.EventHandler(this.TestCatagory_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.label1.Font = new System.Drawing.Font("Pristina", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(606, 120);
            this.label1.TabIndex = 7;
            this.label1.Text = "TestCode Generator";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.ClientSize = new System.Drawing.Size(606, 570);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TestCatagory);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.chap_choise);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form3";
            this.Text = "程式碼產生器";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form3_FormClosed);
            this.Load += new System.EventHandler(this.Form3_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox chap_choise;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ComboBox TestCatagory;
        private System.Windows.Forms.Label label1;
    }
}