﻿namespace WindowsFormsApplication5
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
            this.setting = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.White;
            this.button1.Font = new System.Drawing.Font("標楷體", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button1.ForeColor = System.Drawing.Color.Black;
            this.button1.Location = new System.Drawing.Point(196, 226);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(291, 60);
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
            this.chap_choise.Location = new System.Drawing.Point(227, 334);
            this.chap_choise.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.chap_choise.Name = "chap_choise";
            this.chap_choise.Size = new System.Drawing.Size(234, 32);
            this.chap_choise.TabIndex = 2;
            this.chap_choise.SelectedIndexChanged += new System.EventHandler(this.Chap_choise_SelectedIndexChanged);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.White;
            this.button2.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button2.Location = new System.Drawing.Point(227, 527);
            this.button2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(234, 42);
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
            this.TestCatagory.Location = new System.Drawing.Point(227, 425);
            this.TestCatagory.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.TestCatagory.Name = "TestCatagory";
            this.TestCatagory.Size = new System.Drawing.Size(234, 30);
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
            this.label1.Size = new System.Drawing.Size(682, 144);
            this.label1.TabIndex = 7;
            this.label1.Text = "TestCode Generator";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // setting
            // 
            this.setting.BackColor = System.Drawing.Color.White;
            this.setting.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.setting.Font = new System.Drawing.Font("微軟正黑體", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.setting.Location = new System.Drawing.Point(571, 616);
            this.setting.Name = "setting";
            this.setting.Size = new System.Drawing.Size(87, 37);
            this.setting.TabIndex = 8;
            this.setting.Text = "設定";
            this.setting.UseVisualStyleBackColor = false;
            this.setting.Click += new System.EventHandler(this.setting_Click);
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.ClientSize = new System.Drawing.Size(682, 684);
            this.Controls.Add(this.setting);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TestCatagory);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.chap_choise);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
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
        private System.Windows.Forms.Button setting;
    }
}