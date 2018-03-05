using System;
using System.Drawing;
using System.Windows.Forms;

namespace WindowsFormsApplication5
{
    public partial class Form4 : Form
    {
        /*UI介紹
                       第1攔          第2攔        第3攔        第4攔       第5欄       第6欄
     
    第0列         |   輸入數量   |   輸入編號   |  通道數量  |  Switch  |    com1   |    com2   |
                                                                                                                       不是的話->channel      |
 i=1~TestNumber  
    
    第i列         |第i節測試類型 |  第i節的範圍  | 第i節switch | j=0~MaxOneTestNumber-1  | 第i節第j個switch | 第i節第j個row或channel | 第i節第j個column的輸入
                  |的選擇控制項  |  選擇控制項  | 要開啟的數量  | 第 i,j列                | 的開關           | 
                                                              選擇控制項   |                         |*/

        //以下變數如果有需求請自行變更
        /*-----------------------------------------*/

        //測試資料欄位控制項之間的高度
        int height = 150;

        //一個測試節的最大switch連接數
        int MaxOneTestNumber = 2;

        //最大的測試數目
        int MAX = 200;

        /*----------------------------------------*/

        //測試節的數量
        int Testnumber;

        //紀錄測試輸入編號
        String[,] test_number;

        //總共要連接的數量
        private ComboBox[] SwitchUseNumber;
        //Switch 的連接
        private ComboBox[,] SwitchList;


        //給矩陣轉換器用的com1,com2控制
        private TextBox[,] SwitchCom1;
        private TextBox[,] SwitchCom2;

        //單個測試使用的輸入數量
        private TextBox[] Input_Total_Num;
        private TextBox[] Input_Num;

        //給一般channel用的控制
        private TextBox[,] Channel;


        //每個測試的編號顯示
        private System.Windows.Forms.Label[] lab;

        public InputData[] Inputdata;
        //改完後的輸出資料
        public string[,] Final_OutputData;
        //單個測試輸入資料
        public string[,] outcome;

        public string[,] OutputData1;

        int ChapterNumber;

        //constructor
        public Form4(int ChapterNumber, int InputNumber, InputData[] inputData, int TestNumber, String[,] OutputData)
        {
            init();
            this.OutputData1 = OutputData;
            this.Inputdata = inputData;
            this.Testnumber = TestNumber;
            InitializeComponent();
            //給予控制項的宣告和初始化
            label1.Text = "Chapter" + ChapterNumber.ToString();//章節名稱
            test_number = split(OutputData, TestNumber);
            this.ChapterNumber = ChapterNumber;

            for (int i = 0; i < TestNumber; i++)
            {
                //初始test 標題
                this.lab[i] = new System.Windows.Forms.Label();
                this.lab[i].AutoSize = true;
                this.lab[i].Location = new System.Drawing.Point(10, height * (i + 1));
                this.lab[i].Font = new System.Drawing.Font("Microsoft JhengHei", 12.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                this.lab[i].Text = "Test" + (i + 1).ToString() + "\n" + "量測目的:" + "\n" + OutputData[i + 1, 10] + "\n" + "量測點A:" + OutputData[i + 1, 2] + "\n" + "量測點B:" + OutputData[i + 1, 3];
                this.Controls.Add(lab[i]);

                //初始輸入編號數量 第一欄
                this.Input_Total_Num[i] = addTextBox(165, i, 0, 0);
                this.Input_Total_Num[i].Text = test_number[i, 0];
                //this.Input_Total_Num[i].TextChanged += new System.EventHandler(this.Input_Total_Num_TextChanged);

                //初始輸入編號 第二欄
                this.Input_Num[i] = addTextBox(325, i, 0, 0);
                this.Input_Num[i].Text = OutputData[i + 1, 1];
                /*for (int j=0; j<Int32.Parse(test_number[i,0]);j++)
                {
                    this.Input_Num[i,j] = addTextBox(325, i, j,0);
                    this.Input_Num[i, j].Text =test_number[i, j+1];
                }*/

                //初始通道數量 第三欄
                this.SwitchUseNumber[i] = addComboBox(477, i, 0);
                this.SwitchUseNumber[i].Items.Add("1");
                this.SwitchUseNumber[i].Items.Add("2");
                this.SwitchUseNumber[i].SelectedIndex = 1;
                this.SwitchUseNumber[i].SelectedIndexChanged += new System.EventHandler(this.SwitchNumber_SelectedIndexChanged);

                //第四欄 初始switch
                //第五欄 初始com1
                //第六欄 初始com2
                for (int j = 0; j < MaxOneTestNumber; j++)
                {
                    //產生switch的選擇
                    //i表示是第幾個測試
                    //j表示是第i個通道，數量第幾個要開關通道
                    this.SwitchList[i, j] = addComboBox(620, i, j);
                    //可以選擇的項目
                    this.SwitchList[i, j].Items.AddRange(new object[] {
                    "A_S08",
                    "B_S0203",
                    "B_S0506",
                    "B_S0809",
                    "B_S13",
                    "B_S15",
                    "B_S17",
                    "B_S18",
                    "B_S1920",
                    "B_S22",
                    "B_S2425",
                    });

                    for (int k = 0; k < this.SwitchList[i, j].Items.Count; k++)
                    {
                        if (j == 0 && this.SwitchList[i, j].Items[k].Equals(OutputData[i + 1, 4]))
                        {
                            this.SwitchList[i, j].SelectedIndex = k;
                        }
                        if (j == 1 && this.SwitchList[i, j].Items[k].Equals(OutputData[i + 1, 7]))
                        {
                            this.SwitchList[i, j].SelectedIndex = k;
                        }
                    }
                    //添加當控制項被改變時要執行的function
                    //this.SwitchList[i, j].SelectedIndexChanged += new System.EventHandler(this.switch_SelectedIndexChanged);
                }

                //com1初始化
                for (int j = 0; j < MaxOneTestNumber; j++)
                {
                    this.SwitchCom1[i, j] = addTextBox(751, i, j, 0);
                    if (j == 0 && OutputData[0, 5].Equals("COM1"))
                    {
                        this.SwitchCom1[i, j].Text = OutputData[i + 1, 5];
                    }
                    if (j == 1 && OutputData[0, 8].Equals("COM1"))
                    {
                        this.SwitchCom1[i, j].Text = OutputData[i + 1, 8];
                    }
                }

                //com2初始化
                for (int j = 0; j < MaxOneTestNumber; j++)
                {
                    this.SwitchCom2[i, j] = addTextBox(885, i, j, 0);
                    if (j == 0 && OutputData[0, 6].Equals("COM2"))
                    {
                        this.SwitchCom2[i, j].Text = OutputData[i + 1, 6];
                    }
                    if (j == 1 && OutputData[0, 9].Equals("COM2"))
                    {
                        this.SwitchCom2[i, j].Text = OutputData[i + 1, 9];
                    }
                }
            }
        }
        //給予控制項的宣告和初始化
        public void init()
        {
            this.Inputdata = new InputData[100];
            this.OutputData1 = new string[MAX, MAX];
            //總relay通道數量
            //標題總數宣告
            this.lab = new System.Windows.Forms.Label[MAX];

            //使用輸入總數量宣告
            this.Input_Total_Num = new TextBox[MAX];
            //                                     col[0]        col[1]
            //把分割完後的資料放這    測試編號    輸入總數      輸入編號
            //                       row[0]          2              1           2
            //                       row[1]          3              1           2           4
            this.test_number = new String[MAX, MAX];

            //使用輸入編號 row:第幾個測試 col:輸入編號
            this.Input_Num = new TextBox[MAX];

            //switch總數宣告
            this.SwitchUseNumber = new ComboBox[MAX];

            //每個test使用的 switch
            this.SwitchList = new ComboBox[MAX, MAX];

            //com1宣告
            this.SwitchCom1 = new TextBox[MAX, MaxOneTestNumber];

            //com2宣告
            this.SwitchCom2 = new TextBox[MAX, MaxOneTestNumber];

            this.Final_OutputData = new string[MAX, MAX];

            this.outcome = new string[MAX, MAX];
        }

        //把輸入標號用","分開
        public string[,] split(string[,] s, int TestNumber)
        {
            string[,] outcome = new string[MAX, MAX];
            for (int i = 0, j = 0; i < TestNumber; i++)
            {
                foreach (string str in s[i + 1, 1].Split(','))
                {
                    outcome[i, j + 1] = str;
                    j++;
                }
                outcome[i, 0] = j.ToString();
                j = 0;
            }
            return outcome;
        }

        //建立新的comboBox給予指定位置和參數
        //row對應的是第row節
        //col對應每個欄位
        //commandNumber表示目前是row測試節的第幾個
        public ComboBox addComboBox(int row, int col, int CommandNumber)
        {
            ComboBox x = new ComboBox();
            x.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            x.FormattingEnabled = true;
            x.Location = new System.Drawing.Point(row, height * (col + 1) + CommandNumber * 50);
            x.Name = "comboBox1";
            x.Size = new System.Drawing.Size(90, 11);
            x.Font = new System.Drawing.Font("Microsoft JhengHei", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            x.TabIndex = 2;
            x.Visible = true;
            x.Enabled = true;
            this.Controls.Add(x);
            return x;
        }
        //建立新的TextBox給予指定位置和參數
        //row對應的是第row節
        //col對應每個欄位
        //commandNumber表示目前是row測試節的第幾個
        public TextBox addTextBox(int row, int col, int CommandNumber, int initial)
        {
            TextBox x = new TextBox();
            x.Location = new Point(row, height * (col + 1) + CommandNumber * 40 + initial);
            x.Font = new System.Drawing.Font("Microsoft JhengHei", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            x.Size = new System.Drawing.Size(90, 11);
            x.TabIndex = 2;
            x.Enabled = true;
            x.Visible = true;
            this.Controls.Add(x);
            return x;
        }

        //可更改MaxOneTestNumber這個變數去增加最大數目(目前為2)
        public void SwitchNumber_SelectedIndexChanged(object sender, EventArgs e)
        {
            //尋找被使用者點選的控制項位置是在第幾row的(第幾個測試數目)
            //從0開始 要加1
            int select = 0;
            int switch_use_number;
            /*---------------------------------------------------------*/
            for (int i = 0; i < Testnumber; i++)
            {
                //當使用者選擇的控制項等於第i個測資的SwitchUseNumber時
                if (sender.Equals(SwitchUseNumber[i]))
                {
                    select = i;
                }
            }
            //看初始值是多少選擇是否要建新的 最多兩個switch
            switch_use_number = Convert.ToInt32(SwitchUseNumber[select].SelectedItem);

            if (SwitchList[select, switch_use_number - 1] == null)
            {
                for (int i = 0; i < switch_use_number; i++)
                {
                    if (SwitchList[select, i] == null)
                    {
                        this.SwitchList[select, i] = addComboBox(620, i, i);
                        //可以選擇的項目
                        this.SwitchList[select, i].Items.AddRange(new object[] {
                        "A_S08",
                        "B_S0203",
                        "B_S0506",
                        "B_S0809",
                        "B_S13",
                        "B_S15",
                        "B_S17",
                        "B_S18",
                        "B_S1920",
                        "B_S22",
                        "B_S2425",
                        });
                    }
                }
            }
            else if (SwitchList[select, switch_use_number - 1] != null)
            {
                if (switch_use_number == 1)
                {
                    ComboBoxUnUsed(SwitchList[select, 1]);
                    TextBoxUnUsed(SwitchCom1[select, 1]);
                    TextBoxUnUsed(SwitchCom2[select, 1]);
                }
                else if (switch_use_number == 2)
                {
                    ComboBoxUsed(SwitchList[select, 1]);
                    TextBoxUsed(SwitchCom1[select, 1]);
                    TextBoxUsed(SwitchCom2[select, 1]);
                }
            }
            /*---------------------------------------------------------*/
            //使這個第select個測資下SwitchList這個欄位依照他的Swutch選擇數量而開啟

            //ex:如果select=2 ，表示第2個側資
            //如果使用者在測試項目為2，欄位為通道選擇數量的控制項填2時
            //SwitchUseNumber[select].SelectedIndex的值便為1
            //將會導致
            //SwitchList[1, 0]和SwitchList[1, 1]這兩個欄位可以被使用

            /*---------------------------------------------------------*/
        }
        /*失敗
        public void Input_Total_Num_TextChanged(object sender, EventArgs e)
        {
            //尋找被使用者點選的控制項位置是在第幾row的(第幾個測試數目)
            //從0開始 要加1
            int select = 0,change_inp, input_total_number = 0;
            string change_input_total_number;
            for (int i = 0; i < Testnumber; i++)
            {
                //當使用者選擇的控制項等於第i個測資的SwitchUseNumber時
                if (sender.Equals(Input_Total_Num[i]))
                {
                    select = i;
                }
            }
            for(int i=0; Input_Num[select,i]!=null ; i++)
            {
                input_total_number++;
            }
            change_input_total_number = Input_Total_Num[select].Text;
            if (string.IsNullOrEmpty(change_input_total_number))
            {

            }
            else
            {
                change_inp = Int32.Parse(change_input_total_number);
                if (input_total_number<change_inp)
                {
                    //要新增幾個"測試編號"格子
                    label2.Text = this.Input_Num[select, input_total_number - 1].Location.Y.ToString();
                    this.Input_Num[select, input_total_number] = addTextBox(325, 0, 1, this.Input_Num[select, input_total_number - 1].Location.Y);
                    for (int i = 0; i < change_inp-input_total_number; i++)
                    {
                        
                        input_total_number++;
                    }
                    
                }
                else
                {
                }
            }

    }*/
        public void ComboBoxUsed(ComboBox x)
        {
            x.Visible = true;
            x.Enabled = true;
            // x.SelectedIndex = 0;
        }
        public void ComboBoxUnUsed(ComboBox x)
        {
            x.Visible = false;
            x.Enabled = false;
            //x.SelectedIndex = 0;
        }
        //表示這個控制項變為使用
        //使用可以更動也看的見
        //程式會讀這個欄位值
        //如果使用者沒填會顯示錯誤
        public void TextBoxUsed(TextBox x)
        {
            x.Visible = true;
            x.Enabled = true;
        }
        //表示這個控制項變為不使用
        //使用者無法隨意更動他也看不見
        //程式最終不會讀這個欄位值
        public void TextBoxUnUsed(TextBox x)
        {
            x.Visible = false;
            x.Enabled = false;
        }
        //把Switch在excel的String轉成int
        //找不到的話傳0
        public int SwitchStringToint(String s)
        {
            int x = 0;
            switch (s)
            {
                case "A08": x = 1; break;
                case "B0203": x = 2; break;
                case "B0506": x = 3; break;
                case "B0809": x = 4; break;
                case "B13": x = 5; break;
                case "B15": x = 6; break;
                case "B17": x = 7; break;
                case "B18": x = 8; break;
                case "B1920": x = 9; break;
                case "B22": x = 10; break;
                case "B2425": x = 11; break;
                default:
                    x = 0;
                    break;
            }
            return x;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[,] outcome1 = new string[MAX, MAX];
            //0錯1對
            int check = 1;
            //col[0]  
            //  編號
            for (int i = 0, j = 0; i < Testnumber; i++)
            {
                foreach (string str in Input_Num[i].Text.Split(','))
                {
                    this.outcome[i, j + 1] = str;
                    outcome1[i, j] = str;
                    j++;
                }
                this.outcome[i, 0] = j.ToString();
                j = 0;
            }
            //檢查是否有漏填
            for (int i = 0; i < Testnumber; i++)
            {
                if (string.IsNullOrEmpty(Input_Total_Num[i].Text))
                {
                    check = 0;
                    MessageBox.Show("第" + (i + 1).ToString() + "個測試的輸入數量未填");
                }
                if (string.IsNullOrEmpty(Input_Num[i].Text))
                {
                    check = 0;
                    MessageBox.Show("第" + (i + 1).ToString() + "個測試的輸入編號未填");
                }
                if (!Input_Total_Num[i].Text.Equals(outcome[i, 0]))
                {
                    check = 0;
                    MessageBox.Show("第" + (i + 1).ToString() + "個測試的輸入數量與輸入編號未填不符");
                }

                for (int j = 0; j < MaxOneTestNumber; j++)
                {
                    if (SwitchCom1[i, j].Enabled == true && SwitchCom2[i, j].Enabled == true)
                    {
                        if (string.IsNullOrEmpty(SwitchCom1[i, j].Text))
                        {
                            check = 0;
                            MessageBox.Show("第" + (i + 1).ToString() + "個測試的COM1未填");
                        }
                        if (SwitchList[i, j].SelectedItem.ToString() == "B_S13" || SwitchList[i, j].SelectedItem.ToString() == "B_S15" || SwitchList[i, j].SelectedItem.ToString() == "B_S17" || SwitchList[i, j].SelectedItem.ToString() == "B_S18")
                        {
                            if (string.IsNullOrEmpty(SwitchCom2[i, j].Text))
                            {
                                check = 0;
                                MessageBox.Show("第" + (i + 1).ToString() + "個測試的COM2未填");
                            }
                        }
                    }
                }
            }
            if (check == 1)
            {
                for (int i = 0; i < 11;)
                {
                    if (!OutputData1[0, i].Contains("輸入編號"))
                    {
                        Final_OutputData[0, i] = OutputData1[0, i];
                    }
                    i++;
                }
                //10列
                for (int i = 1, j = 0; i <= Testnumber; i++, j++)
                {
                    //測試編號 0
                    Final_OutputData[i, 0] = OutputData1[i, 0];
                    //量測點A 1
                    Final_OutputData[i, 1] = OutputData1[i, 2];
                    //量測點B 2
                    Final_OutputData[i, 2] = OutputData1[i, 3];
                    //量測點A連接的relay 3
                    //com1 4
                    //com2 5
                    //量測點B連接的relay 6
                    //com1 7
                    //com2 8
                    if (Convert.ToInt32(SwitchUseNumber[j].SelectedItem) == 1)
                    {
                        Final_OutputData[i, 3] = SwitchList[j, 0].SelectedItem.ToString();
                        Final_OutputData[i, 6] = "null";
                        Final_OutputData[i, 4] = SwitchCom1[j, 0].Text;
                        if (!String.IsNullOrEmpty(SwitchCom2[j, 0].Text))
                        {
                            Final_OutputData[i, 5] = SwitchCom2[j, 0].Text;
                        }
                        else
                        {
                            Final_OutputData[i, 5] = "null";
                        }
                        Final_OutputData[i, 7] = "null";
                        Final_OutputData[i, 8] = "null";
                    }
                    else if (Convert.ToInt32(SwitchUseNumber[j].SelectedItem) == 2)
                    {
                        Final_OutputData[i, 3] = SwitchList[j, 0].SelectedItem.ToString();
                        Final_OutputData[i, 6] = SwitchList[j, 1].SelectedItem.ToString();
                        Final_OutputData[i, 4] = SwitchCom1[j, 0].Text;
                        if (!String.IsNullOrEmpty(SwitchCom2[j, 0].Text))
                        {
                            Final_OutputData[i, 5] = SwitchCom2[j, 0].Text;
                        }
                        else
                        {
                            Final_OutputData[i, 5] = "null";
                        }
                        Final_OutputData[i, 7] = SwitchCom1[j, 1].Text;
                        if (!String.IsNullOrEmpty(SwitchCom2[j, 1].Text))
                        {
                            Final_OutputData[i, 8] = SwitchCom2[j, 1].Text;
                        }
                        else
                        {
                            Final_OutputData[i, 8] = "null";
                        }
                        // Final_OutputData[i, 8] = SwitchCom1[j, 1].Text;
                    }
                    Final_OutputData[i, 9] = OutputData1[i, 10];
                }
                /*for(int i = 0;i<Testnumber;i++)
                {
                    for(int j = 0;j<Testnumber;j++)
                    {
                        label2.Text += Final_OutputData[i, j] + "  ";
                    }
                    label2.Text += "\n";
                }*/                                                      //從col 1開始  
                //                                    測試數量   輸入資料   輸入編號      10列
                DynamicWriteCS writeCS = new DynamicWriteCS(ChapterNumber, Testnumber, Inputdata, outcome1, Final_OutputData);
            }
        }
        
        private void Form4_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
          
