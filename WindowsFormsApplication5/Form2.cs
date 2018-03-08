using System;
using System.Windows.Forms;
using System.Drawing;

namespace WindowsFormsApplication5
{
    /*UI介紹
                    第0攔             第1攔        第2攔         第3攔                                第4攔               第5欄                    第6欄
     
    第0列        初始化和新增的按鈕 |   測試類型 |      範圍       |  數量   |                         |Switch             如果選擇矩陣轉換器->row|   column
                                                                                                                       不是的話->channel      |
 i=1~TestNumber  
    
    第i列         顯示第i節      |第i節測試類型|  第i節的範圍  | 第i節switch | j=0~MaxOneTestNumber-1  | 第i節第j個switch | 第i節第j個row或channel | 第i節第j個column的輸入
                  的內容和註解   |的選擇控制項 |  選擇控制項  | 要開啟的數量  | 第 i,j列                | 的開關           | 
                                                              選擇控制項   |                         |
    */
    public partial class Form2 : Form
    {
        
        //以下變數如果有需求請自行變更
        /*-----------------------------------------*/

        //測試資料控制項之間的寬度
        int width = 200;
        
        //測試資料欄位控制項之間的高度
        int height = 150;
        
        //一個測試節的最大switch連接數
        int MaxOneTestNumber = 2;
        
        //最大的測試數目
        int MAX = 200;

        //最大的欄位數
        int MAXcol = 10;

        /*----------------------------------------*/

        //測試節的數量
        int Testnumber;
        //目前的章節
        int chapter;

        //紀錄量測點A和B
        string[,] pin_name;
        //紀錄每個測試的測試內容
        string[] ann;

        //結果
        int[,,] result;
        //excel傳入結果
        int[,,] exceldata;
        //測試名稱
        private ComboBox[] TestType;
        //總共要連接的數量
        private ComboBox[] SwitchUseNumber;
        //Switch 的連接
        private ComboBox[,] SwitchList;


        //給矩陣轉換器用的row column控制
        private TextBox[,] SwitchRow;
        private TextBox[,] SwitchCol;
        //給一般channel用的控制
        private TextBox[,] Channel;

        //給測量阻值範圍用的控制項
        private ComboBox[] ResistRange;

        //每個測試的編號顯示
        private Label[] lab;
        //每個欄位的顯示內容
        private Label[] title;

        //constructor
        public Form2(int n, int c, String[,] ss, String[,] s)
        {
            //欄位初始化，內建的可不用理他
            InitializeComponent();

            //給予控制項的宣告和初始化，n表示測試數量,c表示第幾章
            init(n, c);
            //給予註解
            for (int i = 0; i <= Testnumber; i++)
            {
                //s[i,2]表示 EXCEL K欄位
                //ann[i, 1] = s[i, 1];
                //ann[i, 2] = s[i, 2];
                ann[i] = s[i, 2];
                //s[i,0]表示 EXCEL C欄位
                //s[i,1]表示 EXCEL D欄位
                pin_name[i, 0] = s[i, 0];
                pin_name[i, 1] = s[i, 1];
            }

            //從excel讀值
            for (int i = 1; i <= Testnumber; i++)
            {
                //excel測試預設為阻抗量測
                //如果要改可能要在excel上新增給予程式知道
                exceldata[i, 1, 0] = 1;
                
                if (SwitchStringToint(ss[i, 3]) > 0)//當第2個測試物的switch(EXCEL上欄位為H)有符合條件的Switch時
                {
                    //i測試節有兩個以上switch連接時
                    exceldata[i, 3, 0] = 2;//數量為2
                    exceldata[i, 4, 0] = SwitchStringToint(ss[i, 0]);//第i節第1個測試物的switch 在EXCEL上欄位為E
                    exceldata[i, 4, 1] = SwitchStringToint(ss[i, 3]);//第i節第2個測試物的switch 在EXCEL上欄位為H
                    exceldata[i, 5, 0] = StringToInt(ss[i, 1]);//第i節第1個測試物的channel/cloumn 在EXCEL上欄位為F
                    exceldata[i, 5, 1] = StringToInt(ss[i, 4]);//第i節第2個測試物的channel/column 在EXCEL上欄位為I
                    exceldata[i, 6, 0] = StringToInt(ss[i, 2]);//第i節第1個測試物的channel/row 在EXCEL上欄位為G
                    exceldata[i, 6, 1] = StringToInt(ss[i, 5]);//第i節第2個測試物的channel/row 在EXCEL上欄位為J
                }
                else
                {
                    //i測試節只有1個switch連接時
                    exceldata[i, 3, 0] = 1;//數量為1
                    exceldata[i, 4, 0] = SwitchStringToint(ss[i, 0]);//第i節第1個測試物的switch 在EXCEL上欄位為E
                    exceldata[i, 5, 0] = StringToInt(ss[i, 1]); //第i節第1個測試物的channel/cloumn 在EXCEL上欄位為F
                    exceldata[i, 6, 0] = StringToInt(ss[i, 2]);//第i節第1個測試物的channel/row 在EXCEL上欄位為G
                }
            }
            //建立控制項
            ToolDesign();
        }
        //把Switch在excel的String轉成int
        //找不到的話傳0
        public int SwitchStringToint(String s)
        {
            int x = 0;
            switch (s)
            {
                case "A_S08": x = 1; break;
                case "B_S0203": x = 2; break;
                case "B_S0506": x = 3; break;
                case "B_S0809":x = 4; break;
                case "B_S13":x = 5; break;
                case "B_S15":x = 6; break;
                case "B_S17":x = 7; break;
                case "B_S18":x = 8; break;
                case "B_S1920":x = 9; break;
                case "B_S22":x = 10; break;
                case "B_S2425":x = 11; break;
                default:
                    x = 0;
                    break;

            }
            return x;
        }
 
        //變數的初始化，要新增欄位的宣告請放這
        public void init(int n, int c)
        {
            Testnumber = n;
            chapter = c;

            this.ann = new String[MAX];
            pin_name = new String[MAX, MAX];
            this.result = new int[MAX, MAXcol, MaxOneTestNumber];

            this.lab = new Label[MAX];
            this.title = new Label[MAXcol];

            this.TestType = new ComboBox[MAX];
            //Switch
            this.SwitchUseNumber = new ComboBox[MAX];

            this.SwitchList = new ComboBox[MAX, MaxOneTestNumber];

            this.Channel = new TextBox[MAX, MaxOneTestNumber];

            this.SwitchRow = new TextBox[MAX, MaxOneTestNumber];
            this.SwitchCol = new TextBox[MAX, MaxOneTestNumber];
            //
            this.ResistRange = new ComboBox[MAX];

            this.exceldata = new int[MAX, MAXcol, MaxOneTestNumber];

        }
        //表示這個控制項變為不使用
        //使用者無法隨意更動他也看不見
        //程式最終不會讀這個欄位值
        public void TextBoxUnused(TextBox x)
        {
            x.Visible = false;
            x.Enabled = false;
        }
        //表示這個控制項變為不使用
        //使用者無法隨意更動他也看不見
        //程式最終不會讀這個欄位值
        public void ComboBoxUnused(ComboBox x)
        {
            x.Visible = false;
            x.Enabled = false;
            x.SelectedIndex = 0;
        }
        //表示這個控制項變為使用
        //使用可以更動也看的見
        //程式會讀這個欄位值
        //如果使用者沒填會顯示錯誤
        public void TextBoxUse(TextBox x)
        {
            x.Visible = true;
            x.Enabled = true;
        }
        //表示這個控制項變為使用
        //使用可以更動也看的見
        //程式會讀這個欄位值
        //如果使用者沒填會顯示錯誤
        public void ComboBoxUse(ComboBox x)
        {
            x.Visible = true;
            x.Enabled = true;
            x.SelectedIndex = 0;
        }
        //建立新的comboBox給予指定位置和參數
        //row對應的是第row節
        //col對應每個欄位
        //commandNumber表示目前是row測試節的第幾個
        public ComboBox addComboBox(int row, int col, int CommandNumber)
        {
            ComboBox x = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                FormattingEnabled = true,
                Location = new Point(50 + width * col, height * (row + 1) + CommandNumber * 50),
                Name = "comboBox1",
                Size = new Size(150, 50),
                Font = new Font("Microsoft JhengHei", 12.25F, FontStyle.Bold, GraphicsUnit.Point, 136),
                TabIndex = 2,
                Visible = true,
                Enabled = false
            };

            this.Controls.Add(x);
            return x;
        }
        //建立新的TextBox給予指定位置和參數
        //row對應的是第row節
        //col對應每個欄位
        //commandNumber表示目前是row測試節的第幾個
        public TextBox addTextBox(int row, int col, int CommandNumber)
        {
            TextBox x = new TextBox
            {
                Location = new Point(50 + width * col, height * (row + 1) + CommandNumber * 50),
                Name = "comboBox1",
                Font = new Font("Microsoft JhengHei", 12.25F, FontStyle.Bold, GraphicsUnit.Point, 136),
                Size = new Size(120, 50),
                TabIndex = 2,
                Enabled = false,
                Visible = true
            };
            this.Controls.Add(x);
            return x;
        }
        //創立控制項並設定他的屬性，在constructor時最後會呼叫 line 122
        public void ToolDesign()
        {
            //欄位名稱
            for (int i = 0; i < 10; i++)
            {
                this.title[i] = new Label();
                this.title[i].AutoSize = true;
                this.title[i].Location = new Point(50 + width * i, 20);

                this.title[i].Font = new Font("Microsoft JhengHei", 12.25F, FontStyle.Bold, GraphicsUnit.Point, 136);
                this.title[i].Text = "";
                this.Controls.Add(title[i]);
            }
            for (int i = 0; i <= Testnumber; i++)
            {
                /*----------------------------------------*/
                //第0欄
                this.lab[i] = new Label();
                this.lab[i].AutoSize = true;
                this.lab[i].Location = new Point(10, height * (i + 1));
                this.lab[i].Font = new Font("Microsoft JhengHei", 12.25F, FontStyle.Bold, GraphicsUnit.Point, 136);
                if (i != 0)
                {
                    this.lab[i].Text = "test" + (i) + "\n";
                    //一行的最大數字
                    int oneLineTextNumber = 16;
                    
                    /*
                    //換行控制
                    for (int j=0;j< ann[i].Length;j += oneLineTextNumber)
                    {
                        //大於最大限制時++
                        if (ann[i].Length - j >= oneLineTextNumber)
                            this.lab[i].Text += ann[i].Substring(j, oneLineTextNumber) + "\n";
                        else//小於的時候
                            this.lab[i].Text += ann[i].Substring(j, ann[i].Length - j);
                    }
                    */
                   
                }
                else
                    this.lab[i].Text = "一鍵輸入";
                this.Controls.Add(lab[i]);

                /*----------------------------------------*/
                //第一欄

                //新增位置到第i個測試節的第1欄位(測試類型)
                this.TestType[i] = addComboBox(i, 1, 0);
                //新增的選擇項目
                this.TestType[i].Items.AddRange(new object[] { "未指定", "靜態阻值測試" });
                //初始預設為未指定
                this.TestType[i].SelectedIndex = 0;
                //當選擇項目改變時呼叫的function
                this.TestType[i].SelectedIndexChanged += new EventHandler(this.TestType_SelectedIndexChanged);

                /*----------------------------------------*/
                //第二欄 輸入範圍(如果選擇靜態阻抗的話)
                this.ResistRange[i] = addComboBox(i, 2, 0);
                this.ResistRange[i].Items.AddRange(new object[] {
                "未指定",
                "自動",
                "100 Ohm",
                "1   KOhm",
                "10  KOhm",
                "100 KOhm",
                "1   MOhm",
                "10  MOhm",
                "100 MOhm",
                "1   GOhm",
                });
                //初始時測試範圍為未指定
                this.ResistRange[i].SelectedIndex = 0;
                //當使用者使用一鍵輸入時
                //所呼叫的function
                if (i == 0)
                    this.ResistRange[i].SelectedIndexChanged += new EventHandler(this.ResistRange_SelectedIndexChanged);
               
                /*----------------------------------------*/
                //第三欄
                //有switch要使用時
                //輸入需要測量的通道數量

                //新增第i節的 SwitchUseNumber  到(row =i, col = 3)裡面
                this.SwitchUseNumber[i] = addComboBox(i, 3, 0);

                this.SwitchUseNumber[i].Items.Add("未指定");
                for (int j = 1; j <= MaxOneTestNumber; j++)
                    this.SwitchUseNumber[i].Items.Add(j.ToString());

                this.SwitchUseNumber[i].SelectedIndex = 0;
                this.SwitchUseNumber[i].SelectedIndexChanged += new EventHandler(this.SwitchNumber_SelectedIndexChanged);

                /*----------------------------------------*/
                //第四欄
                //產生switch的選擇
                for (int j = 0; j < MaxOneTestNumber; j++)
                {
                    //i表示是第幾個測試
                    //j表示是第i個通道，數量第幾個要開關通道
                    this.SwitchList[i, j] = addComboBox(i, 4, j);
                    //可以選擇的項目
                    this.SwitchList[i, j].Items.AddRange(new object[] {
                    "未指定",
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
                    //初始選擇為0也就是未指定
                    this.SwitchList[i, j].SelectedIndex = 0;
                    //添加當控制項被改變時要執行的function
                    this.SwitchList[i, j].SelectedIndexChanged += new EventHandler(this.switch_SelectedIndexChanged);

                }

                /*----------------------------------------*/
                //第五欄 如果(控制器是channel的話)
                for (int j = 0; j < MaxOneTestNumber; j++)
                {
                    //新增channel[i][j]的參數
                    //其中i表示第i個測試節
                    //j表示第i個測試節的數量第j個通道
                    this.Channel[i, j] = addTextBox(i, 5, j);
                    this.Channel[i, j].Text = "";
                }

                /*----------------------------------------*/
                //第五欄 如果(控制器是row的話)
                for (int j = 0; j < MaxOneTestNumber; j++)
                {
                    this.SwitchRow[i, j] = addTextBox(i, 5, j);
                    this.SwitchRow[i, j].Text = "";
                }

                /*----------------------------------------*/
                //第六欄 如果(控制器是column的話)
                for (int j = 0; j < MaxOneTestNumber; j++)
                {
                    this.SwitchCol[i, j] = addTextBox(i, 6, j);
                    this.SwitchCol[i, j].Text = "";
                }
                /*----------------------------------------*/
            }
        }
        //switch類型的選擇改變了
        public void switch_SelectedIndexChanged(object sender, EventArgs e)
        {
            int a = 0;//紀錄哪個測試節的控制項
            int b = 0;//紀錄第a項測試節的第幾個開關的switch通道

            /*---------------------------------------------------------*/
            //尋找到改變的控制項
            for (int i = 0; i <= Testnumber; i++)
            {
                for (int j = 0; j < MaxOneTestNumber; j++)
                {
                    //i表示是哪個測試節的控制項
                    //j 表示是該項測試節的第幾個開關的switch通道
                    if (sender.Equals(SwitchList[i, j]))
                    {
                        a = i;
                        b = j;
                    }
                }
            }

            /*---------------------------------------------------------*/
            //一鍵輸入
            if (a == 0)
            {
                for (int i = 1; i <= Testnumber; i++)
                    for (int j = 0; j < SwitchUseNumber[i].SelectedIndex; j++)
                    {
                        //把所有使用者可變更的switchList欄位的值全部變得跟一鍵輸入的值一樣
                        if (SwitchList[i, j].Visible == true)
                            SwitchList[i, j].SelectedIndex = SwitchList[a, 0].SelectedIndex;
                    }
            }

            /*---------------------------------------------------------*/
            //5~8表示使用矩陣轉換器
            if (SwitchList[a, b].SelectedIndex >= 5 && SwitchList[a, b].SelectedIndex <= 8)
            {
                //此時channel的欄位便不需要
                TextBoxUnused(this.Channel[a, b]);
                //把row和col的欄位變為可使用的
                TextBoxUse(this.SwitchRow[a, b]);
                TextBoxUse(this.SwitchCol[a, b]);

                //欄位名稱改變
                this.title[5].Text = "row位置";
                this.title[6].Text = "column位置";
            }
            else if (SwitchList[a, b].SelectedIndex == 0)
            {
                //都不使用
                TextBoxUnused(this.Channel[a, b]);

                TextBoxUnused(this.SwitchRow[a, b]);
                TextBoxUnused(this.SwitchCol[a, b]);

                //this.title[5].Text = "";
                //this.title[6].Text = "";
            }
            else
            {
                //使用一般的channel
                TextBoxUse(this.Channel[a, b]);
                //把矩陣轉換器移除
                TextBoxUnused(this.SwitchRow[a, b]);
                TextBoxUnused(this.SwitchCol[a, b]);

                this.title[5].Text = "Channel";
            }
            /*---------------------------------------------------------*/
        }


        //Switch的連接數量選擇
        //ex:
        //如果選1
        //表示
        //Switch.XXX_XXXX(X,1);//某通道的開啟

        //Measure

        //Switch.XXX_XXXX(X,1);;//某通道的關閉

        //如果選2
        //表示
        //Switch.XXX_XXXX(X,1);//某2個通道的開啟
        //Switch.XXX_XXXX(X,1);

        //Measure

        //Switch.XXX_XXXX(X,1);//某2個通道的關閉
        //Switch.XXX_XXXX(X,1);

        //可更改MaxOneTestNumber這個變數去增加最大數目(目前為2)
        public void SwitchNumber_SelectedIndexChanged(object sender, EventArgs e)
        {
            //尋找被使用者點選的控制項位置是在第幾row的(第幾個測試數目)
            int select = 0;

            /*---------------------------------------------------------*/
            for (int i = 1; i <= Testnumber; i++)
            {
                //當使用者選擇的控制項等於第i個測資的SwitchUseNumber時
                if (sender.Equals(SwitchUseNumber[i]))
                {
                    select = i;
                }
            }

            /*---------------------------------------------------------*/
            //一鍵輸入
            if (select == 0)
            {
                //從第一個測試資料開始跑跑到最後一個
                for (int i = 1; i <= Testnumber; i++)
                    if (SwitchUseNumber[i].Visible == true)//當Switch數量選擇這個欄位是可使用的時候(表示是一鍵輸入要改的)
                        SwitchUseNumber[i].SelectedIndex = SwitchUseNumber[select].SelectedIndex;
            }

            /*---------------------------------------------------------*/
            //初始化，使這個第select個測資下SwitchList這個欄位的全部都不使用
            for (int i = 0; i < MaxOneTestNumber; i++)
            {
                //this.title[4].Text = "";
                ComboBoxUnused(this.SwitchList[select, i]);
            }

            /*---------------------------------------------------------*/
            //使這個第select個測資下SwitchList這個欄位依照他的Swutch選擇數量而開啟

            //ex:如果select=2 ，表示第2個側資
            //如果使用者在測試項目為2，欄位為通道選擇數量的控制項填2時
            //SwitchUseNumber[select].SelectedIndex的值便為2
            //將會導致
            //SwitchList[2, 0]和SwitchList[2, 1]這兩個欄位可以被使用
            for (int i = 0; i < SwitchUseNumber[select].SelectedIndex; i++)
            {
                this.title[4].Text = "Switch";
                ComboBoxUse(this.SwitchList[select, i]);
            }
            /*---------------------------------------------------------*/
        }
        //一鍵輸入到阻抗範圍的欄位
        public void ResistRange_SelectedIndexChanged(object sender, EventArgs e)
        {
            //所有阻抗範圍的選擇項目內容變成跟一鍵輸入的一樣
            for (int i = 1; i <= Testnumber; i++)
                if(ResistRange[i].Visible)
                    ResistRange[i].SelectedIndex = ResistRange[0].SelectedIndex;
        }
        //如果測試名稱(第一欄)的控制項改變
        public void TestType_SelectedIndexChanged(object sender, EventArgs e)
        {

            //尋找物件位置
            int select = 0;

            /*---------------------------------------------------------*/
            //select表示改變的是第select個測試節的
            for (int i = 0; i <= Testnumber; i++)
            {
                if (sender.Equals(TestType[i]))
                {
                    select = i;
                }
            }

            /*---------------------------------------------------------*/
            //一鍵輸入全部
            if (select == 0)
            {
                for (int i = 1; i <= Testnumber; i++)
                    if (TestType[i].Visible == true)
                        TestType[i].SelectedIndex = TestType[select].SelectedIndex;
            }

            /*---------------------------------------------------------*/
            //testtype的選擇類型
            switch (TestType[select].SelectedIndex)
            {
                //未指定
                case 0://初始化第i節所有的控制項
                    Allinit(select);
                    //this.title[3].Text = "";
                    //this.title[2].Text = "";
                    break;
                case 1:
                    Allinit(select);
                    this.title[3].Text = "通道數量選擇";
                    this.title[2].Text = "範圍選擇";

                    //新增第i節控制項switch數量
                    ComboBoxUse(SwitchUseNumber[select]);
                    //新增第i節控制項量測範圍
                    ComboBoxUse(ResistRange[select]);
                    break;
                case 2:
                    Allinit(select);//尚未完成
                    //this.title[3].Text = "";
                    //this.title[2].Text = "";

                    break;
            }
            /*---------------------------------------------------------*/

        }
        //使第select節的所有的控制項變數回歸初始
        public void Allinit(int select)
        {
            for (int i = 0; i < MaxOneTestNumber; i++)
            {
                ComboBoxUnused(this.SwitchList[select, i]);
                TextBoxUnused(this.Channel[select, i]);
                TextBoxUnused(this.SwitchRow[select, i]);
                TextBoxUnused(this.SwitchCol[select, i]);

            }
            ComboBoxUnused(this.SwitchUseNumber[select]);
            ComboBoxUnused(this.ResistRange[select]);
        }
        //初始化 當start的按鈕點選時
        //把EXCEL的值輸入到控制項去
        private void start_Click(object sender, EventArgs e)
        {
            title[1].Text = "測試類型";
            for (int i = 0; i <= Testnumber; i++)
            {
                ComboBoxUse(this.TestType[i]);
                Allinit(i);

                this.TestType[i].SelectedIndex = exceldata[i, 1, 0];

                //通道數量
                if (exceldata[i, 3, 0] == 1)//只開起一個relay
                {
                    
                    //使用數量為1
                    ComboBoxUse(this.SwitchUseNumber[i]);
                    this.SwitchUseNumber[i].SelectedIndex = 1;//使用項目為1
                    //量測點A
                    ComboBoxUse(this.SwitchList[i, 0]);
                    this.SwitchList[i, 0].SelectedIndex = exceldata[i, 4, 0];
                    TextBoxUse(this.Channel[i, 0]);

                    //如果使用的switch為矩陣時
                    if (exceldata[i,4,0] == 5 || exceldata[i, 4, 0] == 6 || exceldata[i, 4, 0] == 7 || exceldata[i, 4, 0] == 8)
                    {
                        //此時channel的欄位便不需要
                        TextBoxUnused(this.Channel[i, 0]);
                        //把row和col的欄位變為可使用的
                        TextBoxUse(this.SwitchRow[i, 0]);
                        TextBoxUse(this.SwitchCol[i, 0]);

                        this.SwitchRow[i, 0].Text = exceldata[i, 5, 0].ToString();
                        this.SwitchCol[i, 0].Text = exceldata[i, 6, 0].ToString();
                    }
                    

                    this.Channel[i, 0].Text = exceldata[i, 5, 0].ToString();
                }
                
                if (exceldata[i, 3, 0] > 1)//開起兩個relay
                {

                    this.SwitchUseNumber[i].SelectedIndex = 2;
                    

                    //量測點A
                    ComboBoxUse(this.SwitchList[i, 0]);
                    this.SwitchList[i, 0].SelectedIndex = exceldata[i, 4, 0];
                 
                    TextBoxUse(this.Channel[i, 0]);
                    
                    this.Channel[i, 0].Text = exceldata[i, 5, 0].ToString();

                    //如果使用的switch為矩陣時
                    if (exceldata[i, 4, 0] == 5 || exceldata[i, 4, 0] == 6 || exceldata[i, 4, 0] == 7 || exceldata[i, 4, 0] == 8)
                    {
                        //此時channel的欄位便不需要
                        TextBoxUnused(this.Channel[i, 0]);
                        //把row和col的欄位變為可使用的
                        TextBoxUse(this.SwitchRow[i, 0]);
                        TextBoxUse(this.SwitchCol[i, 0]);

                        this.SwitchRow[i, 0].Text = exceldata[i, 5, 0].ToString();
                        this.SwitchCol[i, 0].Text = exceldata[i, 6, 0].ToString();
                    }

                    //量測點B
                    ComboBoxUse(this.SwitchList[i, 1]);
                    this.SwitchList[i, 1].SelectedIndex = exceldata[i, 4, 1];
                    TextBoxUse(this.Channel[i, 1]);
                    
                    this.Channel[i, 1].Text = exceldata[i, 5, 1].ToString();

                    //如果使用的switch為矩陣時
                    if (exceldata[i, 4, 1] == 5 || exceldata[i, 4, 1] == 6 || exceldata[i, 4, 1] == 7 || exceldata[i, 4, 1] == 8)
                    {
                        //此時channel的欄位便不需要
                        TextBoxUnused(this.Channel[i, 1]);
                        //把row和col的欄位變為可使用的
                        TextBoxUse(this.SwitchRow[i, 1]);
                        TextBoxUse(this.SwitchCol[i, 1]);

                        this.SwitchRow[i, 1].Text = exceldata[i, 5, 1].ToString();
                        this.SwitchCol[i, 1].Text = exceldata[i, 6, 1].ToString();
                    }
                }
            }
        }
        //避免整數轉換發生錯誤
        public static int StringToInt(String s)
        {
            try
            {
                return Int32.Parse(s);
            }
            catch (Exception)
            {
                return 0;
            }
        }
        //建立新檔案 ，當create 的按鈕被點下時
        private void newFile_Click(object sender, EventArgs e)
        {
            //輸入錯誤的警告
            String warning="";
            //判斷欄位有無正確填寫
            bool ok = true;
            
            //先將結果轉成一個三維陣列
            for (int i = 1; i <= Testnumber; i++)
            {
                for (int j = 1; j < MAXcol; j++)
                {

                    if (j <= 3)
                    {
                        switch (j)
                        {
                            
                            case 1://測試類型的結果輸入到result
                                result[i, j, 0] = TestType[i].SelectedIndex;
                                break;
                            case 2://範圍的結果輸入到result
                                result[i, j, 0] = ResistRange[i].SelectedIndex;
                                break;
                            case 3://Switch 數量的結果輸入到result
                                result[i, j, 0] = SwitchUseNumber[i].SelectedIndex;
                                break;
                        }
                        //表示裡面有值尚未輸入
                        if (result[i, j, 0] == 0)
                        {
                            ok = false;
                            warning += "第" + i + "個測資第" + j + "欄的測資未填寫完成!\n";
                        }
                    }
                    else
                    {
                        //k表示第i個測資 第j欄 的第 k 個
                        for (int k = 0; k < SwitchUseNumber[i].SelectedIndex; k++)
                        {
                            switch (j)
                            {
                                case 4://Switch的指定
                                    result[i, j, k] = SwitchList[i, k].SelectedIndex;
                                    if(result[i,j,k]==0)
                                    {
                                        ok = false;
                                        warning += "第" + i + "個測資第" + k + "個的switch未指定!\n";
                                    }
                                    break;
                                case 5://當channel是可看見時，把channel的值轉換到result
                                    if (Channel[i, k].Visible)
                                        result[i, j, k] = StringToInt(Channel[i, k].Text);
                                    else if (SwitchRow[i, k].Visible)//當Row是可看見時，把row的值轉換到result
                                        result[i, j, k] = StringToInt(SwitchRow[i, k].Text);
                                    break;
                                case 6:
                                    if (SwitchCol[i, k].Visible)//當column是可看見時，把column的值轉換到result
                                        result[i, j, k] = StringToInt(SwitchCol[i, k].Text);
                                    break;
                            }
                            if (result[i, j, k] == -1)//轉換String 到 int 的結果是-1時表示轉換失效
                            {
                                ok = false;
                                warning+="第"+i+"個測資第"+j+"欄第"+k+"個欄位有誤，必須為數字!\n";
                            }
                                
                        }
                    }
                }
            }
            if (ok)
            {
                AutoWriteCS writeCs = new AutoWriteCS(Testnumber, chapter, ann, result, pin_name);
            }
            else
            {
                MessageBox.Show(warning);
            }

        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }
    }
}
