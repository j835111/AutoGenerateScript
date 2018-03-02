using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;

namespace WindowsFormsApplication5
{
    class AutoWriteCS
    {
        //最大行數
        int MAXline = 20000;
        String[] SwitchName ={
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
                    };
        String[] ann;
        //記錄總行數
        int lineNumber;
        //紀錄排版數
        int spaceNumber;
        //紀錄測試的數量
        int testNumber;
        //紀錄此次要寫的章節
        int chapter;
        int[,,] result;
        String[] line;
        /*------------------------------*/
        int MAX = 200;
        InputData[] inputData;
        string[,] each_InputNum;
        string[,] Final_OutputData;
        //Constructor
        public AutoWriteCS(int n,int c,String[] a,int[,,] x)
        {
            //初始化
            init();
            testNumber = n;
            chapter = c;
            //紀錄註解
            for (int i = 0; i < testNumber; i++)
                ann[i] = a[i];
            //紀錄result
            result = x;
            makeFile();
        }
        //動態
        public AutoWriteCS(int chapter,int Testnumber,InputData[] inputData,string[,] each_InputNum,string[,] Final_OutputData)
        {
            this.testNumber = Testnumber;
            this.inputData = new InputData[100];
            this.each_InputNum = new string[MAX,MAX];
            this.Final_OutputData = new string[MAX, MAX];
            this.inputData = inputData;
            this.each_InputNum = each_InputNum;
            this.Final_OutputData = Final_OutputData;
            this.chapter = chapter;
            init();
            Dynamic_makeFile();
        }
        void init()
        {
            result = new Int32[MAXline, 10, 5];
            ann = new String[MAXline];
            line = new String[MAXline];
            lineNumber = 1;
            spaceNumber = 0;
            testNumber = 0;
             
        }
        //用來寫入每一行的程式碼，包含排版
        public void WriteLine(String text)
        {
            if(lineNumber==MAXline)
            {
                MessageBox.Show("超過最大行數!!!");
                return;
            }
            //line本身初始化
            line[lineNumber] = "";
            //排版用
            if (text!=""&&String.Compare(text, "}") == 0)
                spaceNumber--;
            //排版用
            for (int i = 0; i < spaceNumber; i++)
                line[lineNumber] += "    ";

            //填入實際的內容
            line[lineNumber++] += text;
            //排版用
            if (text != ""&&String.Compare(text, "{") == 0)
                spaceNumber++;
        }
        /*
              mode
              1 時表示 Measure的指令
              2 時表示 Switch.FORMARLY開啟的指令
              3 時表示 Switch.Diffrly開啟的指令
              4 時表示 DCV
              5 時表示 AC
              test 表示第幾個測試項目
              onoff 表示開或關
              number 在switch上表示第test項目的第幾個通道連接
                     在測量上表示是哪個模式的測量方式
            */
        public void WriteCommand(int test, int number ,int mode, int on_off)
        { 
            //寫入mode類型的指令
            switch (mode)
            {
                /*---------------------------------------------------------*/
                case 1://表示 Measure的指令

                    if (result[test, 2, 0] >=2)//表示範圍值是有指定的
                    {
                        WriteLine("//測量有線電阻 範圍為" + result[test, 2, 0]);
                        WriteLine("UUT_MainFunction.g_objMeasure_Data = DMM.DMMhp34461A_meas(" + number + "," + result[test, 2, 0] + ");");
                    }
                    else 
                    {
                        WriteLine("//測量有線電阻範圍自動");
                        WriteLine("UUT_MainFunction.g_objMeasure_Data = DMM.DMMhp34461A_meas(" + number + ");");
                    }
                    break;
                
                /*---------------------------------------------------------*/
                case 2://表示 Switch.FORMARLY開關的指令
                        
                        //Switch連接註解
                        if(on_off==1)//開啟 
                            WriteLine("//開啟Switch" + SwitchName[result[test, 4, number]]+"的通道"+ result[test, 5, number]);
                        else//關閉
                            WriteLine("//關閉Switch" + SwitchName[result[test, 4, number]] + "的通道" + result[test, 5, number]);

                        //寫入Switch.FORMARLY程式碼
                        WriteLine("Switch.FORMARLY_" + SwitchName[result[test, 4, number]] + "(" + result[test, 5, number] + "," + on_off + ");");
                        break;

                /*---------------------------------------------------------*/
                case 3://Switch.Diffrly開關的指令
                        WriteLine("Switch.DIFFRLY_" + SwitchName[result[test, 4, number]] + "(" + result[test, 5, number] + "," + result[test, 6, number] + "," + on_off + ");");
                        break;
                case 4:
                        //if (on_off == 1)
                            //WriteLine("DCV.DCV_ON(" + dataTable.Rows[i][2] + "," + dataTable.Rows[i][3] + "," + dataTable.Rows[i][4] + ");");
                        //else
                            //WriteLine("DCV.DCV_OFF(" + dataTable.Rows[i][2] + ");");
                        break;
                case 5:
                        //if (on_off == 1)
                            //WriteLine("AC_Source.AC_Source_CPS6000_ON(" + dataTable.Rows[i][2] + "," + dataTable.Rows[i][3] + "," + dataTable.Rows[i][4] + ");");
                        //else
                            //WriteLine("AC_Source.AC_Source_CPS6000_OFF();");
                        break;
               
            }
            WriteLine("");
        }

        public void WriteTest(int test)
        {
            //第test節的註解
            WriteLine("//"+ann[test]);
            //名稱
            WriteLine("public static void test" + chapter + "_" + (test) + "()");
            WriteLine("{");

            //測試架構
            //*---------------------------------------------------------*/
            WriteLine("if (UutModule.test_interrupt == 1)");
            WriteLine("{");
            WriteLine("return;");
            WriteLine("}");

            WriteLine("");
            WriteLine("UutModule.pre_triger();");
            WriteLine("UutModule.pre_probe();");
            /*---------------------------------------------------------*/

            //給予Stimuli
            /*---------------------------------------------------------*/
            WriteLine(" //*************** Stimuli ***********************************");
            WriteLine("");
            //確認有無開啟DCV電流
            //WriteCommand(test, 4, 1);
            //WriteCommand(test, 5, 1)

            //*---------------------------------------------------------*/

            //給予Measure
            /*---------------------------------------------------------*/
            //measure 目前步驟
            //1.先開啟要測的switch和通道
            //2.用dmm測量
            //3.關閉要測的switch和通道
            WriteLine(" //*************** Measure ***********************************");
            WriteLine("");
            /*---------------------------------------------------------*/
            //確認是否有要Switch開啟的指令並寫入
            for (int k = 0; k < result[test, 3, 0]; k++)//result[test, 3, 0]為連接的Switch數量
            {
                //連接哪個Switch 欄位4為switch選擇
                if (result[test, 4, k] >= 5 && result[test, 4, k] <= 8)
                {
                    //矩陣轉換器時
                    WriteCommand(test,k, 3, 1);//mode=3 ,on_off=1 ,k表示通道數
                }
                else //其他時channel
                    WriteCommand(test,k,2, 1);//mode=2 ,on_off=1,k表示通道數
            }

            /*---------------------------------------------------------*/
            //確認是否有測量的種類
            switch (result[test,1,0])
            {
                //靜態阻抗
                case 1:
                    if (result[test, 2, 0] >= 2)//表示有輸入範圍
                    {
                        WriteCommand(test, 13, 1, -1);//模式為13
                    }
                    else//表示沒有輸入範圍
                    {
                        WriteCommand(test, 6, 1, -1);//模式為6
                    }
                    break;
            }

            /*---------------------------------------------------------*/
            //確認是否有要Switch關閉的指令並寫入
            for (int k = 0; k < result[test, 3, 0]; k++)//連接數量
            {
                //連接哪個Switch
                if (result[test, 4, k] >= 5 && result[test, 4, k] <= 8)
                {
                    //矩陣轉換器時
                    WriteCommand(test, k, 3, 0);
                }
                else //其他時channel
                    WriteCommand(test, k, 2, 0);
            }

            /*---------------------------------------------------------*/
            WriteLine(" //*************** Stimuli Off after test *************************");
            WriteLine("");
            WriteLine("UutModule.pre_reset();");
            WriteLine("UutModule.post_reset();");
            WriteLine("}");

        }
        //填一個檔案的初始架構程式碼
        public void WriteChapter()
        {
            WriteLine("using Microsoft.VisualBasic;");
            WriteLine("using System;");
            WriteLine("using System.Collections;");
            WriteLine("using System.Data;");
            WriteLine("using System.Diagnostics;");
            WriteLine("using System.Drawing;");
            WriteLine("using System.Windows.Forms;");
            WriteLine("");

            WriteLine("namespace UUTapplication1");
            WriteLine("{");
            WriteLine("//靜態阻抗測試");
            WriteLine("static class Chapter_" + chapter.ToString());
            WriteLine("{");

            //第i個指定內容的填寫
            /*---------------------------------------------------------*/
            for (int i = 1; i <= testNumber; i++)
            {
                //填每個測試內容
                WriteTest(i);
            }
            /*---------------------------------------------------------*/

            WriteLine("}");
            WriteLine("}");
        }

        public void makeFile()
        {
            //寫入程式碼到line的資料裡面
            WriteChapter();
            //儲存檔案的定義
            SaveFileDialog save = new SaveFileDialog();
            //儲存檔案的名稱
            save.FileName = "chap" + chapter.ToString() + ".cs";
            //允許的副檔名
            save.Filter = "C# File | *.cs";
            //當成功儲存時
            if (save.ShowDialog() == DialogResult.OK)
            {
                //將line的資料寫入write裡面並寫入新創的檔案
                StreamWriter writer = new StreamWriter(save.OpenFile());
                for (int i = 1; i <= lineNumber; i++)
                {
                    writer.WriteLine(line[i]);
                }
                writer.Dispose();
                writer.Close();
                MessageBox.Show("Done!");
            }
        }

        //以下為動態
        /*-----------------------------------------------------------------------------------*/
        public void Dynamic_makeFile()
        {
            //寫入程式碼到line的資料裡面
            Dynamic_WriteChapter();
            //儲存檔案的定義
            SaveFileDialog save = new SaveFileDialog();
            //儲存檔案的名稱
            save.FileName = "chap" + chapter.ToString() + ".cs";
            //允許的副檔名
            save.Filter = "C# File | *.cs";
            //當成功儲存時
            if (save.ShowDialog() == DialogResult.OK)
            {
                //將line的資料寫入write裡面並寫入新創的檔案
                StreamWriter writer = new StreamWriter(save.OpenFile());
                for (int i = 1; i <= lineNumber; i++)
                {
                    writer.WriteLine(line[i]);
                }
                writer.Dispose();
                writer.Close();
                MessageBox.Show("Done!");
            }
        }
        //填一個檔案的初始架構程式碼
        public void Dynamic_WriteChapter()
        {
            WriteLine("using Microsoft.VisualBasic;");
            WriteLine("using System;");
            WriteLine("using System.Collections;");
            WriteLine("using System.Data;");
            WriteLine("using System.Diagnostics;");
            WriteLine("using System.Drawing;");
            WriteLine("using System.Windows.Forms;");
            WriteLine("");

            WriteLine("namespace UUTapplication1");
            WriteLine("{");
            WriteLine("//動態測試");
            WriteLine("static class Chapter_" + chapter.ToString());
            WriteLine("{");

            //第i個指定內容的填寫
            /*---------------------------------------------------------*/
            for (int i = 0; i < testNumber; i++)
            {
                //填每個測試內容
                Dynamic_WriteTest(i);
            }
            /*---------------------------------------------------------*/

            WriteLine("}");
            WriteLine("}");
        }
        public void Dynamic_WriteTest(int test)
        {
            //第test節的註解
            //名稱
            WriteLine("public static void test" + chapter.ToString() + "_" + (test+1).ToString() + "()");
            WriteLine("{");

            //測試架構
            //*---------------------------------------------------------*/
            WriteLine("if (UutModule.test_interrupt == 1)");
            WriteLine("{");
            WriteLine("return;");
            WriteLine("}");

            WriteLine("");
            WriteLine("UutModule.pre_triger();");
            WriteLine("UutModule.pre_probe();");
            /*---------------------------------------------------------*/

            //給予Stimuli
            /*---------------------------------------------------------*/
            WriteLine(" //*************** Stimuli ***********************************");
            WriteLine("");
            //第i個測試使用哪些輸入 這些輸入用了哪些function
            //單個測試
            for (int i = 0; each_InputNum[test, i] != null; i++)
            {
                //打開
                /*---------------------------------------------------------*/

                //輸入編號 
                int inputNumber = Int32.Parse(each_InputNum[test, i]) - 1;//因為輸入編號從0開始存，所以要減1
                //單個輸入開的多個switch
                string[] s = new string[MAX];
                s = inputData[inputNumber].GetRelay;
                WriteLine(" //*************** Measure ***********************************");
                WriteLine("");
                //1.先開啟要送電的switch和通道
                for (int k = 0; !string.IsNullOrEmpty(s[k]); k++)//s[0]--s[1]---
                {
                    //連接哪個Switch 
                    Dynamic_WriteSwitch(s[k], inputNumber, 1, 1);//order=1 on_off=1
                }
                //2.使用送電儀器
                switch (inputData[inputNumber].GetInstrument)
                {
                    //選擇使用function
                    case "AC_Source":
                        AC_Source(inputNumber, 1);
                        break;
                    case "AWG":
                        AWG(inputNumber, 1);
                        break;
                    case "DCV":
                        DCV(inputNumber, 1);
                        break;
                }
                //3.開啟要量測的switch腳位
                //兩個switch
                if (i == 0)//只開一次
                {
                    Dynamic_WriteSwitch(Final_OutputData[test, 3], test, 0, 1);//order=0 on_off=0 
                    if (!Final_OutputData[test, 6].Equals("NULL"))
                    {
                        Dynamic_WriteSwitch(Final_OutputData[test, 6], test, 0, 1);
                    }
                }
                //4.量測
                switch (inputData[inputNumber].GetInstrument)
                {
                    case "DMM":
                        DMM(inputNumber);
                        break;
                    case "LCR":
                        LCR(inputNumber);
                        break;
                }
            }
            /*----------------------------------*/
            //關閉
            for (int i = 0; each_InputNum[test, i] != null; i++)
            {
                //輸入編號 
                int inputNumber = Int32.Parse(each_InputNum[test, i]) - 1;//因為輸入編號從0開始存，所以要減1
                //單個輸入開的多個switch
                string[] s = new string[MAX];
                s = inputData[inputNumber].GetRelay;
                //5.關閉電源
                switch (inputData[inputNumber].GetInstrument)
                {
                    //選擇使用function
                    case "AC_Source":
                        AC_Source(inputNumber, 0);
                        break;
                    case "AWG":
                        AWG(inputNumber, 0);
                        break;
                    case "DCV":
                        DCV(inputNumber, 0);
                        break;
                }
                //6.關閉送電的switch通道
                for (int k = 0; !string.IsNullOrEmpty(s[k]); k++)//s[0]--s[1]---
                {
                    //關閉哪個Switch 
                    Dynamic_WriteSwitch(s[k], inputNumber, 1, 0);//order=1 on_off=1
                }

                //6.關閉要測的switch和通道
                //兩個switch
                if (i == 0)//只關一次
                {
                    Dynamic_WriteSwitch(Final_OutputData[test, 3], test, 0, 0);//order=0 on_off=0 
                    if (!Final_OutputData[test, 6].Equals("NULL"))
                    {
                        Dynamic_WriteSwitch(Final_OutputData[test, 6], test, 0, 1);
                    }
                }
            }

                /*---------------------------------------------------------*/
            WriteLine(" //*************** Stimuli Off after test *************************");
            WriteLine("");
            WriteLine("UutModule.pre_reset();");
            WriteLine("UutModule.post_reset();");
            WriteLine("}");
        }                                                                    //1 :輸入的switch 0:測試的switch
        public void Dynamic_WriteSwitch(string switchName, int inputnumber,int order,int on_off)
        {
            //num: 看是第幾個com[num]
            string[] com1 = new string[MAX];
            com1 = inputData[inputnumber].GetCom1;
            string[] com2 = new string[MAX];
            com2 = inputData[inputnumber].GetCom2;
            string[] pin_name = new string[MAX];
            pin_name = inputData[inputnumber].GetPin_name;
            //switchNumber: 第幾個switch
            //表示 Switch.FORMARLY開關的指令
            //Switch連接註解
            //Switch.Diffrly開關的指令
            if (switchName.Contains("S1") && !switchName.Contains("S19"))
            {
                if(on_off == 1 && order == 1)//開
                {
                    WriteLine("//開啟DIFFRLY_" + switchName+ "的通道" + com1[inputnumber]+","+ com2[inputnumber] + "腳位: " +pin_name[inputnumber]);
                    WriteLine("Switch.DIFFRLY_" + switchName + "(" + com1[inputnumber] + "," + com2[inputnumber] + "," + on_off + ");");
                }
                else if(on_off == 0 && order == 1)//關
                {
                    WriteLine("//關閉DIFFRLY_" + switchName + "的通道" + com1[inputnumber] + "," + com2[inputnumber] + "腳位: " + pin_name[inputnumber]);
                    WriteLine("Switch.DIFFRLY_" + switchName + "(" + com1[inputnumber] + "," + com2[inputnumber] + "," + on_off + ");");
                }
                //測試用
                else if (on_off == 1 && order == 0)//開
                {
                    WriteLine("//開啟DIFFRLY_" + switchName + "的通道" + Final_OutputData[inputnumber,4] + "," + Final_OutputData[inputnumber, 5] + "量測點A: " + Final_OutputData[inputnumber, 1] + "量測點B" + Final_OutputData[inputnumber, 2]);
                    WriteLine("Switch.DIFFRLY_" + switchName + "(" + Final_OutputData[inputnumber, 4] + "," + Final_OutputData[inputnumber, 5] + "," + on_off + ");");
                }
                else if (on_off == 0 && order == 0)//關
                {
                    WriteLine("//關閉DIFFRLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 4] + "," + Final_OutputData[inputnumber, 5] + "量測點A: " + Final_OutputData[inputnumber, 1] + "量測點B" + Final_OutputData[inputnumber, 2]);
                    WriteLine("Switch.DIFFRLY_" + switchName + "(" + Final_OutputData[inputnumber, 4] + "," + Final_OutputData[inputnumber, 5] + "," + on_off + ");");
                }

            }
            else 
            {
                if (on_off == 1 && order == 1)//開啟 //修改開一個，一個com
                {
                    WriteLine("//開啟FORMARLY_" + switchName + "的通道" + com1[inputnumber] + "," + com2[inputnumber] + "腳位: " + pin_name[inputnumber]);
                    WriteLine("Switch.FORMARLY_" + switchName + "(" + com1[inputnumber] + on_off + ");");
                    WriteLine("");
                    WriteLine("Switch.FORMARLY_" + switchName + "(" + com2[inputnumber] + on_off + ");");
                }
                else if(on_off == 0 && order == 1)//關閉
                {
                    WriteLine("//關閉FORMARLY_" + switchName + "的通道" + com1[inputnumber] + "," + com2[inputnumber] + "腳位: " + pin_name[inputnumber]);
                    WriteLine("Switch.FORMARLY_" + switchName + "(" + com1[inputnumber] + on_off + ");");
                    WriteLine("");
                    WriteLine("Switch.FORMARLY_" + switchName + "(" + com2[inputnumber] + on_off + ");");
                }
                //測試用
                else if (on_off == 1 && order == 0)//開
                {
                    WriteLine("//開啟FORMARLY_" + switchName + "的通道" + Final_OutputData[inputnumber,4] + "," + Final_OutputData[inputnumber, 5] + "量測點A: " + Final_OutputData[inputnumber, 1] + "量測點B" + Final_OutputData[inputnumber, 2]);
                    WriteLine("Switch.FORMARLY_" + switchName + "(" + Final_OutputData[inputnumber, 4] + on_off + ");");
                    WriteLine("");
                    WriteLine("Switch.FORMARLY_" + switchName + "(" + Final_OutputData[inputnumber, 5] + on_off + ");");
                }
                else if (on_off == 0 && order == 0)//關
                {
                    WriteLine("//關閉FORMARLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 4] + "," + Final_OutputData[inputnumber, 5] + "量測點A: " + Final_OutputData[inputnumber, 1] + "量測點B" + Final_OutputData[inputnumber, 2]);
                    WriteLine("Switch.FORMARLY_" + switchName + "(" + Final_OutputData[inputnumber, 4] + on_off + ");");
                    WriteLine("");
                    WriteLine("Switch.FORMARLY_" + switchName + "(" + Final_OutputData[inputnumber, 5] + on_off + ");");
                }
            }
            WriteLine("");
        }
        //三相電
        public void AC_Source(int inputNumber, int on_off)
        {
            string[] Parameter = new string[MAX];
            Parameter = inputData[inputNumber].GetParameter;
            WriteLine("AC_Source_CPS6000_init()");
            //無法使用這些測量 excel錯誤
            switch (inputData[inputNumber].GetFunction)
            {
                case "AC_Source_CPS6000_init":showWrong(inputNumber, "AC_Source_CPS6000_init");
                    break;
                case "ShutDown":showWrong(inputNumber, "ShutDown");
                    break;
                case "AC_Source_CPS6000_OFF":showWrong(inputNumber, "AC_Source_CPS6000_OFF");
                    break;
                case "AC_Source_CPS6000_P5VAC_OFF":showWrong(inputNumber, "AC_Source_CPS6000_P5VAC_OFF");
                    break;
            }
            //打開AC_Source 送三相電
            if (on_off == 1)
            {
                if (inputData[inputNumber].GetFunction.Equals("AC_Source_CPS6000_ON"))
                {
                    //三個參數
                    WriteLine("//打開AC_Source 送三相電,參數: " + Parameter[0] + "," + Parameter[1] + "," + Parameter[2]);
                    WriteLine("AC_Source_CPS6000_ON(" + Parameter[0] + "," + Parameter[1] + "," + Parameter[2] + ");");
                }
                else if(inputData[inputNumber].GetFunction.Equals("AC_Source_CPS6000_P5VAC_ON"))
                {
                    WriteLine("//打開AC_Source 送5v三相電");
                    WriteLine("AC_Source_CPS6000_P5VAC_ON();");
                }
                else if(inputData[inputNumber].GetFunction.Equals("AC_Source_CPS6000_P115VAC_ON"))
                {
                    WriteLine("//打開AC_Source 送115v三相電");
                    WriteLine("AC_Source_CPS6000_P115VAC_ON();");
                }
                WriteLine("");
            }
            //關閉AC_Source 
            else if (on_off == 0)
            {
                WriteLine("//關閉AC_Source");
                if (inputData[inputNumber].GetFunction.Equals("AC_Source_CPS6000_ON"))
                {
                    WriteLine("AC_Source_CPS6000_OFF();");
                }
                else if (inputData[inputNumber].GetFunction.Equals("AC_Source_CPS6000_P5VAC_ON"))
                {
                    WriteLine("AC_Source_CPS6000_P5VAC_OFF();");
                }
                else if (inputData[inputNumber].GetFunction.Equals("AC_Source_CPS6000_P115VAC_ON"))
                {
                    WriteLine("AC_Source_CPS6000_OFF();");
                }
                WriteLine("");
            }
        }
        //input錯誤時呼叫
        public void showWrong(int inputNumber, string function)
        {
            MessageBox.Show("第"+inputNumber.ToString()+"個輸入的function: "+function+"錯誤");
        }
        //任意波形產生器
        public void AWG(int inputNumber, int on_off)
        {
            string s;
            string[] Parameter = new string[MAX];
            Parameter = inputData[inputNumber].GetParameter;
            WriteLine("AWG_init()");
            switch (inputData[inputNumber].GetFunction)
            {
                case "AWG_init":
                    showWrong(inputNumber, "AWG_init");
                    break;
            }
            if (on_off == 1)
            {
                if (inputData[inputNumber].Equals("AWG_SinWave"))
                {
                    //五個參數
                    WriteLine("//sin波型產生器 五個參數");
                    WriteLine("AWG_SinWave(" + Parameter[0] + "," + Parameter[1] + "," + Parameter[2] + "," + Parameter[3] + "," + Parameter[4] + ");");
                }
                else if (inputData[inputNumber].Equals("AWG_SquareWave"))
                {
                    //六個參數
                    WriteLine("//方波波型產生器 六個參數");
                    WriteLine("AWG_SquareWave(" + Parameter[0] + "," + Parameter[1] + "," + Parameter[2] + "," + Parameter[3] + "," + Parameter[4] + "," + Parameter[5] + ");");
                }
                else if (inputData[inputNumber].Equals("AWG_Pulse"))
                {
                    //8個參數
                    WriteLine("//AWG_Pulse波型產生器 8個參數");
                    WriteLine("AWG_Pulse(" + Parameter[0] + "," + Parameter[1] + "," + Parameter[2] + "," + Parameter[3] + "," + Parameter[4] + "," + Parameter[5] + "," + Parameter[6] + "," + Parameter[7] + ");");
                }
                //2個參數
                else if (inputData[inputNumber].Equals("AWG_DCoffset"))
                {
                    WriteLine("//");
                    WriteLine("AWG_DCoffset(" + Parameter[0] + "," + Parameter[1] + ");");
                }
            }
            else if (on_off == 0)
            {
                WriteLine("//關閉波型產生器");
                WriteLine("AWG_Close();");
            }
        }
        //直流電
        public void DCV(int inputNumber, int on_off)
        {
            string[] Parameter = new string[MAX];
            Parameter = inputData[inputNumber].GetParameter;
            switch (inputData[inputNumber].GetFunction)
            {
                //N6744B 直流電源模組
                case "N67xx_1_init":
                    WriteLine("N67xx_1_init();");
                    break;
                case "N67xx_2_init":
                    WriteLine("N67xx_2_init();");
                    break;
                //N5771A 電源供應器
                case "N57xx_init":
                    WriteLine("N57xx_init();");
                    break;
            }
            if(on_off == 1 && inputData[inputNumber].GetFunction.Equals("DCV_ON"))
            {
                //3參數
                WriteLine("//送直流電 參數:"+ Parameter[0] + ", " + Parameter[1] + ", " + Parameter[2]);
                WriteLine("DCV_ON(" + Parameter[0] + "," + Parameter[1] + "," + Parameter[2] + ");");
            }
            else if(on_off == 0)
            {
                //1參數
                WriteLine("//停止送直流電 參數:"+Parameter[0]);
                WriteLine("DCV_OFF(" + Parameter[0] + ");");
            }
        }
        //數位電錶
        public void DMM(int inputNumber)
        {
            string[] Parameter = new string[MAX];
            Parameter = inputData[inputNumber].GetParameter;
            switch (inputData[inputNumber].GetFunction)
            {
                // function不應被excel呼叫
                case "DMMhp34461A_init":
                    showWrong(inputNumber, "DMMhp34461A_init");
                    break;
                case "DMMhp34461A_Meas_ID_Res":
                    showWrong(inputNumber, "DMMhp34461A_Meas_ID_Res");
                    break;
                    //重設
                case "DMMhp34461A_Reset":
                    showWrong(inputNumber, "DMMhp34461A_Reset");
                    break;
            }
            
                //WriteLine("DMMhp34461A_init()");
                //1參數
                if (inputData[inputNumber].GetFunction.Equals("DMMhp34461A_meas") || Parameter[1] == null)
                {
                    WriteLine("//使用數位電錶量測 參數: " + Parameter[0]);
                    WriteLine("UUT_MainFunction.g_objMeasure_Data = DMMhp34461A_meas(" + Parameter[0] + ");");
                }
                //二參數
                else if (inputData[inputNumber].GetFunction.Equals("DMMhp34461A_meas") || Parameter[2] != null)
                {
                    WriteLine("//使用數位電錶量測 參數: " + Parameter[0] + "," + Parameter[1]);
                    WriteLine("UUT_MainFunction.g_objMeasure_Data = DMMhp34461A_meas(" + Parameter[0] + "," + Parameter[1] + ");");
                }
                WriteLine("");
            
        }
        //精密型LCR錶
        public void LCR(int inputNumber)
        {
            string[] Parameter = new string[MAX];
            Parameter = inputData[inputNumber].GetParameter;
            switch (inputData[inputNumber].GetFunction)
            {
                case "LCR_init":showWrong(inputNumber, "LCR_init");
                    break;
            }
                //開LCR
                //WriteLine("LCR_init()");
                //二參數
                if (inputData[inputNumber].GetFunction.Equals("LCR_meas") || Parameter[2] == null)
                {
                    WriteLine("//使用精密型LCR錶量測 參數: " + Parameter[0] + "," + Parameter[1]);
                    WriteLine("UUT_MainFunction.g_objMeasure_Data = LCR_meas(" + Parameter[0] + "," + Parameter[1] + ");");
                }
                else if (inputData[inputNumber].GetFunction.Equals("LCR_meas") || Parameter[2] != null)
                {
                    //3參數
                    WriteLine("//使用精密型LCR錶量測 參數: " + Parameter[0] + "," + Parameter[1] + Parameter[2]);
                    WriteLine("UUT_MainFunction.g_objMeasure_Data = LCR_meas(" + Parameter[0] + "," + Parameter[1] + "," + Parameter[2] + ");");
                }
                WriteLine("");
        }
    }
    
 
 






}
