using System;
using System.Windows.Forms;
using System.IO;


namespace WindowsFormsApplication5
{
    class DynamicWriteCS
    {
        //最大行數
        int MAXline = 20000;
        //記錄總行數
        int lineNumber;
        //紀錄排版數
        int spaceNumber;
        //紀錄測試的數量
        int testNumber;
        //紀錄此次要寫的章節
        int chapter;
        //所有測試程式碼都放這 
        string[] line;
        //第幾個測試章節
        int test;
        /*------------------------------*/
        int MAX = 200;
        InputData[] inputData;
        //每個測試的輸入資料
        string[,] each_InputNum;
        //存放輸出資料
        string[,] Final_OutputData;
        //Boolean[,] check;

        string[] DmmMmodeName = {"DC Voltage","DC Voltage Ratio",
                                 "AC Voltage","DC Current",
                                 "AC Current","2 Wire Resistance",
                                 "4 Wire Resistance","Continuity",
                                 "Diode"};
        //動態
        public DynamicWriteCS(int chapter, int Testnumber, InputData[] inputData, string[,] each_InputNum, string[,] Final_OutputData)
        {
            //初始化
            init();
            this.testNumber = Testnumber;
            this.inputData = inputData;
            this.each_InputNum = each_InputNum;
            this.Final_OutputData = Final_OutputData;
            this.chapter = chapter;
            Dynamic_makeFile();
        }
        void init()
        {
            each_InputNum = new string[MAX, MAX];
            Final_OutputData = new string[MAX, MAX];
            line = new string[MAXline];
            lineNumber = 0;
            spaceNumber = 0;
            test = 0;
            inputData = new InputData[100];
            //check = new Boolean[100, 100];
            
            
        }
        //用來寫入每一行的程式碼，包含排版
        public void WriteLine(String text)
        {
            if (lineNumber == MAXline)
            {
                MessageBox.Show("超過最大行數!!!");
                return;
            }
            //line本身初始化
            line[lineNumber] = " ";
            //排版用
            if (text != "" && String.Compare(text, "}") == 0)
                spaceNumber--;
            //排版用
            for (int i = 0; i < spaceNumber; i++)
                line[lineNumber] += "    ";
            //填入實際的內容
            line[lineNumber++] += text;
            //排版用
            if (text != "" && String.Compare(text, "{") == 0)
                spaceNumber++;
        }
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
                    writer.WriteLine(line[i-1]);
                }
                writer.Dispose();
                writer.Close();
                MessageBox.Show("Done!");
                Application.Exit();
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
            this.test = test;
            //第test節的註解
            //註解量測目的
            WriteLine("//量測目的: " + Final_OutputData[test + 1, 9]);
            //名稱
            WriteLine("public static void test" + chapter.ToString() + "_" + (test + 1).ToString() + "()");
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
            bool check=true;
            //第test個測試使用哪些輸入 這些輸入用了哪些function
            //單個測試
            for (int i = 0; each_InputNum[test, i] != null; i++)//0 1 2
            {
                //打開
                //輸入編號 
                int inputNumber = StringToInt(each_InputNum[test, i]) - 1;//因為輸入編號從0開始存，所以要減1
                //單個輸入開的多個switch
                //1.先開啟要送電的switch和通道
                for (int k = 0; !string.IsNullOrEmpty(inputData[inputNumber].GetRelay[k]); k++)//s[0]--s[1]---
                {
                    int checkNum=0;
                    //連接哪個Switch 
                    //先判斷是否會重複開啟
                    for (int j = 0; j<i; j++)//判斷測試編號: 0 1
                    {
                        for (; inputData[checkNum].GetNumber != each_InputNum[test, j]; checkNum++) ;//0 1 * 1 2
                        if (inputData[checkNum].GetNumber.Equals(each_InputNum[test, j]))
                        {
                            for (int x=0; inputData[checkNum].GetCom1[x] != null;x++)
                            {
                                if (inputData[checkNum].GetFunction.Equals(inputData[inputNumber].GetFunction) && inputData[checkNum].GetRelay[x].Equals(inputData[inputNumber].GetRelay[k]) && inputData[checkNum].GetCom1[x].Equals(inputData[inputNumber].GetCom1[k]))
                                {
                                    check = false;
                                }
                            }
                        }
                    }
                    if (check == true)
                    {
                        Dynamic_WriteSwitch(inputData[inputNumber].GetRelay[k], inputNumber, k, 1, 1);//order=1 on_off=1 k=讀到輸入的第幾個com
                    }
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
            }
            WriteLine(" //*************** Measure ***********************************");
            WriteLine("");

            //3.開啟要量測的switch腳位
            Dynamic_WriteSwitch(Final_OutputData[test + 1, 3], test + 1, 1, 2, 1);//order=2 on_off=0 k=1 :開第一個switch(量測點A) 
            if (!Final_OutputData[test+1, 6].Equals("null"))
            {
                Dynamic_WriteSwitch(Final_OutputData[test + 1, 6], test + 1, 2, 2, 1);//k=2 :開第二個switch(量測點B) 
            } 

            /*--------------------------- Measure --------------------------------*/
            //4.量測
            for (int i = 0; each_InputNum[test, i] != null; i++)
            {
                //輸入編號 
                int inputNumber = StringToInt(each_InputNum[test, i]) - 1;//因為輸入編號從0開始存，所以要減1
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

            //5.關閉要測的switch和通道
            Dynamic_WriteSwitch(Final_OutputData[test + 1, 3], test + 1, 1, 2, 0);//k=1 :關第一個switch(量測點A) order=2 on_off=0 
            if (!Final_OutputData[test + 1, 6].Equals("null"))
            {
                Dynamic_WriteSwitch(Final_OutputData[test + 1, 6], test + 1, 2, 2, 0);//k=2 :關第二個switch(量測點B) 
            }

            WriteLine(" //*************** Measure ***********************************");
            WriteLine("");
            /*--------------------------- close --------------------------------*/
            //關閉
            check = true;
            for (int i = 0; each_InputNum[test, i] != null; i++)
            {
                //輸入編號 
                int inputNumber = StringToInt(each_InputNum[test, i]) - 1;//因為輸入編號從0開始存，所以要減1
                //單個輸入開的多個switch
                string[] s = new string[MAX];
                s = inputData[inputNumber].GetRelay;

                //6.關閉電源
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

                //7.關閉送電的switch通道
                for (int k = 0; !string.IsNullOrEmpty(s[k]); k++)//s[0]--s[1]---
                {
                    //關閉哪個Switch 
                    int checkNum = 0;
                    //連接哪個Switch 
                    //先判斷是否會重複開啟
                    for (int j = 0; j < i; j++)//判斷測試編號: 0 1
                    {
                        for (; inputData[checkNum].GetNumber != each_InputNum[test, j]; checkNum++) ;//0 1 * 1 2
                        if (inputData[checkNum].GetNumber.Equals(each_InputNum[test, j]))
                        {
                            for (int x = 0; inputData[checkNum].GetCom1[x] != null; x++)
                            {
                                if (inputData[checkNum].GetFunction.Equals(inputData[inputNumber].GetFunction) && inputData[checkNum].GetRelay[x].Equals(inputData[inputNumber].GetRelay[k]) && inputData[checkNum].GetCom1[x].Equals(inputData[inputNumber].GetCom1[k]))
                                {
                                    check = false;
                                }
                            }
                        }
                    }
                    if (check == true)
                    {
                        Dynamic_WriteSwitch(s[k], inputNumber, k, 1, 0);//order=1 on_off=1
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
        public void Dynamic_WriteSwitch(string switchName, int inputnumber, int k, int order, int on_off)
        {
            //k: 哪個量測點
            //switchNumber: 第幾個switch
            //表示 Switch.FORMARLY開關的指令
            //Switch.Diffrly開關的指令
            //MessageBox.Show("switch=" + switchName + " inputnumber=" + inputnumber + " k=" + k);
            if (switchName.Contains("S1") && !switchName.Contains("S19"))
            {
                //送電用
                if (on_off == 1 && order == 1)//開
                {
                    WriteLine("//開啟DIFFRLY_" + switchName + "的通道" + inputData[inputnumber].GetCom1[k] + "," + inputData[inputnumber].GetCom2[k] + "腳位: " + inputData[inputnumber].GetPin_name[k]);
                    WriteLine("Switch.DIFFRLY_" + switchName + "(" + inputData[inputnumber].GetCom1[k] + "," + inputData[inputnumber].GetCom2[k] + "," + on_off + ");");
                }
                //送電用
                else if (on_off == 0 && order == 1)//關
                {
                    WriteLine("//關閉DIFFRLY_" + switchName + "的通道" + inputData[inputnumber].GetCom1[k] + "," + inputData[inputnumber].GetCom2[k] + "腳位: " + inputData[inputnumber].GetPin_name[k]);
                    WriteLine("Switch.DIFFRLY_" + switchName + "(" + inputData[inputnumber].GetCom1[k] + "," + inputData[inputnumber].GetCom2[k] + "," + on_off + ");");
                }
                //測試用
                else if (on_off == 1 && order == 2)//開
                {
                    if(k == 1)
                    {
                        WriteLine("//開啟DIFFRLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 4] + "," + Final_OutputData[inputnumber, 5] + "  量測點A: " + Final_OutputData[inputnumber, 1]);
                        WriteLine("Switch.DIFFRLY_" + switchName + "(" + Final_OutputData[inputnumber, 4] + "," + Final_OutputData[inputnumber, 5] + "," + on_off + ");");
                    }
                    else if(k == 2)
                    {
                        WriteLine("//開啟DIFFRLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 7] + "," + Final_OutputData[inputnumber, 8] + "  量測點B: " + Final_OutputData[inputnumber, 2]);
                        WriteLine("Switch.DIFFRLY_" + switchName + "(" + Final_OutputData[inputnumber, 7] + "," + Final_OutputData[inputnumber, 8] + "," + on_off + ");");
                    }
                    
                }
                //測試用
                else if (on_off == 0 && order == 2)//關
                {
                    if(k == 1)
                    {
                        WriteLine("//關閉DIFFRLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 4] + "," + Final_OutputData[inputnumber, 5] + "  量測點A: " + Final_OutputData[inputnumber, 1]);
                        WriteLine("Switch.DIFFRLY_" + switchName + "(" + Final_OutputData[inputnumber, 4] + "," + Final_OutputData[inputnumber, 5] + "," + on_off + ");");
                    }
                    else if(k == 2)
                    {
                        WriteLine("//關閉DIFFRLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 4] + "," + Final_OutputData[inputnumber, 5] + "  量測點B: " + Final_OutputData[inputnumber, 2]);
                        WriteLine("Switch.DIFFRLY_" + switchName + "(" + Final_OutputData[inputnumber, 4] + "," + Final_OutputData[inputnumber, 5] + "," + on_off + ");");
                    }
                }
            }
            else if (switchName.Contains("S19") || switchName.Contains("S22") || switchName.Contains("S2425"))
            {
                //送電用
                if (on_off == 1 && order == 1)//開啟 
                {
                    WriteLine("//開啟FORMCRLY_" + switchName + "的通道" + inputData[inputnumber].GetCom1[k] + "腳位: " + inputData[inputnumber].GetPin_name[k]);
                    WriteLine("Switch.FORMCRLY_" + switchName + "(" + inputData[inputnumber].GetCom1[k] + "," + on_off + ");");
                }
                //送電用
                else if (on_off == 0 && order == 1)//關閉
                {
                    WriteLine("//關閉FORMCRLY_" + switchName + "的通道" + inputData[inputnumber].GetCom1[k] + "腳位: " + inputData[inputnumber].GetPin_name[k]);
                    WriteLine("Switch.FORMCRLY_" + switchName + "(" + inputData[inputnumber].GetCom1[k] + "," + on_off + ");");
                }
                //測試用
                else if (on_off == 1 && order == 2)//開
                {
                    if (k == 1)
                    {
                        WriteLine("//開啟FORMCRLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 4] + "  量測點A: " + Final_OutputData[inputnumber, 2]);
                        WriteLine("Switch.FORMCRLY_" + switchName + "(" + Final_OutputData[inputnumber, 4] + "," + on_off + ");");
                        WriteLine("");
                    }
                    else if (k == 2)
                    {
                        WriteLine("//開啟FORMCRLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 7] + "  量測點B: " + Final_OutputData[inputnumber, 2]);
                        WriteLine("Switch.FORMCRLY_" + switchName + "(" + Final_OutputData[inputnumber, 7] + "," + on_off + ");");
                    }

                }
                //測試用
                else if (on_off == 0 && order == 2)//關
                {
                    if (k == 1)
                    {
                        WriteLine("//關閉FORMCRLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 4] + "  量測點A: " + Final_OutputData[inputnumber, 1]);
                        WriteLine("Switch.FORMCRLY_" + switchName + "(" + Final_OutputData[inputnumber, 4] + "," + on_off + ");");
                    }
                    else if (k == 2)
                    {
                        WriteLine("//關閉FORMCRLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 7] + "  量測點B:  " + Final_OutputData[inputnumber, 2]);
                        WriteLine("Switch.FORMCRLY_" + switchName + "(" + Final_OutputData[inputnumber, 7] + "," + on_off + ");");
                    }
                }
            }
            else
            {
                //送電用
                if (on_off == 1 && order == 1)//開啟 
                {
                    WriteLine("//開啟FORMARLY_" + switchName + "的通道" + inputData[inputnumber].GetCom1[k] + "腳位: " + inputData[inputnumber].GetPin_name[k]);
                    WriteLine("Switch.FORMARLY_" + switchName + "(" + inputData[inputnumber].GetCom1[k] + "," + on_off + ");");
                }
                //送電用
                else if (on_off == 0 && order == 1)//關閉
                {
                    WriteLine("//關閉FORMARLY_" + switchName + "的通道" + inputData[inputnumber].GetCom1[k] + "腳位: " + inputData[inputnumber].GetPin_name[k]);
                    WriteLine("Switch.FORMARLY_" + switchName + "(" + inputData[inputnumber].GetCom1[k] + "," + on_off + ");");
                }
                //測試用
                else if (on_off == 1 && order == 2)//開
                {
                    if(k == 1)
                    {
                        WriteLine("//開啟FORMARLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 4] + "  量測點A: " + Final_OutputData[inputnumber, 2]);
                        WriteLine("Switch.FORMARLY_" + switchName + "(" + Final_OutputData[inputnumber, 4] + "," + on_off + ");");
                        WriteLine("");
                    }
                    else if(k == 2)
                    {
                        WriteLine("//開啟FORMARLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 7] + "  量測點B: " + Final_OutputData[inputnumber, 2]);
                        WriteLine("Switch.FORMARLY_" + switchName + "(" + Final_OutputData[inputnumber, 7] + "," + on_off + ");");
                    }
                }
                else if (on_off == 0 && order == 2)//關
                {
                    if(k == 1)
                    {
                        WriteLine("//關閉FORMARLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 4] + "  量測點A: " + Final_OutputData[inputnumber, 1]);
                        WriteLine("Switch.FORMARLY_" + switchName + "(" + Final_OutputData[inputnumber, 4] + "," + on_off + ");");
                    }
                   else if(k == 2)
                    {
                        WriteLine("//關閉FORMARLY_" + switchName + "的通道" + Final_OutputData[inputnumber, 7] + "  量測點B:  " + Final_OutputData[inputnumber, 2]);
                        WriteLine("Switch.FORMARLY_" + switchName + "(" + Final_OutputData[inputnumber, 7] + "," + on_off + ");");
                    }
                }
            }
            WriteLine("");
        }
        //三相電
        public void AC_Source(int inputNumber, int on_off)
        {
            string[] Parameter = new string[MAX];
            Parameter = inputData[inputNumber].GetParameter;
            //無法使用這些測量 excel錯誤
            switch (inputData[inputNumber].GetFunction)
            {
                case "AC_Source_CPS6000_init":
                    showWrong(inputNumber, "AC_Source.AC_Source_CPS6000_init");
                    break;
                case "ShutDown":
                    showWrong(inputNumber, "AC_Source.ShutDown");
                    break;
                case "AC_Source_CPS6000_OFF":
                    showWrong(inputNumber, "AC_Source.AC_Source_CPS6000_OFF");
                    break;
                case "AC_Source_CPS6000_P5VAC_OFF":
                    showWrong(inputNumber, "AC_Source.AC_Source_CPS6000_P5VAC_OFF");
                    break;
            }
            //打開AC_Source 送三相電
            if (on_off == 1)
            {
                if (inputData[inputNumber].GetFunction.Equals("AC_Source_CPS6000_ON"))
                {
                    //三個參數
                    WriteLine("//打開AC_Source 送三相電,參數: " + Parameter[0] + "," + Parameter[1] + "," + Parameter[2]);
                    WriteLine("AC_Source.AC_Source_CPS6000_ON(" + Parameter[0] + "," + Parameter[1] + "," + Parameter[2] + ");");
                }
                else if (inputData[inputNumber].GetFunction.Equals("AC_Source_CPS6000_P5VAC_ON"))
                {
                    WriteLine("//打開AC_Source 送5v三相電");
                    WriteLine("AC_Source.AC_Source_CPS6000_P5VAC_ON();");
                }
                else if (inputData[inputNumber].GetFunction.Equals("AC_Source_CPS6000_P115VAC_ON"))
                {
                    WriteLine("//打開AC_Source 送115v三相電");
                    WriteLine("AC_Source.AC_Source_CPS6000_P115VAC_ON();");
                }
                WriteLine("");
            }
            //關閉AC_Source 
            else if (on_off == 0)
            {
                WriteLine("//關閉AC_Source");
                if (inputData[inputNumber].GetFunction.Equals("AC_Source_CPS6000_ON"))
                {
                    WriteLine("AC_Source.AC_Source_CPS6000_OFF();");
                }
                else if (inputData[inputNumber].GetFunction.Equals("AC_Source_CPS6000_P5VAC_ON"))
                {
                    WriteLine("AC_Source.AC_Source_CPS6000_P5VAC_OFF();");
                }
                else if (inputData[inputNumber].GetFunction.Equals("AC_Source_CPS6000_P115VAC_ON"))
                {
                    WriteLine("AC_Source.AC_Source_CPS6000_OFF();");
                }
                WriteLine("");
            }
        }
        //input錯誤時呼叫
        public void showWrong(int inputNumber, string function)
        {
            MessageBox.Show("第" + inputNumber.ToString() + "個輸入的function: " + function + "錯誤");
        }
        //任意波形產生器
        public void AWG(int inputNumber, int on_off)
        {
            string[] Parameter = new string[MAX];
            Parameter = inputData[inputNumber].GetParameter;
            switch (inputData[inputNumber].GetFunction)
            {
                case "AWG_init":
                    showWrong(inputNumber, "AWG.AWG_init");
                    break;
            }
            if (on_off == 1)
            {
                if (inputData[inputNumber].Equals("AWG_SinWave"))
                {
                    //五個參數
                    WriteLine("//sin波型產生器 五個參數");
                    WriteLine("AWG.AWG_SinWave(" + Parameter[0] + "," + Parameter[1] + "," + Parameter[2] + "," + Parameter[3] + "," + Parameter[4] + ");");
                }
                else if (inputData[inputNumber].Equals("AWG_SquareWave"))
                {
                    //六個參數
                    WriteLine("//方波波型產生器 六個參數");
                    WriteLine("AWG.AWG_SquareWave(" + Parameter[0] + "," + Parameter[1] + "," + Parameter[2] + "," + Parameter[3] + "," + Parameter[4] + "," + Parameter[5] + ");");
                }
                else if (inputData[inputNumber].Equals("AWG_Pulse"))
                {
                    //8個參數
                    WriteLine("//AWG_Pulse波型產生器 8個參數");
                    WriteLine("AWG.AWG_Pulse(" + Parameter[0] + "," + Parameter[1] + "," + Parameter[2] + "," + Parameter[3] + "," + Parameter[4] + "," + Parameter[5] + "," + Parameter[6] + "," + Parameter[7] + ");");
                }
                //2個參數
                else if (inputData[inputNumber].Equals("AWG_DCoffset"))
                {
                    WriteLine("//");
                    WriteLine("AWG.AWG_DCoffset(" + Parameter[0] + "," + Parameter[1] + ");");
                }
            }
            else if (on_off == 0)
            {
                WriteLine("//關閉波型產生器");
                WriteLine("AWG.AWG_Close();");
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
                    showWrong(inputNumber, "DCV.N67xx_1_init");
                    break;
                case "N67xx_2_init":
                    showWrong(inputNumber, "DCV.N67xx_2_init");
                    break;
                //N5771A 電源供應器
                case "N57xx_init":
                    showWrong(inputNumber, "DCV.N57xx_init");
                    break;
            }
            if (on_off == 1 && inputData[inputNumber].GetFunction.Equals("DCV_ON"))
            {
                //3參數
                WriteLine("//送直流電 參數:" + Parameter[0] + "," + Parameter[1] + "," + Parameter[2]);
                WriteLine("DCV.DCV_ON(" + Parameter[0] + "," + Parameter[1] + "," + Parameter[2] + ");");
            }
            else if (on_off == 0)
            {
                //1參數
                WriteLine("//停止送直流電 參數:" + Parameter[0]);
                WriteLine("DCV.DCV_OFF(" + Parameter[0] + ");");
            }
            WriteLine("");
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
                    showWrong(inputNumber, "DMM.DMMhp34461A_init");
                    break;
                case "DMMhp34461A_Meas_ID_Res":
                    showWrong(inputNumber, "DMM.DMMhp34461A_Meas_ID_Res");
                    break;
                //重設
                case "DMMhp34461A_Reset":
                    showWrong(inputNumber, "DMM.DMMhp34461A_Reset");
                    break;
            }

            //WriteLine("DMMhp34461A_init()");
            //1參數
            if (inputData[inputNumber].GetFunction.Equals("DMMhp34461A_meas") || Parameter[1] == null)
            {
                WriteLine("//使用數位電錶量測 參數: " + Parameter[0] + " (" + DmmMmodeName[StringToInt(Parameter[0]) - 1] + ")");
                WriteLine("UUT_MainFunction.g_objMeasure_Data = DMM.DMMhp34461A_meas(" + Parameter[0] + ");");
            }
            //二參數
            else if (inputData[inputNumber].GetFunction.Equals("DMMhp34461A_meas") || Parameter[2] != null)
            {
                WriteLine("//使用數位電錶量測 參數: " + Parameter[0] + "," + Parameter[1]);
                WriteLine("UUT_MainFunction.g_objMeasure_Data = DMM.DMMhp34461A_meas(" + Parameter[0] + "," + Parameter[1] + ");");
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
                case "LCR_init":
                    showWrong(inputNumber, "LCR_init");
                    break;
            }
            //開LCR
            //WriteLine("LCR_init()");
            //二參數
            if (inputData[inputNumber].GetFunction.Equals("LCR_meas") || Parameter[2] == null)
            {
                WriteLine("//使用精密型LCR錶量測 參數: " + Parameter[0] + "," + Parameter[1]);
                WriteLine("UUT_MainFunction.g_objMeasure_Data = LCR.LCR_meas(" + Parameter[0] + "," + Parameter[1] + ");");
            }
            else if (inputData[inputNumber].GetFunction.Equals("LCR_meas") || Parameter[2] != null)
            {
                //3參數
                WriteLine("//使用精密型LCR錶量測 參數: " + Parameter[0] + "," + Parameter[1] + Parameter[2]);
                WriteLine("UUT_MainFunction.g_objMeasure_Data = LCR.LCR_meas(" + Parameter[0] + "," + Parameter[1] + "," + Parameter[2] + ");");
            }
            WriteLine("");
        }
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
    }
}

