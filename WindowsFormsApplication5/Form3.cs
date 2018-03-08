using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;

namespace WindowsFormsApplication5
{

    public partial class Form3 : Form
    {
        string path;//紀錄excel路徑 
        int test, column = 0, max = 1000;
        string[,] data;//excel量測
        string[,] comment;//量測目的
        int ChapterNunber, Catagory;//章節 測試種類
        int testNumber, InputNumber,TestNumber;//靜態測試數量 動態輸入數量 動態測試數量
        string[,] outputData;//存動態輸出
        InputData[] inputData;
        string chapter = "chapter";

        //初始化
        public Form3()
        {
            data = new string[max, max];
            comment = new string[max, max];
            outputData = new string[max, max];
            inputData = new InputData[100];
            InitializeComponent();
            chap_choise.Visible = false;
            button2.Visible = false;
            chap_choise.Items.Add("請選擇章節");
            chap_choise.SelectedIndex = 0;
            TestCatagory.Items.Add("靜態測試");
            TestCatagory.Items.Add("動態測試");
        }

        

        //excel的開啟，當按開啟鍵時觸發
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog filedialog = new OpenFileDialog();
            filedialog.Title = "選取excel資料表";
            filedialog.Filter = "excel測試表 | *.xlsx";
            if (filedialog.ShowDialog() == DialogResult.OK)
            {
                path = filedialog.FileName;
                //當成功選取excel,text=1
                test = 1;
                opene_excel(path);
            }
        }
        //連接
        public void opene_excel(string path)
        {
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'");
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            conn.Open();
            DataTable Table = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            chap_choise.Items.Clear();
            chap_choise.Items.Add("請選擇章節");
            //章選擇的控制項項目新增
            /*---------------------------------------------------------*/
            foreach (DataRow row in Table.Rows)
            {
                string table_name = row["TABLE_NAME"].ToString();
                if (table_name.IndexOf(chapter, StringComparison.OrdinalIgnoreCase) >= 0 && !(table_name.ToString()).Contains("Database"))
                {
                    chap_choise.Items.Add(table_name.TrimEnd('$'));
                }
            }
            /*---------------------------------------------------------*/
            conn.Close();
            chap_choise.Visible = true;
            button2.Visible = true;
            TestCatagory.Visible = true;
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            
        }

        //選擇章節
        private void Chap_choise_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ChapterNunber = StringToInt(chap_choise.SelectedItem.ToString().TrimStart(chapter.ToCharArray()));

            }
            catch
            {
                return;
            }
        }
        
        private void Form3_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        //選擇靜態或動態
        private void TestCatagory_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Catagory:1為靜態，2為動態
            Catagory = TestCatagory.SelectedIndex;                       
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if(ChapterNunber == 0 )
            {
                MessageBox.Show("未選擇章節");
            }
            else if(Catagory == 0)
            {
                MessageBox.Show("未選擇測試種類");
            }
            else if(ChapterNunber == 0 && Catagory == 0)
            {
                MessageBox.Show("兩者皆未選擇");
            }
            else if(Catagory == 1)
            {
                static_test(ChapterNunber);
            }
            else if(Catagory == 2)
            {
                dynamic_test(ChapterNunber);
            }
            
        }
        //動態測試
        public void dynamic_test(int ChapterNunber)
        {
            string name = chap_choise.SelectedItem.ToString();
            int count=1;
            try
            {
                //開excel
                /*---------------------------------------------------------*/

                OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'");
                DataSet ds1 = new DataSet();
                DataTable dt1 = new DataTable();
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("Select * from [" + name + "$]", conn);
                da.Fill(ds1);
                dt1 = ds1.Tables[0]; 

                /*---------------------------------------------------------*/

                //判斷標題是否正確
                if (dt1.Rows[0][1].ToString() != "輸入編號" && dt1.Rows[0][2].ToString() != "儀器名稱" && dt1.Rows[0][3].ToString() != "使用函式" && dt1.Rows[0][4].ToString() != "參數" && dt1.Rows[0][5].ToString() != "腳位" && dt1.Rows[0][6].ToString() != "腳位連接的relay" && dt1.Rows[0][7].ToString() != "com1" && dt1.Rows[0][8].ToString() != "com2" && dt1.Rows[0][9].ToString() != "內容")
                {
                    MessageBox.Show("欄位標題錯誤");
                }
                //讀輸入資料
                for(int i=1,j=-1,k=0; !string.IsNullOrEmpty(dt1.Rows[i][3].ToString()); i++,k++)
                {
                    if(!string.IsNullOrEmpty(dt1.Rows[i][1].ToString()))
                    {
                        ++j;
                        //紀錄幾個輸入
                        inputData[j] = new InputData(Int32.Parse(dt1.Rows[i][1].ToString()));
                        inputData[j].SetData(dt1.Rows[i][2].ToString(), dt1.Rows[i][3].ToString(), dt1.Rows[i][9].ToString(), dt1.Rows[i][4].ToString());
                        k = 0;
                    }
                    inputData[j].SetData(k, dt1.Rows[i][5].ToString(), dt1.Rows[i][6].ToString(), dt1.Rows[i][7].ToString(),dt1.Rows[i][8].ToString());
                    count++;
                    InputNumber = j + 1;
                }
                //讀取輸出在excel的欄位
                while (!dt1.Rows[count][1].ToString().Equals("測試編號"))
                {
                    count++;
                }
                if(dt1.Rows[count][1].ToString()!= "測試編號"  && dt1.Rows[count][2].ToString() != "輸入編號" && dt1.Rows[count][3].ToString() != "量測點A" && dt1.Rows[count][4].ToString() != "量測點B" && dt1.Rows[count][5].ToString() != "量測點A連結的relay" && dt1.Rows[count][6].ToString()!= "COM1" && dt1.Rows[count][7].ToString() != "COM2" && dt1.Rows[count][8].ToString() != "量測點B連結的relay"  && dt1.Rows[count][9].ToString() !="COM1" && dt1.Rows[count][10].ToString() !="COM2" && dt1.Rows[count][11].ToString() !="量測目的")
                {
                    MessageBox.Show("標題錯誤");
                }
                
                //讀取輸出欄位
                for(int j=0; !string.IsNullOrEmpty(dt1.Rows[count][1].ToString()); count++,j++)
                {
                    for(int i=0;i<11;i++)
                    {
                        outputData[j, i] = dt1.Rows[count][i+1].ToString();
                    }
                    TestNumber = j;
                }

                conn.Close();

            }
            catch (Exception x)
            {
                MessageBox.Show("錯誤" + x.ToString());
            }
            Form4 Form4 = new Form4(ChapterNunber, InputNumber,inputData,TestNumber,outputData);
            Form4.Visible = true;
            this.Hide();



        }

        public void static_test(int ChapterNunber)
        {
            string name = chap_choise.SelectedItem.ToString();
            try
            {
                OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'");
                DataSet ds1 = new DataSet();
                DataTable dt1 = new DataTable();
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("Select * from [" + name + "$]", conn);
                da.Fill(ds1);
                dt1 = ds1.Tables[0];

                //寫入註解的說明到comment裡面
                /*---------------------------------------------------------*/
                if ((dt1.Rows[0][2].ToString()).Contains("量測點") && (dt1.Rows[0][3].ToString()).Contains("量測點") && (dt1.Rows[0][10].ToString()).Contains("量測目的"))
                {
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        comment[i, 0] = dt1.Rows[i][2].ToString();//紀錄量測的量測點A (在EXCEL上為欄位C)
                        comment[i, 1] = dt1.Rows[i][3].ToString();//紀錄量測的量測點B (在EXCEL上為欄位D)
                        comment[i, 2] = dt1.Rows[i][10].ToString();//紀錄量測目的 (在EXCEL上為欄位K)
                                                                  //label2.Text += comment[i, 0] + "  " + comment[i, 1] + " " + comment[i, 2];
                                                                  //label2.Text += "\n";
                        testNumber = i;
                    }
                }
                else
                    MessageBox.Show("標題錯誤");
                /*---------------------------------------------------------*/

                //寫入測試節的數目
                /*---------------------------------------------------------*/

                testNumber = StringToInt(dt1.Rows[0][0].ToString());

                /*---------------------------------------------------------*/


                //寫入COM和switch的資料到data
                /*---------------------------------------------------------*/

                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    //label2.Text += "\n";
                    //E欄位寫到data 0 欄位
                    //F欄位寫道data 1 欄位
                    //G欄位寫道data 2 欄位
                    //H欄位寫道data 3 欄位
                    //I欄位寫道data 4 欄位
                    //J欄位寫道data 5 欄位
                    for (int j = 4; j < 10; j++)
                    {
                        data[i, column] = dt1.Rows[i][j].ToString();
                        //label2.Text += data[i, column] + "  ";
                        column++;
                    }
                    //label2.Visible = false;
                    column = 0;
                }
                /*---------------------------------------------------------*/

                conn.Close();
            }
            catch (Exception x)
            {
                MessageBox.Show("錯誤" + x.ToString());
            }
            if (testNumber == 0)
            {
                MessageBox.Show("未輸入正確測試數量!請把測試數量寫在excel A2欄位");
            }
            else
            {
                //傳入form2
                Form2 Form2 = new Form2(testNumber, ChapterNunber, data, comment);
                this.Visible = false;
                Form2.Visible = true;
            }

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
