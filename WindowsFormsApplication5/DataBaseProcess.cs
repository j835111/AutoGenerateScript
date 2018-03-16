using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication5
{
    class DataBaseProcess
    {
        DataColumnName data;
        string expath;

        DataTable dataconent = new DataTable();
        public DataBaseProcess(DataColumnName name, string path)
        {
            data = name;
            expath = path;
            OleDbConnection exconn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'");
            DataTable table = OpenExcel(path);
            exconn.Open();

            foreach (DataRow row in table.Rows)
            {
                OleDbCommand command = new OleDbCommand("SELECT * FROM [" + row["TABLE_NAME"].ToString() + "]", exconn);
                OleDbDataReader reader = command.ExecuteReader();
                dataconent.Load(reader);
                if (row["TABLE_NAME"].ToString().Contains("chapter"))
                {
                    if (row["TABLE_NAME"].ToString().Contains("test"))
                    {
                        CommandQuery.CreateTable(row["TABLE_NAME"].ToString().TrimEnd('$'), data.Chapter_test);
                        CommandQuery.InsertChapterTest(row["TABLE_NAME"].ToString().TrimEnd('$'), CommandQuery.DataTableToList(dataconent));
                    }
                    else
                    {
                        CommandQuery.CreateTable(row["TABLE_NAME"].ToString().TrimEnd('$'), dataconent.Columns);
                        CommandQuery.InsertChapter(row["TABLE_NAME"].ToString().TrimEnd('$'), dataconent);
                    }
                }

                dataconent.Clear();
                dataconent.Columns.Clear(); 
            }

            CommandQuery.CreateTable("component", data.Component);
            //CommandQuery.CreateTable("graph", data.Graph);
            //CommandQuery.InsertGraph();

            exconn.Close();
            CommandQuery.CloseConnect();
        }
        private DataTable OpenExcel(string path)
        {
            DataTable columname = new DataTable();
            columname = Form3.opene_excel(path);
            
            return columname;
        }
    }
    class CommandQuery
    {
        private static OleDbConnection acconn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=board.accdb");

        /// <summary>
        /// 執行資料庫指令
        /// </summary>
        /// <param name="commandtext">資料庫指令</param>
        private static void DataBaseExecute(StringBuilder commandtext)
        {
            OleDbCommand command = new OleDbCommand(commandtext.ToString(), acconn);
            if (!acconn.State.ToString().Contains("Open"))
                acconn.Open();

            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
        }

        /// <summary>
        /// 執行資料庫指令
        /// </summary>
        /// <param name="commandtext"></param>
        private static void DataBaseExecute(string commandtext)
        {
            OleDbCommand command = new OleDbCommand(commandtext, acconn);
            if (!acconn.State.ToString().Contains("Open"))
                acconn.Open();

            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
        }

        /// <summary>
        /// 建立資料表
        /// </summary>
        /// <param name="chaptername">章節名稱</param>
        /// <param name="columname">欄位名稱</param>
        public static void CreateTable(string chaptername, DataColumnCollection columname)
        {
            StringBuilder text = new StringBuilder("CREATE TABLE " + chaptername + "(" + columname[0] + " char(50)", 50);
            for (int i = 1; i < columname.Count; i++)
                text.Append(", " + columname[i] + " char(50)");

            text.Append(");");
            DataBaseExecute(text);
        }

        /// <summary>
        /// 建立資料表
        /// </summary>
        /// <param name="chaptername">章節名稱</param>
        /// <param name="column">欄位名稱</param>
        public static void CreateTable(string chaptername, List<string> column)
        {
            StringBuilder text = new StringBuilder("CREATE TABLE " + chaptername + "(" + column[0] + " char(50)", 50);
            for (int i = 1; i < column.Count; i++)
                text.Append(", " + column[i] + " char(50)");

            text.Append(");");
            DataBaseExecute(text);
        }

        /// <summary>
        /// 插入資料表內容
        /// </summary>
        /// <param name="chaptername">章節名稱</param>
        /// <param name="columcount">總共有幾欄</param>
        /// <param name="data">資料表內容</param>
        private static void InsertContent(string chaptername, int columcount, DataRow data)
        {
            StringBuilder text = new StringBuilder("INSERT INTO " + chaptername + " VALUES('" + data[0] + "'", 100);
            for (int i = 1; i < columcount; i++)
                text.Append(", '" + data[i] + "'");

            text.Append(");");
            try
            {
                DataBaseExecute(text);
            }
            catch
            {
                MessageBox.Show("請將欄位中的graph改為\"graph\"!!");
                CloseConnect();
                return;
            }
        }

        /// <summary>
        /// 插入Chapter_Test(專用)
        /// </summary>
        /// <param name="chaptername">章節名稱</param>
        /// <param name="list">內容</param>
        public static void InsertChapterTest(string chaptername, List<string> list)
        {
            int x = 1;
            foreach (string i in list)
            {
                string text = "INSERT INTO " + chaptername + " VALUES('','" + x + "','" + i + "','','" + x.ToString("000000") + "')";
                DataBaseExecute(text);
                x++;
            }
            x = 1;
        }

        /// <summary>
        /// 插入章節
        /// </summary>
        /// <param name="chaptername">章節名稱</param>
        /// <param name="data">內容</param>
        public static void InsertChapter(string chaptername, DataTable data)
        {
            foreach (DataRow row in data.Rows)
                InsertContent(chaptername, data.Columns.Count, row);
        }

        /// <summary>
        /// 插入Graph
        /// </summary>
        /// <param name="name">IC的名字</param>
        public static void InsertGraph(string name)
        {
            string text1 = "INSERT INTO graph VALUES('','1'," + name + ".png,'','','參考資料')";
            DataBaseExecute(text1);
            string text2 = "INSERT INTO graph VALUES('','2'," + name + "_PreProbe.png,'','','參考資料')";
            DataBaseExecute(text2);
        }

        /// <summary>
        /// DataTableToList方法
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static List<string> DataTableToList(DataTable data)
        {
            List<string> list = new List<string>();
            foreach (DataRow row in data.Rows)
                list.Add(row[0].ToString());

            return list;
        }

        /// <summary>
        /// 插入Chapter_Test(專用)
        /// </summary>
        /// <param name="chaptername">章節名稱</param>
        /// <param name="data">內容</param>
        private static void CreateChapterTest(string chaptername, DataTable data)
        {
            CreateTable(chaptername, data.Columns);
            foreach (DataRow row in data.Rows)
                InsertContent(chaptername, data.Columns.Count, row);
        }

        /// <summary>
        /// 關閉資料庫連接
        /// </summary>
        public static void CloseConnect()
        {
            if (!acconn.State.ToString().Contains("Close"))
                acconn.Close();
            acconn.Dispose();
        }
    }
}
