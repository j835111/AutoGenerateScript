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
                OleDbCommand command = new OleDbCommand("Select * from [" + row["TABLE_NAME"].ToString() + "]", exconn);
                OleDbDataReader reader = command.ExecuteReader();
                dataconent.Load(reader);
                if (row["TABLE_NAME"].ToString().Contains("chapter"))
                    if (row["TABLE_NAME"].ToString().Contains("test"))
                    {
                        CommandQuery.CreateTable(row["TABLE_NAME"].ToString().TrimEnd('$'), data.Chapter_test);
                        CommandQuery.InsertChapterTest(row["TABLE_NAME"].ToString().TrimEnd('$'), CommandQuery.DataTableToList(dataconent));
                    }
                    else
                    {
                        CommandQuery.CreateTable(row["TABLE_NAME"].ToString().TrimEnd('$'), data.Chapter);
                        CommandQuery.InsertChapter(row["TABLE_NAME"].ToString().TrimEnd('$'), dataconent);
                    }
                else if (row["TABLE_NAME"].ToString().Contains("component"))
                    CommandQuery.CreateTable(row["TABLE_NAME"].ToString().TrimEnd('$'), data.Component);
                else
                {
                    CommandQuery.CreateTable(row["TABLE_NAME"].ToString().TrimEnd('$'), data.Graph);

                }
                    

            }
        }
        private DataTable OpenExcel(string path)
        {
            DataTable columname = new DataTable();
            OpenFileDialog filedialog = new OpenFileDialog();
            filedialog.Title = "選取excel資料表";
            filedialog.Filter = "excel測試表 | *.xlsx";
            if (filedialog.ShowDialog() == DialogResult.OK)
            {
                path = filedialog.FileName;

                columname =  Form3.opene_excel(path);
            }
            return columname;
        }
        //private int TableNameToInt(string name)
        //{
        //    int x;
        //    switch (name)
        //    {
        //        case "Chaper_test":
        //            return 0;
        //        case "Chapter"


        //    }
        //}
    }
    class CommandQuery
    {
        private static OleDbConnection acconn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=board.accdb");

        private static void DataBaseExecute(StringBuilder commandtext)
        {
            OleDbCommand command = new OleDbCommand(commandtext.ToString(), acconn);
            if (!acconn.State.ToString().Contains("Open"))
                acconn.Open();

            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
        }
        private static void DataBaseExecute(string commandtext)
        {
            OleDbCommand command = new OleDbCommand(commandtext, acconn);
            if (!acconn.State.ToString().Contains("Open"))
                acconn.Open();

            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
        }
        private static void CreateTable(string chaptername, DataColumnCollection columname)
        {
            StringBuilder text = new StringBuilder("CREATE TABLE " + chaptername + "(" + columname[0] + " char(20)", 50);
            for (int i = 1; i < columname.Count; i++)
                text.Append(", " + columname[i] + " char(20)");

            text.Append(");");
            DataBaseExecute(text);
        }
        public static void CreateTable(string chaptername, List<string> column)
        {
            StringBuilder text = new StringBuilder("CREATE TABLE " + chaptername + "(" + column[0] + " char(20)", 50);
            for (int i = 1; i < column.Count; i++)
                text.Append(", " + column[i] + " char(20)");

            text.Append(");");
            DataBaseExecute(text);
        }
        private static void InsertContent(string chaptername, int columcount, DataRow data)
        {
            StringBuilder text = new StringBuilder("INSERT INTO " + chaptername + " VALUES('" + data[0] + "'", 100);
            for (int i = 1; i < columcount; i++)
                text.Append(", '" + data[i] + "'");

            text.Append(");");
            DataBaseExecute(text);
        }
        public static void InsertChapterTest(string chaptername, List<string> list)
        {
            int x = 1;
            foreach (string i in list)
            {
                string text = "INSERT INTO " + chaptername + " VALUES('','" + x + "','" + i + "','" + x.ToString("000000") + "')";
                DataBaseExecute(text);
                x++;
            }
            x = 1;
        }
        public static void InsertChapter(string chaptername, DataTable data)
        {
            foreach (DataRow row in data.Rows)
                InsertContent(chaptername, data.Rows.Count, row);
        }
        public static void InsertGraph(string name)
        {
            string text1 = "INSERT INTO graph VALUES('','1'," + name + ".png,'','','參考資料')";
            DataBaseExecute(text1);
            string text2 = "INSERT INTO graph VALUES('','2'," + name + "_PreProbe.png,'','','參考資料')";
            DataBaseExecute(text2);
        }
        public static List<string> DataTableToList(DataTable data)
        {
            List<string> list = new List<string>();
            foreach (DataRow row in data.Rows)
                list.Add(row[0].ToString());

            return list;
        }

        private static void CreateChapterTest(string chaptername, DataTable data)
        {
            CreateTable(chaptername, data.Columns);
            foreach (DataRow row in data.Rows)
                InsertContent(chaptername, data.Columns.Count, row);
        }
        public static void CloseConnect()
        {
            if (!acconn.State.ToString().Contains("Close"))
                acconn.Close();
            acconn.Dispose();
        }
    }
}
