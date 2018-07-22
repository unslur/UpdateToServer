
using SQLiteSugar;
using System;

namespace UpdateToServer
{
    /// <summary>
    /// SqlSugar
    /// </summary>
    public class SugarDao
    {
        private SugarDao()
        {

        }
        public static string ConnectionString
        {
            get
            {
                string reval = "DataSource=" + Environment.CurrentDirectory.ToString() + "\\db.sqlite3"; ; //这里可以动态根据cookies或session实现多库切换
                return reval;
            }
        }
        public static SqlSugarClient GetInstance()
        {
            var db = new SqlSugarClient(ConnectionString);
            db.IsEnableLogEvent = true;//Enable log events
            db.LogEventStarting = (sql, par) => { Console.WriteLine(sql + " " + par + "\r\n"); };
            return db;
        }
    }
}
