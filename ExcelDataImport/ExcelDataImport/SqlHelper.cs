using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace ExcelDataImport
{
    public static class SqlHelperExt
    {
        public static int AddRange(this IDataParameterCollection coll, IDataParameter[] par)
        {
            int i = 0;
            foreach (var item in par)
            {
                coll.Add(item);
                i++;
            }
            return i;
        }
    }

    public class SqlHelper
    {
        /// <summary>
        /// 创建数据库连接
        /// </summary>
        /// <returns></returns>
        public SqlConnection CreatSqlConnection(string connectionStr)
        {
            try
            {
                //XmlConfig xc = new XmlConfig();
                SqlConnection sqlConn = null;
                //string connectionStr = xc.GetXmlValue("ServerConfig.xml", "connectionStrings");
                sqlConn = new SqlConnection(connectionStr);

                sqlConn.Open();//打开数据库
                return sqlConn;
            }
            catch (Exception e)
            {
                throw new Exception("SQL Connection Error!");
            }
        }

        /// <summary>
        /// 关闭数据库连接
        /// </summary>
        /// <param name="sqlConn"></param>
        public void CloseSqlConnection(SqlConnection sqlConn)
        {
            sqlConn.Close();
        }

        /// <summary>
        /// 执行SQL查询语句
        /// </summary>
        /// <param name="sql">sql语句Str</param>
        /// <param name="connectionStr">数据库连接地址</param>
        /// <returns>Data表</returns>
        public DataTable ExecuteReader(string sql, string connectionStr)
        {
            DataTable dt = new DataTable();
            SqlConnection sqlConn = CreatSqlConnection(connectionStr);
            using (IDbCommand cmd = new SqlCommand(sql, (SqlConnection)sqlConn))
            {
                using (IDataReader dr = cmd.ExecuteReader())
                {
                    dt.Load(dr);
                }
            }
            sqlConn.Close();
            return dt;
        }


        /// <summary>
        /// 执行查询带参的sql语句
        /// </summary>
        /// <param name="sql">查询sql语句</param>
        /// <param name="par">sql语句中的参数</param>
        /// <returns>返回一个表集合</returns>
        public DataSet GetDataSet(string sql, IDataParameter[] par, string connectionStr)
        {
            DataSet ds = new DataSet();
            SqlConnection sqlConn = CreatSqlConnection(connectionStr);
            try
            {
                using (IDbCommand cmd = new SqlCommand(sql, (SqlConnection)sqlConn))
                {
                    cmd.Parameters.AddRange(par);
                    IDataAdapter da = new SqlDataAdapter((SqlCommand)cmd);
                    da.Fill(ds);
                    cmd.Parameters.Clear();

                }
                sqlConn.Close();
                return ds;
            }
            catch (Exception)
            {
                //cmd.Parameters.Clear();
                sqlConn.Close();
                return ds;
            }
        }

        /// <summary>
        /// 执行查询带不参的sql语句
        /// </summary>
        /// <param name="sql">查询sql语句</param>
        /// <param name="par">sql语句中的参数</param>
        /// <returns>返回一个表集合</returns>
        public DataSet GetDataSet(string sql, string connectionStr)
        {
            DataSet ds = new DataSet();
            SqlConnection sqlConn = CreatSqlConnection(connectionStr);
            try
            {
                using (IDbCommand cmd = new SqlCommand(sql, (SqlConnection)sqlConn))
                {
                    //cmd.Parameters.AddRange(par);
                    IDataAdapter da = new SqlDataAdapter((SqlCommand)cmd);
                    da.Fill(ds);
                    cmd.Parameters.Clear();

                }
                sqlConn.Close();
                return ds;
            }
            catch (Exception)
            {
                //cmd.Parameters.Clear();
                sqlConn.Close();
                return ds;
            }
        }



        /// <summary>
        /// 执行SQL增、删、改语句
        /// </summary>
        /// <param name="sql">sql语句Str</param>
        /// <param name="connectionStr">数据库连接地址</param>
        /// <returns>影响条数 int</returns>
        public int ExecuteNonQuery(string sql, string connectionStr)
        {
            int result = 0;
            SqlConnection sqlConn = CreatSqlConnection(connectionStr);
            using (IDbCommand cmd = new SqlCommand(sql, (SqlConnection)sqlConn))
            {
                result = cmd.ExecuteNonQuery();
            }
            sqlConn.Close();
            return result;
        }

        /// <summary>
        /// 执行带参SQL增、删、改语句
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="par"></param>
        /// <param name="connectionStr"></param>
        /// <returns></returns>
        public int ExecuteNonQuery(string sql, IDbDataParameter[] par, string connectionStr)
        {
            int result = 0;
            SqlConnection sqlConn = CreatSqlConnection(connectionStr);
            using (IDbCommand cmd = new SqlCommand(sql, (SqlConnection)sqlConn))
            {
                cmd.Parameters.AddRange(par);
                try
                {
                    result = cmd.ExecuteNonQuery();
                }
                catch { }
            }
            sqlConn.Close();
            return result;
        }







        public int ExecuteSqlBulkCopy(DataTable dataTable, string connectionStr, string DestinationTableName)
        {

            //SqlConnection sqlConn = new SqlConnection(connectionStr);
            //sqlConn.Open();//打开数据库
            SqlConnection sqlConn = CreatSqlConnection(connectionStr);
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConn))
            {
                bulkCopy.DestinationTableName = DestinationTableName;//目标表
                try
                {
                    bulkCopy.WriteToServer(dataTable);//写入
                }
                catch (Exception ex)
                {
                    //Console.WriteLine(ex.Message);
                }
            }
            sqlConn.Close();//关闭连接
            return dataTable.Rows.Count;
        }


        /// <summary>
        /// 执行带参SQL存储过程
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="par"></param>
        /// <param name="connectionStr"></param>
        /// <returns></returns>
        public int ExecuteStoredProcedure(string StoredProcedureName, IDbDataParameter[] par, string connectionStr)
        {
            int result = 0;
            SqlConnection sqlConn = CreatSqlConnection(connectionStr);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = sqlConn;
            cmd.CommandText = StoredProcedureName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddRange(par);
            try
            {
                result = cmd.ExecuteNonQuery();
                sqlConn.Close();
                return result;
            }
            catch
            {
                return -1;
            }
            //}

        }


        /// <summary>
        /// 执行批量查询语句并返回影响行数 add by shenwenhua 2012-7-21 执行批量插入操作
        /// </summary>
        /// <param name="sql">查询语句</param>
        /// <param name="sqlParameter">参数</param>
        /// <returns>影响行数</returns>
        public int ExecuteSqlStr(string sql, List<SqlParameter[]> sqlParameter, string ConnectionString)
        {
            int tempResault = 0;
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                SqlCommand cmd = new SqlCommand();
                try
                {
                    for (int i = 0; i < sqlParameter.Count; i++)
                    {
                        PrepareCommand(cmd, connection, null, CommandType.Text, sql, null);
                        cmd.Parameters.AddRange(sqlParameter[i]);
                        tempResault += cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                    }
                }
                catch (SqlException se)
                {
                    cmd.Parameters.Clear();
                    connection.Close();
                }

            }
            return tempResault;
        }

        /// <summary>
        /// 添加SqlParameter参数到SqlCommand
        /// </summary>
        /// <param name="cmd">SqlCommand命令对象</param>
        /// <param name="conn">SqlConnection连接对象</param>
        /// <param name="trans">SqlTransaction事务对象</param>
        /// <param name="cmdType">执行SqlCommand的类型</param>
        /// <param name="cmdText">执行SqlCommand的SQL语句</param>
        /// <param name="commandParameters">SQL参数</param>
        private void PrepareCommand(SqlCommand cmd, SqlConnection conn, SqlTransaction trans, CommandType cmdType, string cmdText, SqlParameter[] commandParameters)
        {
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            if (trans != null)
            {
                cmd.Transaction = trans;
            }
            cmd.CommandType = cmdType;
            if (commandParameters != null)
            {
                foreach (SqlParameter parm in commandParameters)
                {
                    cmd.Parameters.Add(parm);
                }
            }
        }




    }
}
