using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Web;
using System.Collections;

    /// <summary>
    /// OleDb Server数据库访问助手类
    /// 本类为静态类 不可以被实例化 需要使用时直接调用即可
    /// Copyright © 2013 Wedn.Net
    /// </summary>
    public static partial class OleDbHelper
    {
        private static readonly string[] localhost = new[] { "localhost", ".", "(local)" };
        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        private readonly static string connStr;

        static string mappath = HttpContext.Current.Server.MapPath("~/App_Data/jdkj.mdb");
        static string conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mappath;

        static OleDbHelper()
        {
            //var conn = System.Configuration.ConfigurationManager.ConnectionStrings["ConnStr"];
            if (conn!=null)
            {
                connStr = new OleDbConnection(conn).ConnectionString;
            }
        }

        #region 获取指定表中指定字段的最大值, 确保字段为INT类型
        /// <summary>
        /// 获取指定表中指定字段的最大值, 确保字段为INT类型
        /// </summary>
        /// <param name="fieldName">字段名</param>
        /// <param name="tableName">表名</param>
        /// <returns>最大值</returns>
        public static int QueryMaxId(string fieldName, string tableName)
        {
            string OleDb = string.Format("select max([{0}]) from [{1}];", fieldName, tableName);
            object res = ExecuteScalar(OleDb);
            if (res == null)
                return 0;
            return Convert.ToInt32(res);
        }

        #endregion

        #region OleDb执行方法

        #region ExecuteNonQuery +static int ExecuteNonQuery(string cmdText, params OleDbParameter[] parameters)
        /// <summary>
        /// 执行一个非查询的T-OleDb语句，返回受影响行数，如果执行的是非增、删、改操作，返回-1
        /// </summary>
        /// <param name="cmdText">要执行的T-OleDb语句</param>
        /// <param name="parameters">参数列表</param>
        /// <exception cref="链接数据库异常"></exception>
        /// <returns>受影响的行数</returns>
        public static int ExecuteNonQuery(string cmdText, params OleDbParameter[] parameters)
        {
            return ExecuteNonQuery(cmdText, CommandType.Text, parameters);
        }
        #endregion

        #region ExecuteNonQuery +static int ExecuteNonQuery(string cmdText, CommandType type, params OleDbParameter[] parameters)
        /// <summary>
        /// 执行一个非查询的T-OleDb语句，返回受影响行数，如果执行的是非增、删、改操作，返回-1
        /// </summary>
        /// <param name="cmdText">要执行的T-OleDb语句</param>
        /// <param name="type">命令类型</param>
        /// <param name="parameters">参数列表</param>
        /// <exception cref="链接数据库异常"></exception>
        /// <returns>受影响的行数</returns>
        public static int ExecuteNonQuery(string cmdText, CommandType type, params OleDbParameter[] parameters)
        {
            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
                using (OleDbCommand cmd = new OleDbCommand(cmdText, conn))
                {
                    if (parameters != null)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddRange(parameters);
                    }
                    cmd.CommandType = type;
                    try
                    {
                        conn.Open();
                        int res = cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                        return res;
                    }
                    catch (System.Data.OleDb.OleDbException e)
                    {
                        conn.Close();
                        throw e;
                    }
                }
            }
        }
        #endregion

        #region ExecuteScalar +static object ExecuteScalar(string cmdText, params OleDbParameter[] parameters)
        /// <summary>
        /// 执行一个查询的T-OleDb语句，返回第一行第一列的结果
        /// </summary>
        /// <param name="cmdText">要执行的T-OleDb语句</param>
        /// <param name="parameters">参数列表</param>
        /// <exception cref="链接数据库异常"></exception>
        /// <returns>返回第一行第一列的数据</returns>
        public static object ExecuteScalar(string cmdText, params OleDbParameter[] parameters)
        {
            return ExecuteScalar(cmdText, CommandType.Text, parameters);
        }
        #endregion

        #region ExecuteScalar +static object ExecuteScalar(string cmdText, CommandType type, params OleDbParameter[] parameters)
        /// <summary>
        /// 执行一个查询的T-OleDb语句，返回第一行第一列的结果
        /// </summary>
        /// <param name="cmdText">要执行的T-OleDb语句</param>
        /// <param name="type">命令类型</param>
        /// <param name="parameters">参数列表</param>
        /// <exception cref="链接数据库异常"></exception>
        /// <returns>返回第一行第一列的数据</returns>
        public static object ExecuteScalar(string cmdText, CommandType type, params OleDbParameter[] parameters)
        {
            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
                using (OleDbCommand cmd = new OleDbCommand(cmdText, conn))
                {
                    if (parameters != null)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddRange(parameters);
                    }
                    cmd.CommandType = type;
                    try
                    {
                        conn.Open();
                        object res = cmd.ExecuteScalar();
                        cmd.Parameters.Clear();
                        return res;
                    }
                    catch (System.Data.OleDb.OleDbException e)
                    {
                        conn.Close();
                        throw e;
                    }
                }
            }
        }
        #endregion

        #region ExecuteReader +static void ExecuteReader(string cmdText, Action<OleDbDataReader> action, params OleDbParameter[] parameters)
        /// <summary>
        /// 利用委托 执行一个大数据查询的T-OleDb语句
        /// </summary>
        /// <param name="cmdText">要执行的T-OleDb语句</param>
        /// <param name="action">传入执行的委托对象</param>
        /// <param name="parameters">参数列表</param>
        /// <exception cref="链接数据库异常"></exception>
        public static void ExecuteReader(string cmdText, Action<OleDbDataReader> action, params OleDbParameter[] parameters)
        {
            ExecuteReader(cmdText, action, CommandType.Text, parameters);
        }
        #endregion

        #region ExecuteReader +static void ExecuteReader(string cmdText, Action<OleDbDataReader> action, CommandType type, params OleDbParameter[] parameters)
        /// <summary>
        /// 利用委托 执行一个大数据查询的T-OleDb语句
        /// </summary>
        /// <param name="cmdText">要执行的T-OleDb语句</param>
        /// <param name="action">传入执行的委托对象</param>
        /// <param name="type">命令类型</param>
        /// <param name="parameters">参数列表</param>
        /// <exception cref="链接数据库异常"></exception>
        public static void ExecuteReader(string cmdText, Action<OleDbDataReader> action, CommandType type, params OleDbParameter[] parameters)
        {
            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
                using (OleDbCommand cmd = new OleDbCommand(cmdText, conn))
                {
                    if (parameters != null)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddRange(parameters);
                    }
                    cmd.CommandType = type;
                    try
                    {
                        conn.Open();
                        using (OleDbDataReader reader = cmd.ExecuteReader())
                        {
                            action(reader);
                        }
                        cmd.Parameters.Clear();
                    }
                    catch (System.Data.OleDb.OleDbException e)
                    {
                        conn.Close();
                        throw e;
                    }
                }
            }
        }
        #endregion

        #region ExecuteReader +static OleDbDataReader ExecuteReader(string cmdText, params OleDbParameter[] parameters)
        /// <summary>
        /// 执行一个查询的T-OleDb语句, 返回一个OleDbDataReader对象, 如果出现OleDb语句执行错误, 将会关闭连接通道抛出异常
        ///  ( 注意：调用该方法后，一定要对OleDbDataReader进行Close )
        /// </summary>
        /// <param name="cmdText">要执行的T-OleDb语句</param>
        /// <param name="parameters">参数列表</param>
        /// <exception cref="链接数据库异常"></exception>
        /// <returns>OleDbDataReader对象</returns>
        public static OleDbDataReader ExecuteReader(string cmdText, params OleDbParameter[] parameters)
        {
            return ExecuteReader(cmdText, CommandType.Text, parameters);
        }
        #endregion

        #region ExecuteReader +static OleDbDataReader ExecuteReader(string cmdText, CommandType type, params OleDbParameter[] parameters)
        /// <summary>
        /// 执行一个查询的T-OleDb语句, 返回一个OleDbDataReader对象, 如果出现OleDb语句执行错误, 将会关闭连接通道抛出异常
        ///  ( 注意：调用该方法后，一定要对OleDbDataReader进行Close )
        /// </summary>
        /// <param name="cmdText">要执行的T-OleDb语句</param>
        /// <param name="type">命令类型</param>
        /// <param name="parameters">参数列表</param>
        /// <exception cref="链接数据库异常"></exception>
        /// <returns>OleDbDataReader对象</returns>
        public static OleDbDataReader ExecuteReader(string cmdText, CommandType type, params OleDbParameter[] parameters)
        {
            OleDbConnection conn = new OleDbConnection(connStr);
            using (OleDbCommand cmd = new OleDbCommand(cmdText, conn))
            {
                if (parameters != null)
                {
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddRange(parameters);
                }
                cmd.CommandType = type;
                conn.Open();
                try
                {
                    OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    cmd.Parameters.Clear();
                    return reader;
                }
                catch (System.Data.OleDb.OleDbException ex)
                {
                    //出现异常关闭连接并且释放
                    conn.Close();
                    throw ex;
                }
            }
        }
        #endregion

        #region ExecuteDataSet +static DataSet ExecuteDataSet(string cmdText, params OleDbParameter[] parameters)
        /// <summary>
        /// 执行一个查询的T-OleDb语句, 返回一个离线数据集DataSet
        /// </summary>
        /// <param name="cmdText">要执行的T-OleDb语句</param>
        /// <param name="parameters">参数列表</param>
        /// <returns>离线数据集DataSet</returns>
        public static DataSet ExecuteDataSet(string cmdText, params OleDbParameter[] parameters)
        {
            return ExecuteDataSet(cmdText, CommandType.Text, parameters);
        }
        #endregion

        #region ExecuteDataSet +static DataSet ExecuteDataSet(string cmdText, CommandType type, params OleDbParameter[] parameters)
        /// <summary>
        /// 执行一个查询的T-OleDb语句, 返回一个离线数据集DataSet
        /// </summary>
        /// <param name="cmdText">要执行的T-OleDb语句</param>
        /// <param name="type">命令类型</param>
        /// <param name="parameters">参数列表</param>
        /// <returns>离线数据集DataSet</returns>
        public static DataSet ExecuteDataSet(string cmdText, CommandType type, params OleDbParameter[] parameters)
        {
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmdText, connStr))
            {
                using (DataSet ds = new DataSet())
                {
                    if (parameters != null)
                    {
                        adapter.SelectCommand.Parameters.Clear();
                        adapter.SelectCommand.Parameters.AddRange(parameters);
                    }
                    adapter.SelectCommand.CommandType = type;
                    adapter.Fill(ds);
                    adapter.SelectCommand.Parameters.Clear();
                    return ds;
                }
            }
        }
        #endregion

        #region ExecuteDataTable +static DataTable ExecuteDataTable(string cmdText, params OleDbParameter[] parameters)
        /// <summary>
        /// 执行一个数据表查询的T-OleDb语句, 返回一个DataTable
        /// </summary>
        /// <param name="cmdText">要执行的T-OleDb语句</param>
        /// <param name="parameters">参数列表</param>
        /// <returns>查询到的数据表</returns>
        public static DataTable ExecuteDataTable(string cmdText, params OleDbParameter[] parameters)
        {
            return ExecuteDataTable(cmdText, CommandType.Text, parameters);
        }
        #endregion

        #region ExecuteDataTable +static DataTable ExecuteDataTable(string cmdText, CommandType type, params OleDbParameter[] parameters)
        /// <summary>
        /// 执行一个数据表查询的T-OleDb语句, 返回一个DataTable
        /// </summary>
        /// <param name="cmdText">要执行的T-OleDb语句</param>
        /// <param name="type">命令类型</param>
        /// <param name="parameters">参数列表</param>
        /// <returns>查询到的数据表</returns>
        public static DataTable ExecuteDataTable(string cmdText, CommandType type, params OleDbParameter[] parameters)
        {
            return ExecuteDataSet(cmdText, type, parameters).Tables[0];
        }
        #endregion
        #endregion

        #region 分页
        /// <summary>

        /// 分页查询数据并返回DataTable的公共方法

        /// </summary>

        /// <param name="tableName">表名</param>

        /// <param name="field">需要查询的字段</param>

        /// <param name="pageSize">每页显示数据的条数</param>

        /// <param name="start">排除的数据量</param>

        /// <param name="sqlWhere">where条件</param>

        /// <param name="sortName">排序名称</param>

        /// <param name="sortOrder">排序方式</param>

        /// <returns></returns>

        public static DataTable GetTable(String tableName, String field, int pageSize, int start, String sqlWhere, String sortName, String sortOrder, String primaryKey, out Int32 total)
        {

            //String sql = String.Format("select top {0} {1} from {2} where {7} and {6} not in (select top {3} {6} from {2} where {7} order by {4} {5}) order by {4} {5} ",

            //    pageSize, field, tableName, start, sortName, sortOrder, primaryKey, sqlWhere);

            String sql = String.Format("select a.* from ( select top {1} * from {2} where {7} order by {3} {4}) a left join ( select top {5} * from {2} where {7} order by {3} {4}) b on a.{6}=b.{6} where iif(b.{6},'0','1')='1'",

               field, start + pageSize, tableName, sortName, sortOrder, start, primaryKey, sqlWhere);

            if (start <= 0)
            {

                sql = String.Format("select top {0} {1} from {2} where {3} order by {4} {5} ",

                pageSize, field, tableName, sqlWhere, sortName, sortOrder);

            }

            DataTable dt = ExecuteDataTable(sql, CommandType.Text, null);



            sql = "select count(1) from " + tableName + " where " + sqlWhere;

            total = Convert.ToInt32(ExecuteScalar(sql, CommandType.Text, null));



            return dt;

        }
        #endregion

        #region 生成页码条
        //前端调用
        /*
          <div class="fy">
                     <div id="gallery">
                        <%=pagebar%>
                     </div>
          </div>
         */
        public static string RenderPager(string format, int totalPages, int current, int showCount)
        {
            string ulContainerClass = "pagination";
            string activeLiClass = "active";
            char separator='@';
            var tempFormats = format.Split(separator);
            // url 前缀
            var prefix = tempFormats[0];
            // url 后缀
            var suffix = tempFormats.Length > 1 ? tempFormats[1] : string.Empty;
            // var totalPages = Math.Max((totalCount + pageSize - 1) / pageSize, 1); //总页数
            // 左右区间
            var region = (int)Math.Floor(showCount / 2.0);
            // 开始页码数
            var beginNum = current - region <= 0 ? 1 : current - region;
            // 结束页码数
            var endNum = beginNum + showCount;
            if (endNum > totalPages)
            {
                endNum = totalPages + 1;
                beginNum = endNum - showCount;
                beginNum = beginNum < 1 ? 1 : beginNum;
            }
            var pager = new StringBuilder(string.Format("<ul class=\"{0}\">\r\n", ulContainerClass));
            if (current != 1)
            {
                pager.AppendFormat("\t<li><a href=\"{1}{0}{2}\">上一页</a></li>\r\n", current - 1, prefix, suffix);
            }
            if (beginNum != 1)
            {
                pager.Append("\t<li><span>&hellip;</span></li>\r\n");
            }
            for (var i = beginNum; i < endNum; i++)
            {
                if (i != current)
                {
                    pager.AppendFormat("\t<li><a href=\"{1}{0}{2}\">{0}</a></li>\r\n", i, prefix, suffix);
                }
                else
                {
                    pager.AppendFormat("\t<li class=\"active\"><span>{0}</span></li>\r\n", current);
                }
            }
            if (endNum != totalPages + 1)
            {
                pager.Append("\t<li><span>&hellip;</span></li>\r\n");
            }
            if (current != totalPages)
            {
                pager.AppendFormat("\t<li><a href=\"{1}{0}{2}\">下一页</a></li>\r\n", current + 1, prefix, suffix);
            }
            pager.Append("</ul>");
            return pager.ToString();
        }
        #endregion

        #region 通过事务执行
         #region 执行多条SQL语句，实现数据库事务
         /// <summary>    
         /// 执行多条SQL语句，实现数据库事务。    
         /// </summary>    
         /// <param name="SQLStringList">多条SQL语句</param>        
         public static void ExecuteSqlTran(ArrayList SQLStringList)
         {
             using (OleDbConnection conn = new OleDbConnection(connStr))
             {
                 conn.Open();
                 OleDbCommand cmd = new OleDbCommand();
                 cmd.Connection = conn;
                 OleDbTransaction tx = conn.BeginTransaction();
                 cmd.Transaction = tx;
                 try
                 {
                     for (int n = 0; n < SQLStringList.Count; n++)
                     {
                         string strsql = SQLStringList[n].ToString();
                         if (strsql.Trim().Length > 1)
                         {
                             cmd.CommandText = strsql;
                             cmd.ExecuteNonQuery();
                         }
                     }
                     tx.Commit();
                 }
                 catch (System.Data.OleDb.OleDbException E)
                 {
                     tx.Rollback();
                     throw new Exception(E.Message);
                 }
             }
         }
         #endregion
        #endregion
    }
