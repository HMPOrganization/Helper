using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using Newtonsoft.Json;

using Oracle.ManagedDataAccess.Client;

namespace Helper.OracleHelper
{
    /// <summary>
    /// Oracle数据库帮助类
    /// </summary>
    public abstract class OracleHelper
    {

        //为缓存的参数创建哈希表
        private static Hashtable parmCache = Hashtable.Synchronized(new Hashtable());

        /// <summary>
        /// 执行不包含选择的数据库查询
        /// </summary>
        /// <param name="connectionString">到数据库的连接字符串</param>
        /// <param name="cmdType">命令类型：存储过程或SQL</param>
        /// <param name="cmdText">ActualSQL命令</param>
        /// <param name="commandParameters">绑定到命令的参数</param>
        /// <returns></returns>
        public static int ExecuteNonQuery(string connectionString, CommandType cmdType, string cmdText, params OracleParameter[] commandParameters)
        {
            // 创建新的Oracle命令
            OracleCommand cmd = new OracleCommand();

            //创建连接
            using (OracleConnection connection = new OracleConnection(connectionString))
            {

                //准备命令
                PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);

                //执行命令
                int val = cmd.ExecuteNonQuery();
                connection.Close();
                cmd.Parameters.Clear();
                return val;
            }
        }


        /// <summary>
        /// 执行查询语句，返回DataSet
        /// </summary>
        /// <param name="connectionString">到数据库的连接字符串</param>
        /// <param name="SQLString">查询语句</param>
        /// <param name="cmdParms"></param>
        /// <returns></returns>
        public static DataSet Query(string connectionString, string SQLString, params OracleParameter[] cmdParms)
        {
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                OracleCommand cmd = new OracleCommand();
                PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                using (OracleDataAdapter da = new OracleDataAdapter(cmd))
                {
                    DataSet ds = new DataSet();
                    try
                    {
                        da.Fill(ds, "ds");
                        cmd.Parameters.Clear();
                    }
                    catch (System.Data.SqlClient.SqlException ex)
                    {
                        throw new Exception(ex.Message);
                    }
                    return ds;
                }
            }
        }

        private static void PrepareCommand(OracleCommand cmd, OracleConnection conn, OracleTransaction trans, string cmdText, OracleParameter[] cmdParms)
        {
            if (conn.State != ConnectionState.Open)
                conn.Open();
            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            if (trans != null)
                cmd.Transaction = trans;
            cmd.CommandType = CommandType.Text;//cmdType;
            if (cmdParms != null)
            {
                foreach (OracleParameter parameter in cmdParms)
                {
                    if ((parameter.Direction == ParameterDirection.InputOutput || parameter.Direction == ParameterDirection.Input) &&
                        (parameter.Value == null))
                    {
                        parameter.Value = DBNull.Value;
                    }
                    cmd.Parameters.Add(parameter);
                }
            }
        }

        /// <summary>
        /// 执行一条计算查询结果语句，返回查询结果（object）。
        /// </summary>
        /// <param name="connectionString">到数据库的连接字符串</param>
        /// <param name="SQLString">计算查询结果语句</param>
        /// <returns>查询结果（object）</returns>
        public static object GetSingle(string connectionString, string SQLString)
        {
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                using (OracleCommand cmd = new OracleCommand(SQLString, connection))
                {
                    try
                    {
                        connection.Open();
                        object obj = cmd.ExecuteScalar();
                        if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
                        {
                            return null;
                        }
                        else
                        {
                            return obj;
                        }
                    }
                    catch (OracleException ex)
                    {
                        throw new Exception(ex.Message);
                    }
                    finally
                    {
                        if (connection.State != ConnectionState.Closed)
                        {
                            connection.Close();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 需要查询的信息是否存在
        /// </summary>
        /// <param name="connectionString">到数据库的连接字符串</param>
        /// <param name="strOracle">sql字符串</param>
        /// <returns></returns>
        public static bool Exists(string connectionString, string strOracle)
        {
            object obj = OracleHelper.GetSingle(connectionString, strOracle);
            int cmdresult;
            if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
            {
                cmdresult = 0;
            }
            else
            {
                cmdresult = int.Parse(obj.ToString());
            }
            if (cmdresult == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// 使用所提供的参数对现有数据库事务执行OraceCoMMand（不返回结果集）
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  int result = ExecuteNonQuery(trans, CommandType.StoredProcedure, "PublishOrders", new OracleParameter(":prodid", 24));
        /// </remarks>
        /// <param name="trans">现有数据库事务</param>
        /// <param name="cmdType">命令类型（存储过程、文本等）the CommandType (stored procedure, text, etc.)</param>
        /// <param name="cmdText">存储过程名称或PL/SQL命令 the stored procedure name or PL/SQL command</param>
        /// <param name="commandParameters">用于执行命令的Oracle PARAMTER数组 an array of OracleParamters used to execute the command</param>
        /// <returns>表示受命令影响的行数的int an int representing the number of rows affected by the command</returns>
        public static int ExecuteNonQuery(OracleTransaction trans, CommandType cmdType, string cmdText, params OracleParameter[] commandParameters)
        {
            OracleCommand cmd = new OracleCommand();
            PrepareCommand(cmd, trans.Connection, trans, cmdType, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }

        /// <summary>
        /// 使用所提供的参数对现有数据库连接执行OraceCoMMand（不返回结果集）
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new OracleParameter(":prodid", 24));
        /// </remarks>
        /// <param name="connection">现有数据库连接</param>
        /// <param name="cmdType">命令类型（存储过程、文本等）the CommandType (stored procedure, text, etc.)</param>
        /// <param name="cmdText">存储过程名称或PL/SQL命令 the stored procedure name or PL/SQL command</param>
        /// <param name="commandParameters">用于执行命令的Oracle PARAMTER数组 an array of OracleParamters used to execute the command</param>
        /// <returns>表示受命令影响的行数的int an int representing the number of rows affected by the command</returns>
        public static int ExecuteNonQuery(OracleConnection connection, CommandType cmdType, string cmdText, params OracleParameter[] commandParameters)
        {

            OracleCommand cmd = new OracleCommand();

            PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }
        /// <summary>
        /// 使用所提供的参数对现有数据库连接执行OraceCoMMand（不返回结果集）。
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new OracleParameter(":prodid", 24));
        /// </remarks>
        /// <param name="connectionString">现有数据库连接</param>
        /// <param name="cmdText">存储过程名称或PL/SQL命令</param>
        /// <returns>表示受命令影响的行数的int</returns>
        public static int ExecuteNonQuery(string connectionString, string cmdText)
        {

            OracleCommand cmd = new OracleCommand();
            OracleConnection connection = new OracleConnection(connectionString);
            PrepareCommand(cmd, connection, null, CommandType.Text, cmdText, null);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }

        /// <summary>
        /// Execute a select query that will return a result set
        /// </summary>
        /// <param name="connectionString">Connection string</param>
        /// <param name="cmdType">命令类型（存储过程、文本等）</param>
        /// <param name="cmdText">存储过程名称或PL/SQL命令</param>
        /// <param name="commandParameters">用于执行命令的Oracle PARAMTER数组</param>
        /// <returns></returns>
        public static OracleDataReader ExecuteReader(string connectionString, CommandType cmdType, string cmdText, params OracleParameter[] commandParameters)
        {
            OracleCommand cmd = new OracleCommand();
            OracleConnection conn = new OracleConnection(connectionString);
            try
            {
                //Prepare the command to execute
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                OracleDataReader rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                cmd.Parameters.Clear();
                return rdr;
            }
            catch
            {
                conn.Close();
                throw;
            }
        }

        /// <summary>
        /// 执行OraceCoMMand，使用所提供的参数返回连接字符串中指定的数据库中的第一条记录的第一列。
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  Object obj = ExecuteScalar(connString, CommandType.StoredProcedure, "PublishOrders", new OracleParameter(":prodid", 24));
        /// </remarks>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdType">命令类型（存储过程、文本等）</param>
        /// <param name="cmdText">存储过程名称或PL/SQL命令</param>
        /// <param name="commandParameters">用于执行命令的Oracle PARAMTER数组</param>
        /// <returns>应该使用转换为{Type }转换为预期类型的对象。</returns>
        public static object ExecuteScalar(string connectionString, CommandType cmdType, string cmdText, params OracleParameter[] commandParameters)
        {
            OracleCommand cmd = new OracleCommand();

            using (OracleConnection conn = new OracleConnection(connectionString))
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                object val = cmd.ExecuteScalar();
                cmd.Parameters.Clear();
                return val;
            }
        }

        ///    <summary>
        ///    使用所提供的参数对指定的SQL事务执行OracleCommand（返回1x1结果集）。
        ///    </summary>
        ///    <param name="transaction">现有的SQL事务</param>
        ///    <param name="commandType">命令类型（存储过程、文本等）</param>
        ///    <param name="commandText">存储过程名称或PL/SQL命令</param>
        ///    <param name="commandParameters">用于执行命令的Oracle PARAMTER数组</param>
        ///    <returns>包含由命令生成的1x1结果集中的值的对象</returns>
        public static object ExecuteScalar(OracleTransaction transaction, CommandType commandType, string commandText, params OracleParameter[] commandParameters)
        {
            if (transaction == null)
                throw new ArgumentNullException("transaction");
            if (transaction != null && transaction.Connection == null)
                throw new ArgumentException("The transaction was rollbacked    or commited, please    provide    an open    transaction.", "transaction");

            // Create a    command    and    prepare    it for execution
            OracleCommand cmd = new OracleCommand();

            PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters);

            // Execute the command & return    the    results
            object retval = cmd.ExecuteScalar();

            // Detach the SqlParameters    from the command object, so    they can be    used again
            cmd.Parameters.Clear();
            return retval;
        }

        /// <summary>
        /// 执行OraceCopMand，使用所提供的参数返回现有数据库连接的第一条记录的第一列。
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  Object obj = ExecuteScalar(conn, CommandType.StoredProcedure, "PublishOrders", new OracleParameter(":prodid", 24));
        /// </remarks>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdType">命令类型（存储过程、文本等）</param>
        /// <param name="cmdText">存储过程名称或PL/SQL命令</param>
        /// <param name="commandParameters">用于执行命令的Oracle PARAMTER数组</param>
        /// <returns>应该使用转换为{Type }转换为预期类型的对象。</returns>
        public static object ExecuteScalar(OracleConnection connectionString, CommandType cmdType, string cmdText, params OracleParameter[] commandParameters)
        {
            OracleCommand cmd = new OracleCommand();

            PrepareCommand(cmd, connectionString, null, cmdType, cmdText, commandParameters);
            object val = cmd.ExecuteScalar();
            cmd.Parameters.Clear();
            return val;
        }

        /// <summary>
        /// 向缓存添加一组参数
        /// </summary>
        /// <param name="cacheKey">查找参数的键值</param>
        /// <param name="commandParameters">缓存的实际参数</param>
        public static void CacheParameters(string cacheKey, params OracleParameter[] commandParameters)
        {
            parmCache[cacheKey] = commandParameters;
        }

        /// <summary>
        /// 从缓存中获取参数
        /// </summary>
        /// <param name="cacheKey">查找参数的关键</param>
        /// <returns></returns>
        public static OracleParameter[] GetCachedParameters(string cacheKey)
        {
            OracleParameter[] cachedParms = (OracleParameter[])parmCache[cacheKey];

            if (cachedParms == null)
                return null;

            // If the parameters are in the cache
            OracleParameter[] clonedParms = new OracleParameter[cachedParms.Length];

            // return a copy of the parameters
            for (int i = 0, j = cachedParms.Length; i < j; i++)
                clonedParms[i] = (OracleParameter)((ICloneable)cachedParms[i]).Clone();

            return clonedParms;
        }
        /// <summary>
        /// 内部函数准备由数据库执行的命令
        /// </summary>
        /// <param name="cmd">现有命令对象</param>
        /// <param name="conn">数据库连接对象</param>
        /// <param name="trans">可选事务对象</param>
        /// <param name="cmdType">命令类型，例如存储过程</param>
        /// <param name="cmdText">命令测试</param>
        /// <param name="commandParameters">命令的参数</param>
        private static void PrepareCommand(OracleCommand cmd, OracleConnection conn, OracleTransaction trans, CommandType cmdType, string cmdText, OracleParameter[] commandParameters)
        {

            //Open the connection if required
            if (conn.State != ConnectionState.Open)
                conn.Open();

            //Set up the command
            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            cmd.CommandType = cmdType;

            //Bind it to the transaction if it exists
            if (trans != null)
                cmd.Transaction = trans;

            // Bind the parameters passed in
            if (commandParameters != null)
            {
                foreach (OracleParameter parm in commandParameters)
                    cmd.Parameters.Add(parm);
            }
        }

        /// <summary>
        /// 使用Oracle使用布尔数据类型的转换器
        /// </summary>
        /// <param name="value">转换值 true或false</param>
        /// <returns>返回 Y或N</returns>
        public static string OraBit(bool value)
        {
            if (value)
                return "Y";
            else
                return "N";
        }

        /// <summary>
        /// 使用Oracle使用布尔数据类型的转换器
        /// </summary>
        /// <param name="value">转换值 Y或N</param>
        /// <returns>返回true或false</returns>
        public static bool OraBool(string value)
        {
            if (value.Equals("Y"))
                return true;
            else
                return false;
        }

    }
}
