using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

/// <summary>
/// sqlserver数据库类库
/// </summary>
namespace Helper.SQLSERVER
{
    /// <summary>
    /// sqlserver数据库的工具类
    /// </summary>
   public class SqlServer_tool
    {

        #region 按照dt批量插入数据库表

        /// <summary>
        /// 按照DT插入数据库
        /// </summary>
        /// <param name="dt">需要insert的数据表（必须和数据库表结构一样） </param>
        /// <param name="GetConnStr">连接字符串</param>
        /// <param name="datatable_name">目的表名称</param>
        /// <returns>是否成功 成功返回ok 不成功返回错误信息</returns>
        public string Insertbulk(DataTable dt, string GetConnStr, string datatable_name)
        {

            SqlConnection conn = new SqlConnection(GetConnStr);
            conn.Open();
            SqlTransaction sqlTran = conn.BeginTransaction(); // 开始事务

            SqlBulkCopy bulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, sqlTran);//添加事务到链接
            bulkCopy.DestinationTableName = datatable_name; //设定目的表
            bulkCopy.BatchSize = dt.Rows.Count; //设定行数
            try
            {
                if (dt != null && dt.Rows.Count != 0)
                {
                    bulkCopy.WriteToServer(dt);//开始进行
                    sqlTran.Commit();
                }
            }
            catch (Exception ex)
            {
                sqlTran.Rollback();
                bulkCopy.Close();
                conn.Close();
                return ex.Message;
            }
            bulkCopy.Close();
            conn.Close();
            return "ok";

        }

        #endregion

        #region 按照dt批量更新数据库表

        /// <summary>
        /// 按照dt批量更新数据库表
        /// </summary>
        /// <param name="dt">需要更新的数据</param>
        /// <param name="data_base">目标数据库名</param>
        /// <param name="serverdt_name">目标数据库表名</param>
        /// <param name="connStr">连接字符串</param>
        /// <param name="where_list">更新数据库KEY字段</param>
        /// <returns>成功返回“ok” 失败返回“报错信息”</returns>
        public string updata_dt(DataTable dt ,string data_base, string serverdt_name,string connStr,List<string> where_list)
        {
            Helper.SQLSERVER.SqlServer_tool SqlServer_tool = new Helper.SQLSERVER.SqlServer_tool();
            DataTable dt_new = new DataTable();
            string sql = "";

            try
            {

                sql = SqlServer_tool.sys_datatable(serverdt_name);
                dt_new = SqlServer_tool.select_data(sql, connStr);

                //创建键值集合  
               Dictionary<string, string> listDic = new Dictionary<string, string>();

                for (int i = 0; i < dt_new.Rows.Count; i++)
                {
                    //用一个键值来保存数据库表的字段和数据类型  
                    //Dictionary<string, string> dic = new Dictionary<string, string>();
                    listDic.Add(dt_new.Rows[i]["列名"].ToString(), dt_new.Rows[i]["类型"].ToString());
                 
                }

                if (SqlServer_tool.IsTableExist(data_base, "datatable_temporary_ljy03042018", connStr))//判断临时表是否存在
                {
                    List<string> dt_name = new List<string>();
                    string dt_name1 = "datatable_temporary_ljy03042018";

                    dt_name.Add(dt_name1);

                    DropDataTable(data_base, dt_name, connStr); //删除临时表                     
                }

                string CreateDataTable_return = SqlServer_tool.CreateDataTable(data_base, "datatable_temporary_ljy03042018", listDic, connStr);
                if (CreateDataTable_return == "ok")
                {
                    string Insertbulk_return = SqlServer_tool.Insertbulk(dt, connStr, "datatable_temporary_ljy03042018");
                    if (Insertbulk_return == "ok")
                    {

                        #region 拼sql字符串

                        string set = "";

                        for (int i = 0; i < dt_new.Rows.Count; i++)
                        {
                            bool key_state = false;
                            foreach (string item in where_list)
                            {
                                if (dt_new.Rows[i]["列名"].ToString() == item)
                                {
                                    key_state = true;
                                }
                            }
                            if (key_state==false)
                            {
                                set = set + "a." + dt_new.Rows[i]["列名"].ToString() + "= b." + dt_new.Rows[i]["列名"].ToString() + " , ";
                            }
                          
                        }

                        set = set.Substring(0, set.Length - 3);


                        string where = "";
                        foreach (string item_where in where_list)
                        {
                            where = where + "a." + item_where + "= b." + item_where + " and ";
                        }

                        where = where.Substring(0, where.Length - 4);


                        string update = "";

                        update = "update a set " + set + " from " + serverdt_name + " a,datatable_temporary_ljy03042018 b where " + where;

                        #endregion

                        ExecuteNonQuery(update, connStr);
                    }
                    else
                    {
                        if (SqlServer_tool.IsTableExist(data_base, "datatable_temporary_ljy03042018", connStr))//判断临时表是否存在
                        {
                            List<string> dt_name = new List<string>();
                            string dt_name1 = "datatable_temporary_ljy03042018";

                            dt_name.Add(dt_name1);

                            DropDataTable(data_base, dt_name, connStr); //删除临时表                     
                        }
                        return Insertbulk_return;
                    }
                }
                else
                {
                    if (SqlServer_tool.IsTableExist(data_base, "datatable_temporary_ljy03042018", connStr))//判断临时表是否存在
                    {
                        List<string> dt_name = new List<string>();
                        string dt_name1 = "datatable_temporary_ljy03042018";

                        dt_name.Add(dt_name1);

                        DropDataTable(data_base, dt_name, connStr); //删除临时表                     
                    }
                    return CreateDataTable_return;
                }
                if (SqlServer_tool.IsTableExist(data_base, "datatable_temporary_ljy03042018", connStr))//判断临时表是否存在
                {
                    List<string> dt_name = new List<string>();
                    string dt_name1 = "datatable_temporary_ljy03042018";

                    dt_name.Add(dt_name1);

                    DropDataTable(data_base, dt_name, connStr); //删除临时表                     
                }

                return "ok";
            }
            catch (Exception ex)
            {
                if (SqlServer_tool.IsTableExist(data_base, "datatable_temporary_ljy03042018", connStr))//判断临时表是否存在
                {
                    List<string> dt_name = new List<string>();
                    string dt_name1 = "datatable_temporary_ljy03042018";

                    dt_name.Add(dt_name1);

                    DropDataTable(data_base, dt_name, connStr); //删除临时表                     
                }
                return ex.Message.ToString();
                //throw;
            }
        }

        #endregion

        #region 判断数据库表是否存在  
        /// <summary>  
        /// 判断数据库表是否存在  
        /// </summary>  
        /// <param name="db">数据库</param>  
        /// <param name="tb">数据库表名</param>  
        /// <param name="connKey">连接数据库的key</param>  
        /// <returns></returns>  
        public Boolean IsTableExist(string db, string tb, string connKey)
        {
            Helper.SQLSERVER.SqlServer_tool SqlServer_tool = new Helper.SQLSERVER.SqlServer_tool();
            //string connToMaster = ConfigurationManager.ConnectionStrings[connKey].ToString();
            string createDbStr = "use " + db + " select 1 from sysobjects where id = object_id('" + tb + "') and type ='U'";
            //在指定的数据库中 查找该表是否存在  
            bool dt_count = SqlServer_tool.Select_count(createDbStr, connKey);

            return dt_count;

        }
        #endregion

        #region 执行SQL语句并返回影响数据库中的行数

        /// <summary>
        /// 执行SQL语句并返回影响数据库中的行数
        /// </summary>
        /// <param name="sql">SQL</param>
        /// <param name="connStr">连接字符串</param>
        /// <returns>返回影响数据库中的行数</returns>
        public int ExecuteNonQuery(string sql,string connStr)
        {
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        #endregion

        #region 执行的SELECT是否返回行

        /// <summary>
        /// 执行的SELECT是否返回行
        /// </summary>
        /// <param name="sql">select语句需要count()</param>
        /// <param name="connStr">连接字符串</param>
        /// <returns>true-有行数 false-无行数</returns>
        public bool Select_count(string sql, string connStr)
        {
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    SqlDataReader dreader = cmd.ExecuteReader();

                    if (dreader.Read() == false)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                    
                }
            }
        }

        #endregion

        #region 判断数据库是否存在  
        /// <summary>  
        /// 判断数据库是否存在  
        /// </summary>  
        /// <param name="db">数据库的名称</param>  
        /// <param name="connKey">数据库的连接Key</param>  
        /// <returns>true:表示数据库已经存在；false，表示数据库不存在</returns>  
        public Boolean IsDBExist(string db, string connKey)
        {
      
            Helper.SQLSERVER.SqlServer_tool SqlServer_tool = new Helper.SQLSERVER.SqlServer_tool();
            //string connToMaster = ConfigurationManager.ConnectionStrings[connKey].ToString();
            string createDbStr = " select * from master.dbo.sysdatabases where name " + "= '" + db + "'";

            bool dt_count = SqlServer_tool.Select_count(createDbStr, connKey);

            return dt_count;

        }
        #endregion

        #region 创建数据库  
        /// <summary>  
        /// 创建数据库  
        /// </summary>  
        /// <param name="db">数据库名称</param>  
        /// <param name="connKey">连接数据库的key</param>  
        public void CreateDataBase(string db, string connKey)
        {            
            Helper.SQLSERVER.SqlServer_tool SqlServer_tool = new Helper.SQLSERVER.SqlServer_tool();
            //符号变量，判断数据库是否存在  
            Boolean flag = IsDBExist(db, connKey);

            //如果数据库存在，则抛出  
            if (flag == true)
            {

                throw new Exception("数据库已经存在！");

            }
            else
            {
                //数据库不存在，创建数据库  
                //string connToMaster = ConfigurationManager.ConnectionStrings[connKey].ToString();
                string createDbStr = "Create database " + db;
                SqlServer_tool.ExecuteNonQuery(createDbStr, connKey);
            }
        }

        #endregion

        #region 在指定的数据库中，创建数据库表  
        /// <summary>  
        ///在指定的数据库中，创建数据库表  
        /// </summary>  
        /// <param name="db">指定的数据库</param>  
        /// <param name="dt">要创建的数据库表</param>  
        /// <param name="dic">数据表中的字段及其数据类型</param>  
        /// <param name="connKey">数据库的连接Key</param>  
        /// <returns>成功返回“ok” 失败返回“报错信息”</returns>
        public string CreateDataTable(string db, string dt, Dictionary<string, string> dic, string connKey)
        {
            Helper.SQLSERVER.SqlServer_tool SqlServer_tool = new Helper.SQLSERVER.SqlServer_tool();

            //string connToMaster = ConfigurationManager.ConnectionStrings[connKey].ToString();

            //判断数据库是否存在  
            if (IsDBExist(db, connKey) == false)
            {
                return "数据库不存在！";
                //throw new Exception("数据库不存在！");


            }
            //如果数据库表存在，则抛出错误  
            if (IsTableExist(db, dt, connKey) == true)
            {
                return "数据库表已经存在！";
                //throw new Exception("数据库表已经存在！");
            }
            //数据库表不存在，创建表  
            else
            {
                //拼接字符串，该串为创建内容  
                //string content = "serial int identity(1,1) primary key";
                string content = "";
                //取出dic中的内容，进行拼接  
                List<string> Keys = new List<string>(dic.Keys);
                List<string> Values = new List<string>(dic.Values);
                for (int i = 0; i < dic.Count(); i++)
                {
                    content = content + "," + Keys[i] + " " + Values[i];
                }
                content = content.Remove(0, 1);
                //其后判断数据库表是否存在，然创建表  
                string createTableStr = "use " + db + " create table " + dt + "(" + content + ")";
                SqlServer_tool.ExecuteNonQuery(createTableStr, connKey);

            }

            return "ok";
        }
        #endregion

        #region 批量删除数据表  
        /// <summary>  
        /// 批量删除数据库  
        /// </summary>  
        /// <param name="db">指定的数据库</param>  
        /// <param name="dt">要删除的数据库表集合</param>  
        /// <param name="connKey">数据库连接串</param>  
        /// <returns>删除是否成功，true表示删除成功，false表示删除失败</returns>  
        public bool DropDataTable(string db, List<string> dt, string connKey)
        {
            //string connToMaster = ConfigurationManager.ConnectionStrings[connKey].ToString();

            Helper.SQLSERVER.SqlServer_tool SqlServer_tool = new Helper.SQLSERVER.SqlServer_tool();

            //判断数据库是否存在  
            if (IsDBExist(db, connKey) == false)
            {
                throw new Exception("数据库不存在！");
            }

            for (int i = 0; i < dt.Count(); i++)
            {
                //如果数据库表存在，则抛出错误  
                if (IsTableExist(db, dt[i], connKey) == false)
                {
                    //如果数据库表已经删除，则跳过该表  
                    continue;
                }
                else//数据表存在，则进行删除数据表  
                {
                    //其后判断数据表是否存在，然后创建数据表  
                    string createTableStr = "use " + db + " drop table " + dt[i] + " ";
                    SqlServer_tool.ExecuteNonQuery(createTableStr, connKey);
                }
            }
            return true;
        }
        #endregion

        #region 删除数据库  
        /// <summary>  
        /// 删除数据库  
        /// </summary>  
        /// <param name="db">数据库名</param>  
        /// <param name="connKey">数据库连接串</param>  
        /// <returns>删除成功为true，删除失败为false</returns>  
        public bool DropDataBase(string db, string connKey)
        {
            Helper.SQLSERVER.SqlServer_tool SqlServer_tool = new Helper.SQLSERVER.SqlServer_tool();
            //SQLHelper helper = SQLHelper.GetInstance();
            //符号变量，判断数据库是否存在  
            Boolean flag = IsDBExist(db, connKey);

            //如果数据库不存在，则抛出  
            if (flag == false)
            {
                return false;
            }
            else
            {
                //数据库存在，删除数据库  
                //string connToMaster = ConfigurationManager.ConnectionStrings[connKey].ToString();
                string createDbStr = "Drop database " + db;
                SqlServer_tool.ExecuteNonQuery(createDbStr, connKey);
                return true;
            }
        }
        #endregion

        #region 创建数据库表(多张表)  
        /// <summary>  
        ///  在指定的数据库中，创建数据表  
        /// </summary>  
        /// <param name="db">指定的数据库</param>  
        /// <param name="dt">要创建的数据表集合</param>  
        /// <param name="dic">数据表中的字段及其数据类型  Dictionary集合</param>  
        /// <param name="connKey">数据库的连接Key</param>  
        public void CreateDataTable(string db, string[] dt, List<Dictionary<string, string>> dic, string connKey)
        {
            Helper.SQLSERVER.SqlServer_tool SqlServer_tool = new Helper.SQLSERVER.SqlServer_tool();

            //string connToMaster = ConfigurationManager.ConnectionStrings[connKey].ToString();

            //判断数据库是否存在  
            if (IsDBExist(db, connKey) == false)
            {
                throw new Exception("数据库不存在！");
            }

            for (int i = 0; i < dt.Count(); i++)
            {
                //如果数据库表存在，则抛出错误  
                if (IsTableExist(db, dt[i], connKey) == true)
                {
                    //如果数据库表已经存在，则跳过该表  
                    continue;
                }
                else//数据表不存在，创建数据表  
                {
                    //其后判断数据表是否存在，然后创建数据表  
                    string createTableStr = PinjieSql(db, dt[i], dic[i]);
                    SqlServer_tool.ExecuteNonQuery(createTableStr, connKey);
                }
            }
        }
        #endregion

        #region 拼接创建数据库表的Sql语句  
        /// <summary>  
        /// 拼接创建数据库表的Sql语句  
        /// </summary>  
        /// <param name="db">指定的数据库</param>  
        /// <param name="dt">要创建的数据表</param>  
        /// <param name="dic">数据表中的字段及其数据类型</param>  
        /// <returns>拼接完的字符串</returns>  
        public string PinjieSql(string db, string dt, Dictionary<string, string> dic)
        {
            //拼接字符串，（该串为创建内容）  
            string content = "serial int identity(1,1) primary key ";
            //取出dic中的内容，进行拼接  
            List<string> test = new List<string>(dic.Keys);
            for (int i = 0; i < dic.Count(); i++)
            {
                content = content + " , " + test[i] + " " + dic[test[i]];
            }

            //其后判断数据表是否存在，然后创建数据表  
            string createTableStr = "use " + db + " create table " + dt + " (" + content + ")";
            return createTableStr;
        }

        #endregion

        #region 查询数据库并返回dt

        /// <summary>
        /// 查询数据库并返回dt
        /// </summary>
        /// <param name="sql">sql字符串</param>
        /// <param name="connStr">连接字符串</param>
        /// <returns></returns>
        public DataTable select_data(string sql,string connStr)
        {          
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    DataTable dt = new DataTable();
                    SqlDataAdapter myda = new SqlDataAdapter(sql, conn); // 实例化适配器
                    myda.Fill(dt); // 保存数据                 

                    conn.Close(); // 关闭数据库连接

                    return dt;
                }
            }          
        }

        #endregion

        #region 返回数据库表的结构的SQL字符串

        /// <summary>
        /// 返回数据库表的结构的SQL字符串
        /// </summary>
        /// <param name="datatable_name">数据库表名</param>
        /// <returns>输入表的结构</returns>
        public string sys_datatable(string datatable_name)
        {

            string sql = "select   " +
                    "    [表名]=c.Name,  " +
                    "    [表说明]=isnull(f.[value],''),   " +
                    "    [列序号]=a.Column_id,   " +
                    "    [列名]=a.Name,   " +
                    "    [列说明]=isnull(e.[value],''),   " +
                    "    [数据库类型]=b.Name,     " +
                    "    [类型]= case when b.Name = 'image' then 'byte[]' " +
                    "                  when b.Name in('image','uniqueidentifier','ntext','varchar','ntext','nchar','nvarchar','text','char') then 'nvarchar(4000)' " +
                    "                  when b.Name in('tinyint','smallint','int','bigint') then 'int' " +
                    "                  when b.Name in('datetime','smalldatetime') then 'DateTime' " +
                    "                  when b.Name in('float','decimal','numeric','money','real','smallmoney') then 'decimal' " +
                    "                  when b.Name ='bit' then 'bool' else b.name end , " +
                    "    [标识]= case when is_identity=1 then '是' else '' end,   " +
                    "    [主键]= case when exists(select 1 from sys.objects x join sys.indexes y on x.Type=N'PK' and x.Name=y.Name   " +
                    "                         join sysindexkeys z on z.ID=a.Object_id and z.indid=y.index_id and z.Colid=a.Column_id)   " +
                    "                     then '是' else '' end,       " +
                    "    [字节数]=case when a.[max_length]=-1 and b.Name!='xml' then 'max/2G'   " +
                    "                   when b.Name='xml' then '2^31-1字节/2G'  " +
                    "                   else rtrim(a.[max_length]) end,   " +
                    "    [长度]=case when ColumnProperty(a.object_id,a.Name,'Precision')=-1 then '2^31-1'  " +
                    "                 else rtrim(ColumnProperty(a.object_id,a.Name,'Precision')) end,   " +
                    "    [小数位]=isnull(ColumnProperty(a.object_id,a.Name,'Scale'),0),   " +
                    "    [是否为空]=case when a.is_nullable=1 then '是' else '' end,       " +
                    "    [默认值]=isnull(d.text,'')       " +
                    " from   " +
                    "     sys.columns a   " +
                    " left join  " +
                    "     sys.types b on a.user_type_id=b.user_type_id   " +
                    " inner join  " +
                    "     sys.objects c on a.object_id=c.object_id and c.Type='U'  " +
                    " left join  " +
                    "     syscomments d on a.default_object_id=d.ID   " +
                    " left join  " +
                    "     sys.extended_properties e on e.major_id=c.object_id and e.minor_id=a.Column_id and e.class=1    " +
                    " left join  " +
                    "     sys.extended_properties f on f.major_id=c.object_id and f.minor_id=0 and f.class=1  " +
                    "where c.name = '"+ datatable_name + "' ";

            return sql;
        }

        #endregion

    }
}
