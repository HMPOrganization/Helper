using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Reflection;
using System.ComponentModel;
using System.Web.Script.Serialization;
using System.Runtime.Serialization;
using System.IO;
using System.Text.RegularExpressions;

/// <summary>
/// 数据转换类类库
/// </summary>
namespace Helper.Transformation
{
    public class Transformation
    {
        #region 数据转换类  
        /// <summary>  
        /// 数据转换类  
        /// 说明    ：实体,List集合,DataTable相互转换<br/>  
        /// </summary>  
        public class DataConvert
        {

            #region DataRow转实体  
            /// <summary>  
            /// DataRow转实体  
            /// </summary>  
            /// <typeparam name="T">数据型类</typeparam>  
            /// <param name="dr">DataRow</param>  
            /// <returns>模式</returns>  

            public static T DataRowToModel<T>(DataRow dr) where T : new()
            {
                //T t = (T)Activator.CreateInstance(typeof(T));  
                T t = new T();
                if (dr == null) return default(T);
                // 获得此模型的公共属性  
                PropertyInfo[] propertys = t.GetType().GetProperties();
                DataColumnCollection Columns = dr.Table.Columns;
                foreach (PropertyInfo p in propertys)
                {
                    //string columnName = ((DBColumn)p.GetCustomAttributes(typeof(DBColumn), false)[0]).ColName;
                    string columnName = p.Name; //如果不用属性，数据库字段对应model属性,就用这个
                    if (Columns.Contains(columnName))
                    {
                        // 判断此属性是否有Setter或columnName值是否为空  
                        object value = dr[columnName];
                        if (!p.CanWrite || value is DBNull || value == DBNull.Value) continue;
                        try
                        {
                            #region SetValue  
                            switch (p.PropertyType.ToString())
                            {
                                case "System.String":
                                    p.SetValue(t, Convert.ToString(value), null);
                                    break;
                                case "System.Int32":
                                    p.SetValue(t, Convert.ToInt32(value), null);
                                    break;
                                case "System.Int64":
                                    p.SetValue(t, Convert.ToInt64(value), null);
                                    break;
                                case "System.DateTime":
                                    p.SetValue(t, Convert.ToDateTime(value), null);
                                    break;
                                case "System.Boolean":
                                    p.SetValue(t, Convert.ToBoolean(value), null);
                                    break;
                                case "System.Double":
                                    p.SetValue(t, Convert.ToDouble(value), null);
                                    break;
                                case "System.Decimal":
                                    p.SetValue(t, Convert.ToDecimal(value), null);
                                    break;
                                default:
                                    p.SetValue(t, value, null);
                                    break;
                            }
                            #endregion
                        }
                        catch (Exception)
                        {

                            continue;
                            /*使用 default 关键字，此关键字对于引用类型会返回空，对于数值类型会返回零。对于结构， 
                             * 此关键字将返回初始化为零或空的每个结构成员，具体取决于这些结构是值类型还是引用类型*/
                        }
                    }
                }
                return t;
            }
            #endregion

            #region DataTable转List<T>  
            /// <summary>  
            /// DataTable转List<T>  
            /// </summary>  
            /// <typeparam name="T">数据项类型</typeparam>  
            /// <param name="dt">DataTable</param>  
            /// <returns>List数据集</returns>  
            public static IList<T> DataTableToList<T>(DataTable dt) where T : new()
            {
                IList<T> tList = new List<T>();
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        T t = DataRowToModel<T>(dr);
                        tList.Add(t);
                    }
                }
                return tList;
            }
            #endregion

            #region DataReader转实体  
            /// <summary>  
            /// DataReader转实体  
            /// </summary>  
            /// <typeparam name="T">数据类型</typeparam>  
            /// <param name="dr">DataReader</param>  
            /// <returns>实体</returns>  
            public static T DataReaderToModel<T>(IDataReader dr) where T : new()
            {
                T t = new T();
                if (dr == null) return default(T);
                using (dr)
                {
                    if (dr.Read())
                    {
                        // 获得此模型的公共属性  
                        PropertyInfo[] propertys = t.GetType().GetProperties();
                        List<string> listFieldName = new List<string>(dr.FieldCount);
                        for (int i = 0; i < dr.FieldCount; i++)
                        {
                            listFieldName.Add(dr.GetName(i).ToLower());
                        }

                        foreach (PropertyInfo p in propertys)
                        {
                            string columnName = p.Name;
                            if (listFieldName.Contains(columnName.ToLower()))
                            {
                                // 判断此属性是否有Setter或columnName值是否为空  
                                object value = dr[columnName];
                                if (!p.CanWrite || value is DBNull || value == DBNull.Value) continue;
                                try
                                {
                                    #region SetValue  
                                    switch (p.PropertyType.ToString())
                                    {
                                        case "System.String":
                                            p.SetValue(t, Convert.ToString(value), null);
                                            break;
                                        case "System.Int32":
                                            p.SetValue(t, Convert.ToInt32(value), null);
                                            break;
                                        case "System.DateTime":
                                            p.SetValue(t, Convert.ToDateTime(value), null);
                                            break;
                                        case "System.Boolean":
                                            p.SetValue(t, Convert.ToBoolean(value), null);
                                            break;
                                        case "System.Double":
                                            p.SetValue(t, Convert.ToDouble(value), null);
                                            break;
                                        case "System.Decimal":
                                            p.SetValue(t, Convert.ToDecimal(value), null);
                                            break;
                                        default:
                                            p.SetValue(t, value, null);
                                            break;
                                    }
                                    #endregion
                                }
                                catch
                                {
                                    //throw (new Exception(ex.Message));  
                                }
                            }
                        }
                    }
                }
                return t;
            }
            #endregion

            #region DataReader转List<T>  
            /// <summary>  
            /// DataReader转List<T>  
            /// </summary>  
            /// <typeparam name="T">数据类型</typeparam>  
            /// <param name="dr">DataReader</param>  
            /// <returns>List数据集</returns>  
            public static List<T> DataReaderToList<T>(IDataReader dr) where T : new()
            {
                List<T> tList = new List<T>();
                if (dr == null) return tList;
                T t1 = new T();
                // 获得此模型的公共属性  
                PropertyInfo[] propertys = t1.GetType().GetProperties();
                using (dr)
                {
                    List<string> listFieldName = new List<string>(dr.FieldCount);
                    for (int i = 0; i < dr.FieldCount; i++)
                    {
                        listFieldName.Add(dr.GetName(i).ToLower());
                    }
                    while (dr.Read())
                    {
                        T t = new T();
                        foreach (PropertyInfo p in propertys)
                        {
                            string columnName = p.Name;
                            if (listFieldName.Contains(columnName.ToLower()))
                            {
                                // 判断此属性是否有Setter或columnName值是否为空  
                                object value = dr[columnName];
                                if (!p.CanWrite || value is DBNull || value == DBNull.Value) continue;
                                try
                                {
                                    #region SetValue  
                                    switch (p.PropertyType.ToString())
                                    {
                                        case "System.String":
                                            p.SetValue(t, Convert.ToString(value), null);
                                            break;
                                        case "System.Int32":
                                            p.SetValue(t, Convert.ToInt32(value), null);
                                            break;
                                        case "System.DateTime":
                                            p.SetValue(t, Convert.ToDateTime(value), null);
                                            break;
                                        case "System.Boolean":
                                            p.SetValue(t, Convert.ToBoolean(value), null);
                                            break;
                                        case "System.Double":
                                            p.SetValue(t, Convert.ToDouble(value), null);
                                            break;
                                        case "System.Decimal":
                                            p.SetValue(t, Convert.ToDecimal(value), null);
                                            break;
                                        default:
                                            p.SetValue(t, value, null);
                                            break;
                                    }
                                    #endregion
                                }
                                catch
                                {
                                    //throw (new Exception(ex.Message));  
                                }
                            }
                        }
                        tList.Add(t);
                    }
                }
                return tList;
            }
            #endregion

            #region 泛型集合转DataTable  
            /// <summary>  
            /// 泛型集合转DataTable  
            /// </summary>  
            /// <typeparam name="T">集合类型</typeparam>  
            /// <param name="entityList">泛型集合</param>  
            /// <returns>DataTable</returns>  
            public static DataTable ListToDataTable<T>(IList<T> entityList)
            {
                if (entityList == null) return null;
                DataTable dt = CreateTable<T>();
                Type entityType = typeof(T);
                //PropertyInfo[] properties = entityType.GetProperties();  
                PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(entityType);
                foreach (T item in entityList)
                {
                    DataRow row = dt.NewRow();
                    foreach (PropertyDescriptor property in properties)
                    {
                        row[property.Name] = property.GetValue(item);
                    }
                    dt.Rows.Add(row);
                }

                return dt;
            }

            #endregion

            #region 创建DataTable的结构  
            private static DataTable CreateTable<T>()
            {
                Type entityType = typeof(T);
                //PropertyInfo[] properties = entityType.GetProperties();  
                PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(entityType);
                //生成DataTable的结构  
                DataTable dt = new DataTable();
                foreach (PropertyDescriptor prop in properties)
                {
                    dt.Columns.Add(prop.Name);
                }
                return dt;
            }
            #endregion

            #region 数据库字段对应属性类

            /// <summary>  
            /// 数据库字段对应属性类  
            /// 说明    ：数据库字段对应属性类<br/>  
            /// 作者    ：niu<br/>  
            /// 创建时间：2011-07-21<br/>  
            /// 最后修改：2011-07-21<br/>  
            /// </summary>  
            [AttributeUsage(AttributeTargets.Property | AttributeTargets.Method, AllowMultiple = true, Inherited = false)]
            public class DBColumn : Attribute
            {
                private string _colName;
                /// <summary>  
                /// 数据库字段  
                /// </summary>  
                public string ColName
                {
                    get { return _colName; }
                    set { _colName = value; }
                }
                /* 
         AttributeTargets 枚举  
         成员名称 说明  
         All 可以对任何应用程序元素应用属性。   
         Assembly 可以对程序集应用属性。   
         Class 可以对类应用属性。   
         Constructor 可以对构造函数应用属性。   
         Delegate 可以对委托应用属性。   
         Enum 可以对枚举应用属性。   
         Event 可以对事件应用属性。   
         Field 可以对字段应用属性。   
         GenericParameter 可以对泛型参数应用属性。   
         Interface 可以对接口应用属性。   
         Method 可以对方法应用属性。   
         Module 可以对模块应用属性。 注意  
         Module 指的是可移植的可执行文件（.dll 或 .exe），而非 Visual Basic 标准模块。 
         Parameter 可以对参数应用属性。   
         Property 可以对属性 (Property) 应用属性 (Attribute)。   
         ReturnValue 可以对返回值应用属性。   
         Struct 可以对结构应用属性，即值类型  
             */

                /* 
                 这里会有四种可能的组合：  

          [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false ]  
          [AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = false ]  
          [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true ]  
          [AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = true ] 

        　　第一种情况：  

        　　如果我们查询Derive类，我们将会发现Help特性并不存在，因为inherited属性被设置为false。 

        　　第二种情况： 

        　　和第一种情况相同，因为inherited也被设置为false。 

        　　第三种情况： 

        　　为了解释第三种和第四种情况，我们先来给派生类添加点代码：  

          [Help("BaseClass")]  
          public class Base  
          {  
          }  
          [Help("DeriveClass")]  
          public class Derive : Base  
          {  
          } 

        　　现在我们来查询一下Help特性，我们只能得到派生类的属性，因为inherited被设置为true，但是AllowMultiple却被设置为false。因此基类的Help特性被派生类Help特性覆盖了。 

        　　第四种情况： 

        　　在这里，我们将会发现派生类既有基类的Help特性，也有自己的Help特性，因为AllowMultiple被设置为true。 
                 */
            }

            #endregion

            #region datatable转换为json
            /// <summary>
            /// 将datatable转换为json  
            /// </summary>
            /// <param name="dtb">Dt</param>
            /// <returns>JSON字符串</returns>
            public static string Dtb2Json(DataTable dtb)
            {
                JavaScriptSerializer jss = new JavaScriptSerializer();
                System.Collections.ArrayList dic = new System.Collections.ArrayList();
                foreach (DataRow dr in dtb.Rows)
                {
                    System.Collections.Generic.Dictionary<string, object> drow = new System.Collections.Generic.Dictionary<string, object>();
                    foreach (DataColumn dc in dtb.Columns)
                    {
                        drow.Add(dc.ColumnName, dr[dc.ColumnName]);
                    }
                    dic.Add(drow);

                }
                //序列化  
                return jss.Serialize(dic);
            }
            #endregion

            #region json转datatable

            /// <summary>
            /// 将json转换为DataTable
            /// </summary>
            /// <param name="strJson">得到的json</param>
            /// <returns></returns>
            public static DataTable JsonToDataTable(string strJson)
            {
                //转换json格式
                strJson = strJson.Replace(",\"", "*\"").Replace("\":", "\"#").ToString();
                //取出表名   
                var rg = new Regex(@"(?<={)[^:]+(?=:\[)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                string strName = rg.Match(strJson).Value;
                DataTable tb = null;
                //去除表名   
                strJson = strJson.Substring(strJson.IndexOf("[") + 1);
                strJson = strJson.Substring(0, strJson.IndexOf("]"));

                //获取数据   
                rg = new Regex(@"(?<={)[^}]+(?=})");
                MatchCollection mc = rg.Matches(strJson);
                for (int i = 0; i < mc.Count; i++)
                {
                    string strRow = mc[i].Value;
                    string[] strRows = strRow.Split('*');

                    //创建表   
                    if (tb == null)
                    {
                        tb = new DataTable();
                        tb.TableName = strName;
                        foreach (string str in strRows)
                        {
                            var dc = new DataColumn();
                            string[] strCell = str.Split('#');

                            if (strCell[0].Substring(0, 1) == "\"")
                            {
                                int a = strCell[0].Length;
                                dc.ColumnName = strCell[0].Substring(1, a - 2);
                            }
                            else
                            {
                                dc.ColumnName = strCell[0];
                            }
                            tb.Columns.Add(dc);
                        }
                        tb.AcceptChanges();
                    }

                    //增加内容   
                    DataRow dr = tb.NewRow();
                    for (int r = 0; r < strRows.Length; r++)
                    {
                        dr[r] = strRows[r].Split('#')[1].Trim().Replace("，", ",").Replace("：", ":").Replace("\"", "");
                    }
                    tb.Rows.Add(dr);
                    tb.AcceptChanges();
                }

                return tb;
            }

            #endregion

            #region  DataSet转换为Json
            /// <summary>    
            /// DataSet转换为Json   
            /// </summary>    
            /// <param name="dataSet">DataSet对象</param>   
            /// <returns>Json字符串</returns>    
            public static string ToJson(DataSet dataSet)
            {
                string jsonString = "{";
                foreach (DataTable table in dataSet.Tables)
                {
                    jsonString += "\"" + table.TableName + "\":" + Dtb2Json(table) + ",";
                }
                jsonString = jsonString.TrimEnd(',');
                return jsonString + "}";
            }
            #endregion

            #region 实体类复制

            /// <summary>
            /// 将一个实体类复制到另一个实体类
            /// </summary>
            /// <param name="objectsrc">源实体类</param>
            /// <param name="objectdest">复制到的实体类</param>
            /// <param name="excudeFields">不复制的属性</param>
            public static void EntityToEntity(object objectsrc, object objectdest, params string[] excudeFields)
            {
                var sourceType = objectsrc.GetType();
                var destType = objectdest.GetType();
                foreach (var item in destType.GetProperties())
                {
                    #region 原代码
                    //if (excudeFields.Any(x => x.ToUpper() == item.Name))
                    //    continue;
                    //item.SetValue(objectdest, sourceType.GetProperty(item.ToString().ToLower()).GetValue(objectsrc, null), null);
                    #endregion

                    foreach (var item1 in sourceType.GetProperties())
                    {
                        if (item.Name == item1.Name)
                        {
                            item.SetValue(objectdest, item1.GetValue(objectsrc, null), null);
                            break;
                        }
                    }

                }         

            }

            #endregion
          
        }
        #endregion
    }
}

