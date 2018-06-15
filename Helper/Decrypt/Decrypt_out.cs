using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/// <summary>
/// 解密类库
/// </summary>
namespace Helper.Decrypt
{
    /// <summary>
    /// 解密后得到连接字符串
    /// </summary>
    public class Decrypt_out
    {
        #region 解密后得到字符串
        /// <summary>
        /// 得到解密后的字符串
        /// </summary>
        /// <returns>返回解密后的字符串</returns>
        public string Decrypt_out_Str(string AppSettings)
        {



            //Configuration config = System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            ////根据Key读取<connectionString>元素的Value
            //string name = config.AppSettings.Settings["connectionString"].Value;

            //string name=ConfigurationSettings.AppSettings["connectionString"];

            string name = System.Configuration.ConfigurationManager.AppSettings[AppSettings];
            string connstr = Helper.DEncrypt.Security.DecryptDES(name);

            return connstr;

        }
        #endregion
    }
}
