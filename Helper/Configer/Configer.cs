using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace Helper.Configer
{
    public class Configer
    { 
        /// <summary>
        /// Configer.configSections配置节点
        /// </summary>
        public class ConfigSection : ConfigurationSection
       {
         /// <summary>
         /// 接口名称
         /// </summary>
         [ConfigurationProperty("Interface_name", DefaultValue = "", IsRequired = true, IsKey = false)]
            public string Interface_name
            {
                get { return (string)this["Interface_name"]; }
                set { this["Interface_name"] = value; }
            }
         /// <summary>
        /// MAIL地址
        /// </summary>
         [ConfigurationProperty("mailing_address", DefaultValue = "", IsRequired = true, IsKey = false)]
            public string mailing_address
            {
                get { return (string)this["mailing_address"]; }
                set { this["mailing_address"] = value; }
            }

        /// <summary>
        /// 发送人MAIL地址
        /// </summary>
        [ConfigurationProperty("Sender", DefaultValue = "", IsRequired = true, IsKey = false)]
            public string Sender
            {
                get { return (string)this["Sender"]; }
                set { this["Sender"] = value; }
            }

        /// <summary>
        /// 发送人密码
        /// </summary>
        [ConfigurationProperty("password", DefaultValue = "", IsRequired = true, IsKey = false)]
            public string password
            {
                get { return (string)this["password"]; }
                set { this["password"] = value; }
            }

        }
        /// <summary>
        /// 读取Config获取邮件信息
        /// </summary>
        public class ConfigManager
        {
            /// <summary>
            /// 配置Config信息实体
            /// </summary>
            public static readonly ConfigSection Instance = GetSection();

            private static ConfigSection GetSection()
            {
                ConfigSection config = ConfigurationManager.GetSection("Mymail") as ConfigSection;
                if (config == null)
                    throw new ConfigurationErrorsException();
                return config;
            }
        }
    }
}
