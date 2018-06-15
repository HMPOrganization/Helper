using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Helper.WriteLog
{
    public class WriteLog
    {

        public static void WriteLog1(string strLog)
        {
            string sFilePath = System.Environment.CurrentDirectory + "\\72Craft\\" + DateTime.Now.ToString("yyyyMM");
            string sFileName = "operator" + DateTime.Now.ToString("dd") + ".log";

            sFileName = sFilePath + "\\" + sFileName; //文件的绝对路径


            //sFileName = "C:\\1.log";
            if (!Directory.Exists(sFilePath))//验证路径是否存在
            {
                Directory.CreateDirectory(sFilePath);
                //不存在则创建
            }
            FileStream fs;
            StreamWriter sw;
            if (File.Exists(sFileName))
            //验证文件是否存在，有则追加，无则创建
            {
                fs = new FileStream(sFileName, FileMode.Append, FileAccess.Write);
            }
            else
            {
                fs = new FileStream(sFileName, FileMode.Create, FileAccess.Write);
            }
            sw = new StreamWriter(fs);
            sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + "  ------   " + strLog);

            sw.Close();
            fs.Close();
        }
    }
}
