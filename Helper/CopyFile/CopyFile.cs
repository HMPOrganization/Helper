using Helper.Decrypt;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace Helper.CopyFile
{
    /// <summary>
    /// 复制文件工具类
    /// </summary>
   public class CopyFile
    {

        #region 复制文件夹中的所有指定内容到指定目录并备份

        /// <summary>
        /// 复制文件夹中的所有内容
        ///<!--图片服务器可支持类型-->
        ///<add key = "imgtype" value="Eb9hkLwmRSbtQONpOwRsgUOWpq0AH3suP9pAl9UzD1E=" />
        ///<!--图片服务器源路径-->
        ///<add key = "SourceDirectory" value="NYOgCGQALKWbZschXyBRp0zNqMxonMOSzW4wSd9oqgztkzbIo44YRmXY8kGB+EO9" />
        ///<!--图片服务器源路径用户名-->
        ///<add key = "SourceUserName" value="9ZwjbCk/prqh6911XEkS1A==" />
        ///<!--图片服务器源路径密码-->
        ///<add key = "SourcePassWord" value="0S6j9HpcoU6S+cxaeOpELg==" />
        ///<!--图片服务器目标路径-->
        ///<add key = "TargetDirectory" value="kfDslHiD/rod/T/Fy+/EHIEU8e5vX6m/bor5x0TyxWPk+0iDQZ4IeN14jtq2H4to" />
        ///<!--图片服务器目标路径用户名-->
        ///<add key = "TargetUserName" value="9ZwjbCk/prqh6911XEkS1A==" />
        ///<!--图片服务器目标路径用密码-->
        ///<add key = "TargetPassWord" value="0S6j9HpcoU6S+cxaeOpELg==" />
        ///<!--图片服务器备份路径-->
        ///<add key = "BackupDirectory" value="ZsrlG20TkntrfkO3aFrxtwNJl4UlO/cmMyADI5cqlvQ=" />
        ///<!--图片备份路径用户名-->
        ///<add key = "BackupUserName" value="3c8dJ7js9MzXlV6s110+brtuflPlfRDeatulp1w507U=" />
        ///<!--图片备份路径用密码-->
        ///<add key = "BackupPassWord" value="xaiFOTRzJQGOUhVHj3/5Cg==" />
        /// </summary>
        /// <param name="sourceDirPath">源文件夹目录</param>
        /// <param name="source_username">源文件夹目录用户名（密文字符串）</param>
        /// <param name="source_userpwd">源文件夹目录密码（密文字符串）</param>
        /// <param name="saveDirPath">指定文件夹目录</param>
        /// <param name="sourceDirPath_state">输入路径的状态（默认为true = 文件 可选为false = 路径）  </param>
        /// <param name="BackupDirectory">备份文件夹目录（可选）</param>
        /// <param name="sourceDirPath_key">源文件夹目录登陆信息（可选） </param>
        /// <param name="saveDirPath_key">指定文件夹目录登陆信息（可选）</param>
        /// <param name="BackupDirectory_key">备份文件夹目录登陆信息（可选）</param>
        /// <returns></returns>
        public string CopyDirectory(string sourceDirPath, string saveDirPath, string source_username="", string source_userpwd="", bool sourceDirPath_state = true, string BackupDirectory = "", bool sourceDirPath_key = false, bool saveDirPath_key = false, bool BackupDirectory_key = false)
        {
            try
            {
                
                Decrypt_out decrypt_out = new Helper.Decrypt.Decrypt_out();
                List<string> files_list = new List<string>();
                sourceDirPath = sourceDirPath.Replace("file:", "");
                #region 判断输入的是路径还是文件
                if (sourceDirPath_state == false)
                {
                  
                    if (sourceDirPath_key != false)
                    {
                        bool status = false;

                        //连接共享文件夹
                        status = connectState(@sourceDirPath, source_username, source_userpwd);
                        //status = connectState(@"\\172.16.1.2\MESfiles", "highrock\\mes_pic", "123456");
                        if (!status)
                        {
                            return "源文件夹用户名密码连接失败！";
                        }

                    }
                    if (!Directory.Exists(saveDirPath))
                    {
                        return "保存目标路径不存在！！";
                    }



                    string imgtype = decrypt_out.Decrypt_out_Str("imgtype");
                    //string imgtype = "*.MP3|*.JPG|*.GIF|*.PNG|*.PDF";
                    string[] ImageType = imgtype.Split('|');



                    for (int i = 0; i < ImageType.Length; i++)
                    {
                        string[] files = Directory.GetFiles(sourceDirPath, ImageType[i]);

                        foreach (string item in files)
                        {
                            files_list.Add(item);
                        }
                    }

                }
                else
                {

                    if (sourceDirPath_key != false)
                    {
                        bool status = false;

                        //连接共享文件夹
                        status = connectState(@sourceDirPath, source_username, source_userpwd);
                        //status = connectState(@"\\172.16.1.2\MESfiles", "highrock\\mes_pic", "123456");
                        if (!status)
                        {
                            return "源文件夹用户名密码连接失败！";
                        }

                    }
                    if (!Directory.Exists(saveDirPath))
                    {
                        return "保存目标路径不存在！！";
                    }

                    string imgtype = decrypt_out.Decrypt_out_Str("imgtype");
                    //string imgtype = "*.MP3|*.JPG|*.GIF|*.PNG|*.PDF";
                    string[] ImageType = imgtype.Split('|');



                    for (int i = 0; i < ImageType.Length; i++)
                    {
                        string P_str_fileexc = sourceDirPath.Substring(sourceDirPath.LastIndexOf(".") + 1,sourceDirPath.Length - sourceDirPath.LastIndexOf(".") - 1);
                        P_str_fileexc = "*." + P_str_fileexc;

                        if (ImageType[i] == P_str_fileexc.ToUpper())
                        {
                            files_list.Add(sourceDirPath);
                        }
                    }
                }
                #endregion


                foreach (string file in files_list)
                {
               
                    if (saveDirPath_key != false)
                    {
                        bool status = false;

                        //连接共享文件夹
                        status = connectState(@decrypt_out.Decrypt_out_Str("TargetDirectory"), decrypt_out.Decrypt_out_Str("TargetUserName"), decrypt_out.Decrypt_out_Str("TargetPassWord"));
                        //status = connectState(@"\\172.16.1.2\MESfiles", "highrock\\mes_pic", "123456");
                        if (!status)
                        {
                            return "目标文件夹用户名密码连接失败！";
                        }

                    }
                    string pFilePath = saveDirPath + "\\" + Path.GetFileName(file.Trim());
                    //if (File.Exists(pFilePath))
                    //{
                    //    continue;
                    //}

                    File.Copy(file.Replace("file:", "").Trim(), pFilePath, true);

                    if (BackupDirectory != "")
                    {
                        if (BackupDirectory_key != false)
                        {
                            bool status = false;

                            //连接共享文件夹
                            status = connectState(@decrypt_out.Decrypt_out_Str("BackupDirectory"), decrypt_out.Decrypt_out_Str("BackupUserName"), decrypt_out.Decrypt_out_Str("BackupPassWord"));
                            //status = connectState(@"\\172.16.1.2\MESfiles", "highrock\\mes_pic", "123456");
                            if (!status)
                            {
                                return "备份文件夹用户名密码连接失败！";
                            }

                        }

                        string BFilePath = BackupDirectory + "\\" + Path.GetFileName(file.Trim());
                        File.Copy(file.Replace("file:", "").Trim(), BFilePath, true);
                    }
                }

                //string[] dirs = Directory.GetDirectories(sourceDirPath);
                //foreach (string dir in dirs)
                //{
                //    CopyDirectory(dir, saveDirPath + "\\" + Path.GetFileName(dir));
                //}
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return "ok";
        }


        #endregion

        #region 连接远程共享文件夹

        /// <summary>
        /// 连接远程共享文件夹
        /// </summary>
        /// <param name="path">远程共享文件夹的路径</param>
        /// <param name="userName">用户名</param>
        /// <param name="passWord">密码</param>
        /// <returns></returns>
        public static bool connectState(string path, string userName, string passWord)
        {
            CopyFile copyfile = new CopyFile();

            path = copyfile.string_Route(path);
            bool Flag = false;
            Process proc = new Process();
            try
            {
                proc.StartInfo.FileName = "cmd.exe";
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardInput = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.CreateNoWindow = true;
                proc.Start();
                string dosLine = @"net use " + path + " /user:" + userName +" "+ passWord;
                proc.StandardInput.WriteLine(dosLine);
                proc.StandardInput.WriteLine("exit");
                while (!proc.HasExited)
                {
                    proc.WaitForExit(1000);
                }
                string errormsg = proc.StandardError.ReadToEnd();
                proc.StandardError.Close();
                if (string.IsNullOrEmpty(errormsg))
                {
                    Flag = true;
                }
                else
                {
                    throw new Exception(errormsg);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                proc.Close();
                proc.Dispose();
            }
            return Flag;
        }

        #endregion

        #region 遍历文件夹所有目录和子目录
        /// <summary>
        /// 遍历文件夹所有目录和子目录（需要配置返回文件类型 例： <add key="QueryRailPassword" value="*.MP3|*.JPG|*.GIF|*.PNG|*.PDF"></add>）
        /// </summary>
        /// <param name="sSourcePath">源目录</param>
        /// <returns>返回路径的List_string</returns>
        public List<string> FindFile2(string sSourcePath)

        {

            Decrypt_out decrypt_out = new Helper.Decrypt.Decrypt_out();

            string imgtype = decrypt_out.Decrypt_out_Str("imgtype");
            //string imgtype = "*.MP3|*.JPG|*.GIF|*.PNG|*.PDF";
            string[] ImageType = imgtype.Split('|');

            ///_______________________________________________________________________
            List<String> list = new List<string>();

            //遍历文件夹

            DirectoryInfo theFolder = new DirectoryInfo(sSourcePath);

            FileInfo[] thefileInfo = theFolder.GetFiles("*.*", SearchOption.TopDirectoryOnly);

            foreach (FileInfo NextFile in thefileInfo)  //遍历文件
            {
                for (int i = 0; i < ImageType.Length; i++)
                {
                    if (NextFile.Extension == ImageType[i])
                    {
                        list.Add(NextFile.FullName);
                        continue;
                    }
                }
            }

            //遍历子文件夹

            DirectoryInfo[] dirInfo = theFolder.GetDirectories();

            foreach (DirectoryInfo NextFolder in dirInfo)

            {

                //list.Add(NextFolder.ToString());

                FileInfo[] fileInfo = NextFolder.GetFiles("*.*", SearchOption.AllDirectories);

                foreach (FileInfo NextFile in fileInfo)  //遍历文件
                {
                    for (int i = 0; i < ImageType.Length; i++)
                    {
                        if (NextFile.Extension == ImageType[i] && NextFile.CreationTime.Date == DateTime.Now.Date)
                        {
                            list.Add(NextFile.FullName);
                            continue;
                        }
                    }
                }


            }

            return list;
        }
        #endregion

        #region 监视某个文件夹的文件情况

        /// <summary>
        /// 监视某个文件夹的文件情况
        /// </summary>
        /// <param name="path">监视路径</param>
        /// <param name="filter">监视的文件类型 默认 "*.*"</param>
        public void WatcherStrat(string path, string filter)
        {

            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = path;
            watcher.Filter = filter;
            watcher.Changed += new FileSystemEventHandler(OnProcess);
            watcher.Created += new FileSystemEventHandler(OnProcess);
            watcher.Deleted += new FileSystemEventHandler(OnProcess);
            watcher.Renamed += new RenamedEventHandler(OnRenamed);
            watcher.EnableRaisingEvents = true;
            watcher.NotifyFilter = NotifyFilters.Attributes | NotifyFilters.CreationTime | NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.LastAccess
                                   | NotifyFilters.LastWrite | NotifyFilters.Security | NotifyFilters.Size;
            watcher.IncludeSubdirectories = true;
        }

        private static void OnProcess(object source, FileSystemEventArgs e)
        {
            if (e.ChangeType == WatcherChangeTypes.Created)
            {
                OnCreated(source, e);
            }
            else if (e.ChangeType == WatcherChangeTypes.Changed)
            {
                OnChanged(source, e);
            }
            else if (e.ChangeType == WatcherChangeTypes.Deleted)
            {
                OnDeleted(source, e);
            }
            else if (e.ChangeType == WatcherChangeTypes.Renamed)
            {
                OnRenamed(source, e);
            }

        }
        /// <summary>
        /// 监听新建文件事件
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        private static void OnCreated(object source, FileSystemEventArgs e)
        {
            Helper.Decrypt.Decrypt_out decrypt_out = new Helper.Decrypt.Decrypt_out();

            //string SourceDirectory = decrypt_out.Decrypt_out_Str("SourceDirectory"); //源目录
            string TargetDirectory = decrypt_out.Decrypt_out_Str("TargetDirectory"); //目标目录
            string BackupDirectory = decrypt_out.Decrypt_out_Str("BackupDirectory"); //备份目录

            Helper.CopyFile.CopyFile copyfile = new Helper.CopyFile.CopyFile();


            copyfile.CopyDirectory(e.FullPath, TargetDirectory, sourceDirPath_state: true, BackupDirectory: BackupDirectory, BackupDirectory_key: true, saveDirPath_key: true);

        }
        /// <summary>
        /// 监听改变文件事件
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        private static void OnChanged(object source, FileSystemEventArgs e)
        {
            Helper.Decrypt.Decrypt_out decrypt_out = new Helper.Decrypt.Decrypt_out();

            //string SourceDirectory = decrypt_out.Decrypt_out_Str("SourceDirectory"); //源目录
            string TargetDirectory = decrypt_out.Decrypt_out_Str("TargetDirectory"); //目标目录
            string BackupDirectory = decrypt_out.Decrypt_out_Str("BackupDirectory"); //备份目录

            Helper.CopyFile.CopyFile copyfile = new Helper.CopyFile.CopyFile();


            copyfile.CopyDirectory(e.FullPath, TargetDirectory, sourceDirPath_state: true, BackupDirectory: BackupDirectory, BackupDirectory_key: true, saveDirPath_key: true);

        }

        /// <summary>
        /// 监听删除文件事件
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        private static void OnDeleted(object source, FileSystemEventArgs e)
        {
            Console.WriteLine("文件删除事件处理逻辑{0}  {1}   {2}", e.ChangeType, e.FullPath, e.Name);
        }
        /// <summary>
        /// 监听重命名文件事件
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        private static void OnRenamed(object source, FileSystemEventArgs e)
        {
            Helper.Decrypt.Decrypt_out decrypt_out = new Helper.Decrypt.Decrypt_out();

            //string SourceDirectory = decrypt_out.Decrypt_out_Str("SourceDirectory"); //源目录
            string TargetDirectory = decrypt_out.Decrypt_out_Str("TargetDirectory"); //目标目录
            string BackupDirectory = decrypt_out.Decrypt_out_Str("BackupDirectory"); //备份目录

            Helper.CopyFile.CopyFile copyfile = new Helper.CopyFile.CopyFile();


            copyfile.CopyDirectory(e.FullPath, TargetDirectory, sourceDirPath_state: true, BackupDirectory: BackupDirectory, BackupDirectory_key: true, saveDirPath_key: true);

        }

        #endregion

        #region 从全路径string中得到路径

        /// <summary>
        /// 从全路径string中得到路径
        /// </summary>
        /// <param name="FullPath">全路径</param>
        /// <returns>路径</returns>
        public string string_Route(string FullPath)
        {
            string string_Route = FullPath.Substring(0, FullPath.LastIndexOf("\\"));

            return string_Route;
        }

        #endregion

        #region 从全路径string中得到文件名

        /// <summary>
        /// 从全路径string中得到文件名
        /// </summary>
        /// <param name="FullPath">全路径</param>
        /// <returns>文件名</returns>
        public string string_filename(string FullPath)
        {
            string P_str_filename =
                    FullPath.Substring(FullPath.LastIndexOf("\\") + 1,
                    FullPath.LastIndexOf(".") - (FullPath.LastIndexOf("\\") + 1));

            return P_str_filename;
        }

        #endregion

        #region 从全路径string中得到扩展名

        /// <summary>
        /// 从全路径string中得到扩展名
        /// </summary>
        /// <param name="FullPath">全路径</param>
        /// <returns>扩展名</returns>
        public string string_fileexc(string FullPath)
        {
            string P_str_fileexc =
                    FullPath.Substring(FullPath.LastIndexOf(".") + 1,
                    FullPath.Length - FullPath.LastIndexOf(".") - 1);

            return P_str_fileexc;
        }

        #endregion

        #region 判断字符串是目录还是文件

        /// <summary>
        /// 判断字符串是目录还是文件
        /// </summary>
        /// <param name="str">需要判断的字符串</param>
        /// <returns>返回 true = '目录' false = '文件'  </returns>
        public bool string_full(string str)
        {
            bool b;
            b = Directory.Exists(str);

            return b;
        }
        #endregion

    }
}
