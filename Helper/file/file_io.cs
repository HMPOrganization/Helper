using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Helper.file
{
    public class file_io
    {

        /// <summary>
        /// 递归拷贝所有子目录。
        /// </summary>
        /// <param >源目录</param>
        /// <param >目的目录</param>
        public static void copyDirectory(string sPath, string dPath)
        {
            string[] directories = System.IO.Directory.GetDirectories(sPath);
            if (!System.IO.Directory.Exists(dPath))
                System.IO.Directory.CreateDirectory(dPath);
            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(sPath);
            System.IO.DirectoryInfo[] dirs = dir.GetDirectories();
            CopyFile(dir, dPath);
            if (dirs.Length > 0)
            {
                foreach (System.IO.DirectoryInfo temDirectoryInfo in dirs)
                {
                    string sourceDirectoryFullName = temDirectoryInfo.FullName;
                    string destDirectoryFullName = sourceDirectoryFullName.Replace(sPath, dPath);
                    if (!System.IO.Directory.Exists(destDirectoryFullName))
                    {
                        System.IO.Directory.CreateDirectory(destDirectoryFullName);
                    }
                    CopyFile(temDirectoryInfo, destDirectoryFullName);
                    copyDirectory(sourceDirectoryFullName, destDirectoryFullName);
                    
                }
            }
            //System.IO.

        }

        /// <summary>
        /// 拷贝目录下的所有文件到目的目录。
        /// </summary>
        /// <param name="path">源路径</param>
        /// <param name="desPath">目的路径</param>
        private static void CopyFile(System.IO.DirectoryInfo path, string desPath)
        {
            string sourcePath = path.FullName;
            System.IO.FileInfo[] files = path.GetFiles();
            foreach (System.IO.FileInfo file in files)
            {
                string sourceFileFullName = file.FullName;
                string destFileFullName = sourceFileFullName.Replace(sourcePath, desPath);
                file.CopyTo(destFileFullName, true);
                
            }
        }

        /// <summary>
        /// 删除指定路径下的文件夹和文件
        /// </summary>
        /// <param name="yourPath">路径</param>
        public static void DeleteFile(string yourPath)
        {
            if (Directory.Exists(yourPath))
            {
                //获取指定路径下所有文件夹
                string[] folderPaths = Directory.GetDirectories(yourPath);

                foreach (string folderPath in folderPaths)
                    Directory.Delete(folderPath, true);
                //获取指定路径下所有文件
                string[] filePaths = Directory.GetFiles(yourPath);

                foreach (string filePath in filePaths)
                    File.Delete(filePath);
            }
        }

        /// <summary>
        /// 删除指定的文件夹
        /// </summary>
        /// <param name="yourPath">路径</param>
        public static void DeleteDierctory(string yourPath)
        {
            if (Directory.Exists(yourPath))
                Directory.Delete(yourPath, true);
        }

        public static void AddCreateDirectory(string yourPath)
        {
            Directory.CreateDirectory(@yourPath);
        }

        /// <summary>
        /// 复制文件
        /// </summary>
        /// <param name="fromFile">要复制的文件</param>
        /// <param name="toFile">要保存的位置</param>
        /// <param name="lengthEachTime">每次复制的长度</param>
        public static void CopyFile(string fromFile, string toFile)//, int lengthEachTime)
        {
            try
            {

        
            FileStream fileToCopy = new FileStream(fromFile, FileMode.Open, FileAccess.Read);
            FileStream copyToFile = new FileStream(toFile, FileMode.Create, FileAccess.Write);
            //int lengthToCopy;
            //if (lengthEachTime < fileToCopy.Length)//如果分段拷贝，即每次拷贝内容小于文件总长度
            //{
            //    byte[] buffer = new byte[lengthEachTime];
            //    int copied = 0;
            //    while (copied <= ((int)fileToCopy.Length - lengthEachTime))//拷贝主体部分
            //    {
            //        lengthToCopy = fileToCopy.Read(buffer, 0, lengthEachTime);
            //        fileToCopy.Flush();
            //        copyToFile.Write(buffer, 0, lengthEachTime);
            //        copyToFile.Flush();
            //        copyToFile.Position = fileToCopy.Position;
            //        copied += lengthToCopy;
            //    }
            //    int left = (int)fileToCopy.Length - copied;//拷贝剩余部分
            //    lengthToCopy = fileToCopy.Read(buffer, 0, left);
            //    fileToCopy.Flush();
            //    copyToFile.Write(buffer, 0, left);
            //    copyToFile.Flush();
            //}
            //else//如果整体拷贝，即每次拷贝内容大于文件总长度
            //{
                byte[] buffer = new byte[fileToCopy.Length];
                fileToCopy.Read(buffer, 0, (int)fileToCopy.Length);
                fileToCopy.Flush();
                copyToFile.Write(buffer, 0, (int)fileToCopy.Length);
                copyToFile.Flush();
            //}
            fileToCopy.Close();
            copyToFile.Close();
            }
            catch (Exception )
            {

                throw;
            }
        }
    }
}
