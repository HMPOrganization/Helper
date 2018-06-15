using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;
using System.IO;

/// <summary>
/// 邮件类库
/// </summary>
namespace Helper.mail
{
    /// <summary>
    /// 发送邮件
    /// </summary>
    public class mail
    {
        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="Title_name">邮件标题</param>
        /// <param name="text">邮件正文</param>
        /// <param name="smtp">SMTP设置</param>
        /// <param name="smtp_port">SMTP端口号</param>
        /// <param name="emailAcount">发送端地址</param>
        /// <param name="emailPassword">发送端密码</param>
        /// <param name="mail">接收人地址</param>
        /// <returns></returns>
        public static bool addmail(string Title_name, string text,string smtp, int smtp_port, string emailAcount,string emailPassword,string[] mail, List<string> strFilePath = null, params string[] CC)
        {

            MailMessage message = new MailMessage();
            //设置发件人,发件人需要与设置的邮件发送服务器的邮箱一致
            //MailAddress fromAddr = new MailAddress(emailAcount, senderDisplayName);
            MailAddress fromAddr = new MailAddress(emailAcount);

            message.From = fromAddr;
            //设置收件人,可添加多个,添加方法与下面的一样
            foreach (string item in mail)
            {
                message.To.Add(item);

            }
            //设置抄送人

            if (CC != null)
            {
                foreach (string item in CC)
                {
                    if (item != null)
                    {
                        message.CC.Add(item);
                    }

                }
            }


            //附件
            //string strFilePath = @"E:\logo.jpg";
            if (strFilePath != null)
            {
                foreach (var item in strFilePath)
                {
                   
                    if (File.Exists(item))
                    {
                        //File.Copy(item + "new", item);       

                        System.Net.Mail.Attachment attachment1 = new System.Net.Mail.Attachment(item);//添加附件 
                        attachment1.Name = System.IO.Path.GetFileName(item);
                        attachment1.NameEncoding = System.Text.Encoding.GetEncoding("gb2312");
                        attachment1.TransferEncoding = System.Net.Mime.TransferEncoding.Base64;
                        attachment1.ContentDisposition.Inline = true;
                        attachment1.ContentDisposition.DispositionType = System.Net.Mime.DispositionTypeNames.Inline;
                        string cid = attachment1.ContentId;//关键性的地方，这里得到一个id数值 
                        message.Attachments.Add(attachment1);


                    }


                }
            }



            //设置邮件标题
            message.Subject = Title_name;

            //设置邮件内容
            message.Body = text;

            //设置邮件发送服务器,服务器根据你使用的邮箱而不同,可以到相应的 邮箱管理后台查看,下面是QQ的
            SmtpClient client = new SmtpClient(smtp, smtp_port);
            client.UseDefaultCredentials = true;
            //设置发送人的邮箱账号和密码
            client.Credentials = new System.Net.NetworkCredential(emailAcount, emailPassword);
            //启用ssl,也就是安全发送
            //client.EnableSsl = true;
            //发送邮件
            client.Send(message);
            message.Dispose();
            

            return true;
        }


    }
}
