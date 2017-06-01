#region FileHeader
/********************************************************************************
** Copyright (C) 2013 Newegg. All rights reserved.
**
**
** File Name:           Process
** Creator:             ez07
** Create date:         9/2/2013 1:26:54 PM        
** CLR Version:         4.0.30319.17929
** NameSpace:           $projectname$ 
** Description:
** Latest Modifier:
** Latest Modify date:     
**
**
** Version number:      1.0.0.0
*********************************************************************************/
#endregion

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Windows.Forms;


namespace LoadMail
{
    public class Process
    {
        public static string url = "http://localhost:3000/api/open/task/outlook/";
        public static void ToAPI(Microsoft.Office.Interop.Outlook.MailItem item)
        {
            bool InsertFlg = false;
            string error = string.Empty;
            try
            {
                string subject = item.Subject;
                string body = item.Body;
                string bodyHtml = item.HTMLBody;

                Encoding myEncoding = Encoding.UTF8;
                //string param = HttpUtility.UrlEncode("subject", myEncoding) + "=" + HttpUtility.UrlEncode(subject, myEncoding)
                //    + "&" + HttpUtility.UrlEncode("body", myEncoding) + "=" + HttpUtility.UrlEncode(bodyHtml, myEncoding);

                string param = "subject=" + Microsoft.JScript.GlobalObject.encodeURIComponent(subject) + "&body=" + Microsoft.JScript.GlobalObject.encodeURIComponent(bodyHtml);

                byte[] postBytes = Encoding.UTF8.GetBytes(param);

                HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded;charset=gb2312";
                request.ContentLength = postBytes.Length;

                using (Stream reqStream = request.GetRequestStream())
                {
                    reqStream.Write(postBytes, 0, postBytes.Length);
                }

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                using (Stream responseStream = response.GetResponseStream())
                {
                    //在这里对接收到的页面内容进行处理
                    string source = new StreamReader(responseStream, myEncoding).ReadToEnd();
                    if (source == "success")
                        InsertFlg = true;
                    else
                        error = source;

                }



                if (InsertFlg)
                {
                    MessageBox.Show("已成功导入 {" + item.Subject + "} 至PPManager系统", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("导入失败:" + error, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("导入PPManager系统时出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static bool BatchToAPI(Microsoft.Office.Interop.Outlook.MailItem item,out string error)
        {
            bool InsertFlg = false;
            try
            {
                string subject = item.Subject;
                string body = item.Body;
                string bodyHtml = item.HTMLBody;

                Encoding myEncoding = Encoding.UTF8;

                
                //string param = HttpUtility.UrlEncode("subject", myEncoding) + "=" + HttpUtility.UrlEncode(subject, myEncoding)
                //    + "&" + HttpUtility.UrlEncode("body", myEncoding) + "=" + HttpUtility.UrlEncode(bodyHtml, myEncoding);

                string param = "subject=" + Microsoft.JScript.GlobalObject.encodeURIComponent(subject) + "&body=" + Microsoft.JScript.GlobalObject.encodeURIComponent(bodyHtml);

                byte[] postBytes = Encoding.UTF8.GetBytes(param);

                HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded;charset=gb2312";
                request.ContentLength = postBytes.Length;

                using (Stream reqStream = request.GetRequestStream())
                {
                    reqStream.Write(postBytes, 0, postBytes.Length);
                }

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                using (Stream responseStream = response.GetResponseStream())
                {
                    //在这里对接收到的页面内容进行处理
                    string source = new StreamReader(responseStream, myEncoding).ReadToEnd();
                    if (source == "success")
                    {
                        error = "";
                        return true;
                    }
                    else
                    {
                        error = source;
                        return false;
                    }

                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }
    }
}