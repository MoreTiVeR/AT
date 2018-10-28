using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using AE.Net.Mail;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Mvc;
using Google.Apis.Drive.v2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using GoogleMailAPI.Models;
using AE.Net.Mail;
using MimeKit;
using System.Net.Mime;

namespace GoogleMailAPI.Controllers
{
    public class HomeController : Controller
    {
        //        private static string _mailFrom = string.Empty;
        //        private static string _mailTo = string.Empty;
        //        private static string _senderPass = string.Empty;
        //        private static string _smptHost = string.Empty;

        static string[] Scopes = { GmailService.Scope.GmailSend };
        static string ApplicationName = "Andaman Transfer";
        public ActionResult Index(VMailContentModel objMail)
        {
            UserCredential credential;
            string path = ConfigurationManager.AppSettings["CREDENTIAL_PATH"].ToString();
            string file_name = ConfigurationManager.AppSettings["CREDENTIAL_BOOKING_FILE_NAME"].ToString();
            string valuePath = ConfigurationManager.AppSettings["CREDENTIAL_BOOKING_VALIDATION_PATH"].ToString();
            string fullPath = Path.Combine(path, file_name);

            using (var stream = new FileStream(fullPath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                //string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                string credPath = string.Empty;
                credPath = Path.Combine(path, string.Format(".credentials/{0}", valuePath));
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            #region Send Mail
            if (objMail != null
                && !string.IsNullOrEmpty(objMail.MailTo)
                && !string.IsNullOrEmpty(objMail.MailSubject)
                && !string.IsNullOrEmpty(objMail.MailBody))
            {

                objMail.MailBody = @"<div id=':w8' class='a3s aXjCH ' role='gridcell' tabindex='-1'><b>Dear Joel Sandeep DSouza</b><br><br><p><font size='6'>Your booking has been confirmed.</font></p><br><b>Your Booking Details</b><br><b>Passeger Name</b>: Joel Sandeep DSouza<br><b>Booking Number</b>: AT-181017-1213-15308<br><b>Arrival DateTime</b>: 22/10/2018 09:30<br><b>Arrival Pickup From</b>: Krabi Airport<br><b>Arrival Destination</b>: Ao Nang/Noppharattara<br><b>Return DateTime</b>: <br><b>Return Pickup From</b>: Aonang Cliff Beach Resort, Krabi<br><b>Return Destination</b>: Krabi Airport<br><br>→ Site : <a href='https://www.andamantransfer.com' target='_blank' data-saferedirecturl='https://www.google.com/url?q=https://www.andamantransfer.com&amp;source=gmail&amp;ust=1540129523878000&amp;usg=AFQjCNGlyfUkOsiaICy-shlxfh67DTPTWg'>Click Here<u></u></a><br>→ Email : <a href='mailto:andamantransfer.th@gmail.com' target='_blank'>andamantransfer.th@gmail.com</a><br>→ Tel : +(66) 92-181-9997<br><br>Best Regards,<br>Andamantransfer.com<div class='yj6qo'></div><div class='adL'><br></div></div>";

                string plainText = string.Format("To: {0}\r\n" +
                                   "Subject: {1}\r\n" +
                                   "Content-Type: text/html; charset=utf-8\r\n\r\n" +
                                   "{2}", objMail.MailTo, objMail.MailSubject, objMail.MailBody);

                var newMsg = new Google.Apis.Gmail.v1.Data.Message();
                newMsg.Raw = Base64UrlEncode(plainText.ToString());
                service.Users.Messages.Send(newMsg, "me").Execute();

                //string fileName = @"C:\Users\Nattapong\Downloads\AT180901235418520_DRIVER.pdf";
                //var bytedd = System.IO.File.ReadAllBytes(fileName);
                //using (MemoryStream ms = new MemoryStream(bytedd))
                //{
                //    var contentType = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Application.Pdf);
                //    var createRequest = service.Users.Messages.Send(newMsg, "me", ms, @"message/rfc822").Upload();
                //    if(createRequest.Status == Google.Apis.Upload.UploadStatus.Completed)
                //    {
                //        service.Users.Messages.Send(newMsg, "me").Execute();
                //    }
                //    ms.Close();
                //}
            }

            #endregion

            #region Miekite
            StringBuilder strDetail = new StringBuilder();
            strDetail.Append("<b>Dear Nattapong Engchaun</b><br /><br />");
            strDetail.Append("<p><font size='6'>Your booking has been confirmed.</font></p><br />");
            strDetail.Append("<b>Your Booking Details</b><br />");
            strDetail.Append("<b>Passeger Name</b>: Nattapong Engchaun<br />");
            strDetail.Append("<b>Booking Number</b>: AT-181025-1459-36521<br />");
            strDetail.Append("<b>Arrival DateTime</b>: 20/12/2018 10:30<br />");
            strDetail.Append("<b>Arrival Pickup From</b>: Phuket -Airport<br />");
            strDetail.Append("<b>Arrival Destination</b>: Ao Nang/Noppharattara<br />");
            strDetail.Append("<b>Return DateTime</b>: -<br />");
            strDetail.Append("<b>Return Pickup From</b>: -<br />");
            strDetail.Append("<b>Return Destination</b>: -<br /><br />");
            strDetail.Append("<b>Site</b> : <a href=\"www.andamantransfer.com\">Click Here</ID></a><br/>");
            strDetail.Append("<b>Email</b> : booking@andamantransfer.com<br />");
            strDetail.Append("<b>Tel</b> : +(66) 92-181-9997<br /><br />");
            strDetail.Append("Best Regards,<br />");
            strDetail.Append("Andamantransfer.com<br />");

            var msg = new System.Net.Mail.MailMessage();
            msg.From = new MailAddress("booking@andamantransfer.com", "Andaman Transfer");
            msg.To.Add("peeakka@hotmail.com");
            msg.CC.Add("eng.nattapong@gmail.com");
            msg.Subject = "Fridbert Hanke (Booking No : AT-181024-1620-40231) 25 Oct 2018 ➨ Krabi Airport - Ao Nang/Noppharattara Travel Itinerary";
            //msg.BodyEncoding = System.Text.Encoding.Unicode;
            //msg.Body = "<div role='gridcell' tabindex='-1'><b>Dear Joel Sandeep DSouza</b><br><br><p><font size='6'>Your booking has been confirmed.</font></p><br><b>Your Booking Details</b><br><b>Passeger Name</b>: Joel Sandeep DSouza<br><b>Booking Number</b>: AT-181017-1213-15308<br><b>Arrival DateTime</b>: 22/10/2018 09:30<br><b>Arrival Pickup From</b>: Krabi Airport<br><b>Arrival Destination</b>: Ao Nang/Noppharattara<br><b>Return DateTime</b>: <br><b>Return Pickup From</b>: Aonang Cliff Beach Resort, Krabi<br><b>Return Destination</b>: Krabi Airport<br><br>→ Site : <a href='https://www.andamantransfer.com' target='_blank' data-saferedirecturl='https://www.google.com/url?q=https://www.andamantransfer.com&amp;source=gmail&amp;ust=1540129523878000&amp;usg=AFQjCNGlyfUkOsiaICy-shlxfh67DTPTWg'>Click Here<u></u></a><br>→ Email : <a href='mailto:andamantransfer.th@gmail.com' target='_blank'>andamantransfer.th@gmail.com</a><br>→ Tel : +(66) 92-181-9997<br><br>Best Regards,<br>Andamantransfer.com<div class='yj6qo'></div><div class='adL'><br></div></div>";
            msg.Body = strDetail.ToString();
            msg.IsBodyHtml = true;
            //msg.SubjectEncoding = System.Text.Encoding.Unicode;

            List<string> attachFiles = new List<string>();
            string fileName1 = @"C:\Users\Nattapong\Downloads\AT180901235418520_DRIVER.pdf";
            string fileName2 = @"C:\Users\Nattapong\Downloads\AT180902021429519.pdf";
            attachFiles.Add(fileName1);
            attachFiles.Add(fileName2);

            int counter = 0;
            MemoryStream msStream = null;
            attachFiles.ForEach(file =>
            {
                counter++;
                var contentType = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Application.Pdf);
                //using (MemoryStream ms = new MemoryStream())
                //{
                //    msg.Attachments.Add(new System.Net.Mail.Attachment(ms, string.Format("File_Name_{0}", counter), contentType.MediaType));
                //    ms.Close();
                //}
                msStream = new MemoryStream(System.IO.File.ReadAllBytes(file));
                System.Net.Mail.Attachment data = new System.Net.Mail.Attachment(file, MediaTypeNames.Application.Octet);
                // Add time stamp information for the file.
                data.ContentDisposition.CreationDate = System.IO.File.GetCreationTime(file);
                data.ContentDisposition.ModificationDate = System.IO.File.GetLastWriteTime(file);
                data.ContentDisposition.ReadDate = System.IO.File.GetLastAccessTime(file);
                // Add the file attachment to this email message.

                msg.Attachments.Add(data);
                //msg.Attachments.Add(new System.Net.Mail.Attachment(msStream, string.Format("File_Name_{0}", counter), contentType.MediaType));

                //msg.Attachments.Add(new System.Net.Mail.Attachment(new MemoryStream(System.IO.File.ReadAllBytes(file))
                //{ }, string.Format("File_Name_{0}", counter), contentType.MediaType));

            });

            var MimeMsg = MimeKit.MimeMessage.CreateFromMailMessage(msg);
            msStream.Close();

            var encodedText = Base64UrlEncode(MimeMsg.ToString());
            var resSend = service.Users.Messages.Send(new Google.Apis.Gmail.v1.Data.Message { Raw = encodedText }, "me").Execute();


            //var bytedd1 = System.IO.File.ReadAllBytes(fileName1);
            //var bytedd2 = System.IO.File.ReadAllBytes(fileName2);
            //using (MemoryStream ms = new MemoryStream())
            //{
            //    var contentType = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Application.Pdf);
            //    msg.Attachments.Add(new System.Net.Mail.Attachment(ms, "File Name", contentType.MediaType));

            //    var MimeMsg = MimeKit.MimeMessage.CreateFromMailMessage(msg);
            //    var encodedText = Base64UrlEncode(MimeMsg.ToString());
            //    var resSend = service.Users.Messages.Send(new Google.Apis.Gmail.v1.Data.Message { Raw = encodedText }, "me").Execute();

            //    ms.Close();
            //}
            //using (MemoryStream ms = new MemoryStream(bytedd2))
            //{
            //    var contentType = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Application.Pdf);
            //    msg.Attachments.Add(new System.Net.Mail.Attachment(ms, "File Name", contentType.MediaType));
            //    ms.Close();
            //}





            //var message = new MimeMessage();
            //message.From.Add(new MailboxAddress("Joey", "booking@andamantransfer.com"));
            //message.To.Add(new MailboxAddress("Alice", "eng.nattapong@gmail.com"));
            //message.Subject = "How you doin?";

            //var builder = new BodyBuilder();

            //// Set the plain-text version of the message text
            //builder.TextBody = @"Hey Alice,What are you up to this weekend? Monica is throwing one of her parties on Saturday and I was hoping you could make it. Will you be my +1 -- Joey";

            //// We may also want to attach a calendar event for Monica's party...
            //builder.Attachments.Add(@"C:\Users\Nattapong\Downloads\AT180901235418520_DRIVER.pdf");

            //// Now we just need to set the message body and we're done
            //message.Body = builder.ToMessageBody();
            #endregion

            #region AE Mail
            ////var msg = new MailMessage();
            ////msg.From = new MailAddress(mailFrom);
            ////msg.To.Add(mailTo);
            ////msg.Body = "<h1>ทดสอบ ส่งเมลอัตโนมัติ by www.andamantransfer.com | Andaman Transfer Co.,Ltd.</h1>";
            ////msg.IsBodyHtml = true;
            ////msg.BodyEncoding = System.Text.Encoding.GetEncoding("utf-8");
            ////msg.Subject = "Mail From Booking Andamantransfer.com";
            ////msg.Attachments.Add(new Attachment(@"C:\Users\Nattapong\Downloads\AT180901235418520_DRIVER.pdf"));

            //var msg = new AE.Net.Mail.MailMessage
            //{
            //    Subject = "Mailing From Booking Andamantransfer.com",
            //    Body = "Hello, World, from Gmail API!<br><h1>ทดสอบ ส่งเมลอัตโนมัติ by www.andamantransfer.com | Andaman Transfer Co.,Ltd.</h1>",
            //    From = new MailAddress("booking@andamantransfer.com"),
            //    //ContentType = "text/html; charset=utf-8"

            //};
            //msg.To.Add(new MailAddress("eng.nattapong@gmail.com"));

            //var firstAttachmentContents = @"C:\Users\Nattapong\Downloads\AT180901235418520_DRIVER.pdf";
            //var attachment = new AE.Net.Mail.Attachment
            //{
            //    Body = Convert.ToBase64String(Encoding.Default.GetBytes(firstAttachmentContents)),
            //    ContentTransferEncoding = "base64",
            //    Encoding = Encoding.ASCII
            //};
            //attachment.Headers.Add("Content-Type", new HeaderValue(@"application/pdf; filename=""AT180901235418520_DRIVER.pdf"""));
            //msg.Attachments.Add(attachment);

            ////string fileName = @"C:\Users\Nattapong\Downloads\AT180901235418520_DRIVER.pdf";
            ////FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            ////byte[] attachByte = new byte[fs.Length];
            ////AE.Net.Mail.Attachment file = new AE.Net.Mail.Attachment(attachByte, fileName, fileName, true);
            ////msg.Attachments.Add(file);

            //var msgStr = new StringWriter();
            //msg.Save(msgStr);

            //var newMsg = new Google.Apis.Gmail.v1.Data.Message
            //{
            //    Raw = Base64UrlEncode(msgStr.ToString())
            //};

            //service.Users.Messages.Send(newMsg, "me").Execute();
            #endregion


            return View();
        }

        private static string GetMimeType(string fileName)
        {
            string mimeType = "application/unknown";
            string ext = Path.GetExtension(fileName).ToLower();
            Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
            if (regKey != null && regKey.GetValue("Content Type") != null)
                mimeType = regKey.GetValue("Content Type").ToString();
            return mimeType;
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public static string Base64UrlEncode(string input)
        {
            var inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(inputBytes).Replace("+", "-").Replace("/", "_").Replace("=", "");
        }
    }
}