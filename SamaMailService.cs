using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Net.Sockets;
using System.Net;
using System.Net.Mail;
using System.Data.SqlClient;
using System.Configuration;
using System.Windows;
using OpenPop.Pop3;
using Limilabs.Client.POP3;
using Limilabs.Mail;
using System.IO;
using System.Timers;
using MailKit.Net.Imap;
using MailKit.Security;
using MailKit.Net;
using MailKit.Search;
using MailKit;
using MimeKit;

namespace SamaMailService
{
    public partial class SamaMailService : ServiceBase
    {
        SqlConnection conn;
        private Timer _timer;
        private DateTime _lastRun = DateTime.Now.AddDays(-1);
        private void AddEvent(string msg, EventLogEntryType type= EventLogEntryType.Information)
        {
            using (EventLog eventLog = new EventLog("Application"))
            {
                eventLog.Source = "SamaMailService";
                eventLog.WriteEntry(msg, type, 0, 1);
            }
        }
        public SamaMailService()
        {
            InitializeComponent();
        }
        private string getAction(string actionName)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.CommandText = "SELECT Value FROM V_CMN_ControlActions where ActionNAme=@ActionName";
            cmd.CommandType = CommandType.Text;
            cmd.Parameters.Add(new SqlParameter("@ActionName", actionName));
            var RAction = cmd.ExecuteReader();
            RAction.Read();
            string result = RAction["Value"].ToString();
            RAction.Close();
            return result;
        }
        private void sendmail()
        {

            SqlCommand cmdSend = new SqlCommand();
            cmdSend.Connection = conn;
            cmdSend.CommandText = getAction("acService_send_email");
            cmdSend.CommandType = CommandType.Text;

            //attachment
            SqlCommand cmdSendAttach = new SqlCommand();
            cmdSendAttach.Connection = conn;
            cmdSendAttach.CommandText = getAction("acService_send_emailAttachment");
            cmdSendAttach.CommandType = CommandType.Text;

            //recipient
            SqlCommand cmdSendrecipient = new SqlCommand();
            cmdSendrecipient.Connection = conn;
            cmdSendrecipient.CommandText = getAction("acService_send_emailrecipient");
            cmdSendrecipient.CommandType = CommandType.Text;

            var readerSend = cmdSend.ExecuteReader();
            readerSend.Read();
            DataTable dt = new DataTable();
            dt.Load(readerSend);
            readerSend.Close();

            SqlCommand cmdUpdate = new SqlCommand();
            cmdUpdate.Connection = conn;
            cmdUpdate.CommandText = "UPDATE V_EMAIL_Email SET delivery=1 where ID=@ID";
            cmdUpdate.CommandType = CommandType.Text;

            MailMessage mail = new MailMessage();
            SmtpClient smtpserver = new SmtpClient("smtp.gmail.com");
            smtpserver.Port = 587;
            smtpserver.Credentials = new NetworkCredential("sama.drs2019@gmail.com", "9338047927");
            smtpserver.EnableSsl = true;
            foreach (DataRow ro in dt.Rows)
            {
                string id = ro["ID"].ToString();
                cmdUpdate.Parameters.Clear();
                cmdUpdate.Parameters.Add(new SqlParameter("@ID", id));
                try
                {
                    mail.Subject = ro["subject"].ToString();
                    mail.From = new MailAddress("sama.drs2019@gmail.com");
                    mail.To.Add(ro["profile_name"].ToString());
                    mail.Body = ro["body"].ToString();

                    cmdSendAttach.Parameters.Clear();
                    cmdSendAttach.Parameters.Add(new SqlParameter("@EmailID", id));
                    var readerAttach = cmdSendAttach.ExecuteReader();
                    while (readerAttach.Read())
                    {
                        Byte[] byteBLOBData = new Byte[0];
                        byteBLOBData = (Byte[])(readerAttach["attachment"]);
                        MemoryStream stmBLOBData = new MemoryStream(byteBLOBData);
                        mail.Attachments.Add(new Attachment(stmBLOBData, readerAttach["attachmentFileName"].ToString()));
                    }
                    cmdSendrecipient.Parameters.Clear();
                    cmdSendrecipient.Parameters.Add(new SqlParameter("@EmailID", id));
                    var readerrecipient = cmdSendrecipient.ExecuteReader();

                    while (readerrecipient.Read())
                    {
                        mail.Bcc.Add(readerrecipient["recipientEmail"].ToString());

                    }
                    smtpserver.Send(mail);
                    cmdUpdate.ExecuteNonQuery();
                    readerAttach.Close();
                    readerrecipient.Close();

                }
                catch (Exception ex)
                {
                    AddEvent(ex.Message.ToString(), EventLogEntryType.Error);
                }
            }
            AddEvent("تعداد" + " " + dt.Rows.Count + " " + "ایمیل با موفقیت ارسال شد  ".ToString());
        }

        private void receivemail()
        {

            SqlCommand cmdR = new SqlCommand();
            cmdR.Connection = conn;
            cmdR.CommandText = getAction("acService_receive_email");
            cmdR.CommandType = CommandType.Text;

            //Second Action
            SqlCommand cmdRA = new SqlCommand();
            cmdRA.Connection = conn;
            cmdRA.CommandText = getAction("acService_Receive_EmailAttachment");
            cmdRA.CommandType = CommandType.Text;

            //Third Action
            SqlCommand cmdRR = new SqlCommand();
            cmdRR.Connection = conn;
            cmdRR.CommandText = getAction("acService_Receive_EmailRecipient");
            cmdRR.CommandType = CommandType.Text;

            int count = 0;

            using (ImapClient client = new ImapClient())
            {

                client.Connect("imap.gmail.com", 993, SecureSocketOptions.SslOnConnect);
                client.Authenticate("sama.drs2019@gmail.com", "9338047927");
                client.Inbox.Open(FolderAccess.ReadWrite);
                client.Inbox.Check();

                var uids = client.Inbox.Search(SearchQuery.NotSeen);

                var items = client.Inbox.Fetch(uids, MessageSummaryItems.UniqueId | MessageSummaryItems.BodyStructure);

                foreach (var item in items)
                {

                    count++;
                    var message = client.Inbox.GetMessage(item.Index);
                    cmdR.Parameters.Clear();
                    cmdR.Parameters.Add(new SqlParameter("@profile_name", message.From.ToString()));
                    cmdR.Parameters.Add(new SqlParameter("@RefID", ""));
                    cmdR.Parameters.Add(new SqlParameter("@subject", message.Subject));
                    var bodyPart = item.TextBody;
                    var body = (TextPart)client.Inbox.GetBodyPart(item.UniqueId, bodyPart);
                    if (body == null)
                        cmdR.Parameters.Add(new SqlParameter("@body", ""));
                    else
                        cmdR.Parameters.Add(new SqlParameter("@body", body.Text));
                    cmdR.Parameters.Add(new SqlParameter("@UID", message.MessageId));

                    cmdR.Parameters.Add(new SqlParameter("@EmailType", "2"));
                    cmdR.Parameters.Add(new SqlParameter("@State", ""));
                    cmdR.Parameters.Add(new SqlParameter("@SendDate", message.Date));
                    var readerEmail = cmdR.ExecuteReader();
                    readerEmail.Read();
                    string id = readerEmail["ID"].ToString();
                    readerEmail.Close();
                    var attachmentIndex = 0;
                    foreach (var attachment in item.Attachments)
                    {
                        attachmentIndex++;
                        cmdRA.Parameters.Clear();
                        cmdRA.Parameters.Add(new SqlParameter("@EmailID", id));
                        cmdRA.Parameters.Add(new SqlParameter("@Rdf", attachmentIndex));
                        cmdRA.Parameters.Add(new SqlParameter("@Sharh", ""));
                        cmdRA.Parameters.Add(new SqlParameter("@attachmentFileName", attachment.FileName));

                        var entity = client.Inbox.GetBodyPart(item.UniqueId, attachment);
                        MemoryStream ms = new MemoryStream();
                        if (entity is MessagePart)
                        {
                            var rfc822 = (MessagePart)entity;
                            rfc822.Message.WriteTo(ms);
                        }
                        else
                        {
                            var part = (MimePart)entity;
                            part.Content.DecodeTo(ms);
                        }
                        ms.Position = 0;
                        cmdRA.Parameters.Add(new SqlParameter("@attachment", ms));
                        cmdRA.ExecuteNonQuery();
                    }

                    if (message.Cc.Count > 0)
                    {
                        foreach (InternetAddress addr in message.Cc)
                        {
                            string addrString = addr.Name;
                            cmdRR.Parameters.Clear();
                            cmdRR.Parameters.Add(new SqlParameter("@EmailID", id));
                            cmdRR.Parameters.Add(new SqlParameter("@recipientType", "Cc"));
                            cmdRR.Parameters.Add(new SqlParameter("@recipientEmail", addrString));
                            cmdRR.ExecuteNonQuery();
                        }
                    }

                    //set mail as seen
                    client.Inbox.SetFlags(item.UniqueId, MessageFlags.Seen, true);
                }
                cmdR.Dispose();
                cmdRA.Dispose();
                cmdRR.Dispose();

                AddEvent("تعداد" + " " + count + " " + "ایمیل با موفقیت درج شد ".ToString());
            }
        }

        private void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            {
                _timer.Stop();

                AddEvent("Timer is starting ...");
                try
                {
                    string connectionString = ConfigurationManager.ConnectionStrings["SamaMailService"].ToString();
                    conn = new SqlConnection(connectionString);
                    conn.Open();
                }

                catch (Exception ex)
                {
                    AddEvent(ex.Message.ToString(), EventLogEntryType.Error);
                }

                sendmail();
                receivemail();

                _lastRun = DateTime.Now;
                _timer.Start();
                AddEvent("Timer is finishing ...");
            }
        }
        protected override void OnStart(string[] args)
        {
            AddEvent("Service is starting ...");
            _timer = new Timer(.1 * 60 * 1000); // every 10 minutes
            _timer.Elapsed += new System.Timers.ElapsedEventHandler(timer_Elapsed);
            _timer.Start();
        }

        protected override void OnStop()
        {
        }
    }
}
