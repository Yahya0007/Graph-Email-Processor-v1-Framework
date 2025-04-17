using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Constants = Microsoft.VisualBasic.Constants;


namespace Email_Processor
{
    public partial class frmEmailProcessor : Form
    {

        enum UnsubscribeStatus
        {
            Unsubscribed,
            NotFound,
            Error
        }

        enum LogChannels
        {
            Access,
            SQLServer,
            Misc
        }

        private static GraphServiceClient _graphServiceClient;
        private GraphServiceClient graphClient = null;
        private const string domain = "high-society.co.uk";

        public const string TEST_TOKEN = "[TEST]";
        public const string SYSTEM_LOG_PATH = @"C:\Graph Email Processor v1 Logs\";
        public const string SYSTEM_LOG_ACCESS = "GraphEmailProcessor-Access.log";
        public const string SYSTEM_LOG_SQLSERVER = "GraphEmailProcessor-SQLServer.log";
        public const string SYSTEM_LOG_ACCESS_NOTFOUND = "GraphEmailProcessor-Access-NotFound.log";
        public const string SYSTEM_LOG_SQLSERVER_NOTFOUND = "GraphEmailProcessor-SQLServer-NotFound.log";
        public const string REFERENCES_LOG = "ProcessAttachmentsv1-References.log";
        public const string UNSUBSCRIBE_LOG = "ProcessAttachmentsv1-Unsubscribe.log";
        public const string TIMESHEETS_DATA_PATH = @"F:\PDF Timesheets\Data\";
        public const string TIMESHEETS_PDF_PATH = @"F:\PDF Timesheets\";
        public const string REFERENCES_DATA_PATH = @"F:\References\Data\";
        public const string REFERENCES_PDF_PATH = @"F:\References\Sent\";
        public const string MSG_NOT_SAVED = "WARNING! Data file not saved, no matching pdf file: ";
        public const string MSG_SAVED = "SAVED: ";
        public const string MSG_NOT_AUTOBOOKED = "NOT AUTO BOOKED: ";
        public const string MSG_AUTOBOOKED = "AUTO BOOKED: ";
        //public const string RemoteSQLServerConnSt = "Data Source=91.232.125.193;Initial Catalog=HS_Staff_Portal;Persist Security Info=True;User ID=HS_Staff_Portal;Password=yMYSyZKY#9";
        public const string RemoteSQLServerConnSt = "Data Source=91.232.125.193;Initial Catalog=EMS_2018_HS;Persist Security Info=True;User ID=EMS_2018_HS;Password=P@ssword01;TrustServerCertificate=True;";
        public const string LocalSQLServerConnSt = @"Data Source=DB-SERVER;Initial Catalog=EMS_2018_HS;Persist Security Info=True;User ID=EMS_2018_HS;Password=P@ssword01;TrustServerCertificate=True;";
        public const string LocalAccessConnSt = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""S:\Events Data\Events Data.mdb"";";
        public bool InProcess = false;

        public enum PDFType
        {
            None = -1,
            Timesheet = 0,
            Reference = 1
        }

        public DateTime CurrentTimeStamp;

        private SecureString ConvertToSecureString(string password)
        {
            if (password == null)
            {
                throw new ArgumentNullException("password");
            }

            var securePassword = new SecureString();
            foreach (char c in password)
                securePassword.AppendChar(c);
            securePassword.MakeReadOnly();
            return securePassword;
        }

        public frmEmailProcessor()
        {
            InitializeComponent();

            // Load appsettings.json
            var config = LoadAppSettings();
            if (null == config)
            {
                Console.WriteLine("Missing or invalid appsettings.json file. Please see README.md for configuration instructions.");
                return;
            }

            //Query using Graph SDK (preferred when possible)
            //graphClient = GetAuthenticatedGraphClient(config);
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            Process();
        }

        #region "ProcessUnsubscribeEmailsAsync"

        private async Task ProcessUnsubscribeEmailsAsync(GraphServiceClient graphClient)
        {
            List<QueryOption> options = new List<QueryOption>
            {
                //new QueryOption("$top", "1")
            };

            //var graphResult = await graphClient.Users.Request(options)
            //    .Filter("startswith(displayName,'Unsubscribe')")
            //    .GetAsync();

            //Console.WriteLine("Email Processor Result");
            //Console.WriteLine(graphResult[0].DisplayName);

            var Users = await graphClient
                         .Users
                         .Request()
                         .Filter("startswith(displayName,'Unsubscribe')").Top(1)
                         .GetAsync();

            var user = Users[0];

            // Get message from the user's inbox
            //var inboxMessages = await graphClient
            //            .Users[User.Id]
            //            .MailFolders.Inbox
            //            .Messages
            //            .Request()
            //            .Select(e => new
            //            {
            //                e.Subject,
            //                e.Body,
            //                e.Sender
            //            })
            //            .OrderBy("ReceivedDateTime/dateTime")
            //            .GetAsync();

            //var inboxMessages = await graphClient
            // .Users[user.Id]
            // .MailFolders.Inbox
            // .Messages
            // .Request()
            // .Filter("startswith(Subject,'Unsub: ')")
            // .OrderBy("receivedDateTime DESC")
            // .GetAsync();

            //var inboxMessages = await graphClient
            //     .Users[user.Id]
            //     .MailFolders.Inbox
            //     .Messages
            //     .Request()
            //     .Filter("not startswith(Subject,'Unsub: ')")
            //     .GetAsync();


            //var inboxMessages1 = await graphClient
            //    .Me
            //    .MailFolders.Inbox
            //    .Messages
            //    .Request()
            //    .OrderBy("")
            //    .GetAsync();

            var inboxMessages = await graphClient
                  .Users[user.Id]
                  .MailFolders.Inbox
                  .Messages
                  .Request()
                  .Filter("isRead eq false")
                  .Top(1000)
                  //.Select("Body, Subject")
                  .GetAsync();

            foreach (Microsoft.Graph.Message x in inboxMessages)
            {
                await ProcessUnsubscribeMessageAsync(x, user);
            }
        }

        private async Task ProcessUnsubscribeMessageAsync(Microsoft.Graph.Message message, User user)
        {
            //Check if message has further message attachments and process them
            //if (message.HasAttachments == true)
            //{
            //    var attachments = await graphClient
            //                        .Users[user.Id]
            //                        .Messages[$"{message.Id}"]
            //                        .Attachments
            //                        .Request()
            //                        .GetAsync();

            //    //foreach (var attachment in message.Attachments)
            //    foreach (var attachment in attachments)
            //    {
            //        if (attachment.ODataType == "#microsoft.graph.itemAttachment")
            //        {

            //            var attachmentRequest = await graphClient
            //                                        .Users[user.Id]
            //                                        .MailFolders
            //                                        .Inbox
            //                                        .Messages[message.Id]
            //                                        .Attachments[attachment.Id]
            //                                        .Request()
            //                                        .Expand("microsoft.graph.itemattachment/item")
            //                                        .GetAsync();

            //            var itemAttachment = (ItemAttachment)attachmentRequest;
            //            var itemMessage = (Microsoft.Graph.Message)itemAttachment.Item;

            //            if (itemMessage != null)
            //            {
            //                await ProcessUnsubscribeMessageAsync(itemMessage, user);
            //            }


            //            //if (attachmentRequest.ODataType == "#microsoft.graph.message")
            //            //{

            //            //    attachmentRequest.GetType
            //            //    var itemMessage = (Microsoft.Graph.Message)attachment;

            //            //    if (itemMessage != null)
            //            //    {
            //            //        await ProcessMessageAsync(itemMessage, user);
            //            //    }
            //            //}

            //            //                    var attachment1 = await graphClient.Me.Messages["{message-id}"].Attachments["{attachment-id}"]
            //            //.Request()
            //            //.GetAsync();
            //            //var itemAttachment = (ItemAttachment)attachmentRequest.Result;


            //        }
            //        //else
            //        //{
            //        //    var fileAttachment = (FileAttachment)attachment;
            //        //    System.IO.File.WriteAllBytes(System.IO.Path.Combine(downloadPath, fileAttachment.Name), fileAttachment.ContentBytes);
            //        //}
            //    }
            //}

            //List<string> Emails = new List<string> { };



            if (!message.Sender.EmailAddress.Address.Contains(domain))
            {
                //Emails.Add(x.Sender.EmailAddress.Address);
                UnsubcribeStaff(message, message, user);
            }
            else
            {
                //x.ConversationId
                //x.ConversationIndex

                //var inboxMessages1 = graphClient
                //    .Users[User.Id]
                //    .Messages
                //    .Request()
                //    .Select(e => new
                //    {
                //        e.Subject,
                //        e.Body,
                //        e.Sender
                //    })
                //    .Filter($"startswith(ConversationId,{x.ConversationId})")
                //    .OrderBy("Recieved/dateTime")
                //    .GetAsync().Result;

                // Get email to subscribe from body text
                var z = message.Body.ContentType;
                message.Body.ContentType = BodyType.Text;
                string EmailBody = message.Body.Content;
                message.Body.ContentType = z;

                Regex emailRegex = new Regex(@"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*", RegexOptions.IgnoreCase);
                //find items that matches with our pattern
                MatchCollection emailMatches = emailRegex.Matches(EmailBody);

                foreach (Match emailMatch in emailMatches)
                {
                    if (!emailMatch.Value.Contains(domain))
                    {
                        //Emails.Add(emailMatch.Value);
                        var m = new Microsoft.Graph.Message();

                        var b = new ItemBody();
                        b.Content = EmailBody;
                        m.Body = b;

                        var e = new EmailAddress();
                        e.Address = emailMatch.Value;
                        e.Name = emailMatch.Value;

                        var s = new Recipient();
                        s.EmailAddress = e;

                        m.Sender = s;

                        //if (UnsubcribeStaff(emailMatch.Value, emailMatch.Value, EmailBody, null)== UnsubscribeStatus.Unsubscribed)
                        UnsubcribeStaff(m, message, user);
                    }
                }
            }

            //Finally update the message subject back into data store
            //if (x.Subject.Contains("Unsub: "))
            //{

            //    //List<QueryOption> options1 = new List<QueryOption>
            //    //{
            //    //    //new QueryOption("$top", "1")
            //    //};


            //    //await graphClient
            //    //    .Users[user.Id]
            //    //    .MailFolders.Inbox
            //    //    .Messages["{x.Id}"]
            //    //    .Request()
            //    //    .UpdateAsync(x);

            //    Update(user, x);


            //    //var inboxMessages = await graphClient
            //    // .Users[user.Id]
            //    // .MailFolders.Inbox
            //    // .Messages
            //    // .Request()
            //    // .OrderBy("receivedDateTime DESC")
            //    // .GetAsync();

            //}

        }

        //private void Update(User user, Microsoft.Graph.Message x)
        //{
        //    graphClient
        //    .Users[user.Id]
        //    .MailFolders.Inbox
        //    .Messages[$"{x.Id}"]
        //    .Request()
        //    .UpdateAsync(x);

        //    //Console.WriteLine("I am here!");

        //    LogMessage("Message from sender " + x.Sender.EmailAddress.Address + " with subject '" + x.Subject + "' updated successfully.", x, TraceEventType.Information);
        //}

        private void UnsubcribeStaff(Microsoft.Graph.Message message, Microsoft.Graph.Message originalmessage, User user)
        {
            LogMessage("Processing unsubscribe request.", message, TraceEventType.Information);
            string Email = message.Sender.EmailAddress.Address;
            string StaffName = message.Sender.EmailAddress.Name;
            string BodyText = message.Body.Content.Replace("\"", "\"\"");
            //return UnsubscribeStaffAccess(Email, StaffName, BodyText, message) | UnsubscribeStaffSQLServer(Email, StaffName, BodyText, message);
            UnsubscribeStaff(Email, StaffName, BodyText, message, originalmessage, user);
            LogMessage(" ", message, TraceEventType.Information, "", true);
        }

        private void UnsubscribeStaff(string Email, string StaffName, string BodyText, Microsoft.Graph.Message message, Microsoft.Graph.Message originalmessage, User user)
        {
            //return

            var xx = UnsubscribeStaffAccess(Email, StaffName, BodyText, message);
            var yy = UnsubscribeStaffSQLServer(Email, StaffName, BodyText, message);

            if ((xx == UnsubscribeStatus.Error) | (yy == UnsubscribeStatus.Error))
            {
                originalmessage.Subject = "Unsub: Error: " + originalmessage.Subject.Replace("Unsub: Error: ", "").Replace("Unsub: Unsubscribed: ", "").Replace("Unsub: Not Found: ", "");
                //await graphClient.Users[user.Id].Messages["{originalmessage.Id}"]
                //    .Request()
                //    .UpdateAsync(originalmessage);

            }
            else
            {
                if (!originalmessage.Subject.Contains("Unsub: Error: "))
                {
                    if ((xx == UnsubscribeStatus.Unsubscribed) | (yy == UnsubscribeStatus.Unsubscribed))
                    {
                        originalmessage.Subject = "Unsub: Unsubscribed: " + originalmessage.Subject.Replace("Unsub: Unsubscribed: ", "").Replace("Unsub: Not Found: ", "");
                        //await graphClient.Users[user.Id].Messages["{originalmessage.Id}"]
                        //    .Request()
                        //    .UpdateAsync(originalmessage);
                    }
                    else
                    {
                        if (!originalmessage.Subject.Contains("Unsub: Unsubscribed: "))
                        {
                            originalmessage.Subject = "Unsub: Not Found: " + originalmessage.Subject.Replace("Unsub: Not Found: ", "");
                            //await graphClient.Users[user.Id].Messages["{originalmessage.Id}"]
                            //    .Request()
                            //    .UpdateAsync(originalmessage);

                            AddNotFoundEmail(Email);
                        }
                    }
                }
            }

            if (originalmessage.Subject.Contains("Unsub: "))
            {

                //graphClient
                // .Users[user.Id]
                // .MailFolders.Inbox
                // .Messages[$"{originalmessage.Id}"]
                // .Request()
                // .UpdateAsync(originalmessage);

                //originalmessage.IsDraft = true;

                //graphClient
                //    .Users[user.Id]
                //    .MailFolders.Inbox
                //    .Messages[$"{originalmessage.Id}"]
                //    .Request()
                //    .UpdateAsync(originalmessage);

                //originalmessage.IsDraft = false;

                //originalmessage.IsDraft = true;

                var categories = new List<string>() { };

                if (originalmessage.Subject.Contains("Unsub: Error:"))
                {
                    categories.Add("Error not unsubscribed");
                }
                else if (originalmessage.Subject.Contains("Unsub: Unsubscribed:"))
                {
                    categories.Add("Unsubscribed");
                }
                else if (originalmessage.Subject.Contains("Unsub: Not Found:"))
                {
                    categories.Add("Not Found");
                }

                //var outlookCategory = new OutlookCategory
                //{
                //    Color = CategoryColor.Preset15
                //};

                try
                {
                    //graphClient
                    //    .Users[user.Id]
                    //    .MailFolders.Inbox
                    //    .Messages[$"{originalmessage.Id}"]
                    //    .Request()
                    //    .UpdateAsync(new Microsoft.Graph.Message()
                    //    {
                    //        Categories = categories
                    //    });

                    var subject = originalmessage.Subject;

                    originalmessage.IsDraft = true;

                    graphClient
                        .Users[user.Id]
                        .MailFolders.Inbox
                        .Messages[$"{originalmessage.Id}"]
                        .Request()
                        .UpdateAsync(new Microsoft.Graph.Message()
                        {
                            Subject = subject,
                            IsRead = true
                        });

                    originalmessage.IsDraft = false;
                }
                catch (ServiceException ex)
                {
                    LogMessage("UnsubscribeStaff::UpdateAsync: Email message categories for staff with email " + Email + " not updated." + ex.Message, message, TraceEventType.Error);
                }

                //graphClient
                //    .Users[user.Id]
                //    .Outlook
                //    .MasterCategories($"{outlookCategory-id}")
                //    .Request()
                //    .UpdateAsync(originalmessage);

                //originalmessage.IsDraft = false;

                //Console.WriteLine("I am here!");

                //LogMessage("Message from sender " + originalmessage.Sender.EmailAddress.Address + " with subject '" + originalmessage.Subject + "' updated successfully.", originalmessage, TraceEventType.Information);
            }
        }

        private UnsubscribeStatus UnsubscribeStaffAccess(string Email, string StaffName, string BodyText, Microsoft.Graph.Message message)
        {
            object StaffID;
            StaffID = DLookupAccess("Staff ID", "Staff", "[E-Mail] = \"" + Email + "\"");
            if (StaffID == null)
            {
                StaffID = -1;
                LogMessage("Staff with email " + Email + " ***not found in Access database.", message, TraceEventType.Error);
                return UnsubscribeStatus.NotFound;
            }

            string St3 = "UPDATE Staff SET Staff.NoMailing = True, Staff.NoMailingDate = Date() WHERE [E-Mail] = \"" + Email + "\"";
            if (ExecuteNonQuery(St3, message) == null)
            {
                LogMessage("Staff " + StaffID.ToString() + ", Email " + Email + ", staff ***failed to unsubscribe.", message, TraceEventType.Error);
                return UnsubscribeStatus.Error;
            }

            var NoMailing = DLookupAccess("NoMailing", "Staff", "[E-Mail] = \"" + Email + "\"");
            if (Conversions.ToBoolean(NoMailing))
            {
                LogMessage("Staff " + StaffID.ToString() + ", Email " + Email + ", staff successfully unsubscribed.", message, TraceEventType.Information);
                St3 = "INSERT INTO StaffUnsubscribeRequests ( StaffID, Email, StaffName, Unsubsribed, Body ) " + "SELECT " + StaffID.ToString() + ", \"" + Email + "\", \"" + StaffName + "\", Now(), \"" + BodyText + "\"";
                if (Conversions.ToInteger(ExecuteNonQuery(St3, message)) >= 1)
                {
                    LogMessage("Staff " + StaffID.ToString() + ", Email " + Email + ", staff unsubscribe history updated.", message, TraceEventType.Error);
                    return UnsubscribeStatus.Unsubscribed;
                }
                else
                {
                    LogMessage("Staff " + StaffID.ToString() + ", Email " + Email + ", staff unsubscribe history update ***failed.", message, TraceEventType.Error);
                    return UnsubscribeStatus.Error;
                }
            }

            return UnsubscribeStatus.Error;
        }

        private UnsubscribeStatus UnsubscribeStaffSQLServer(string Email, string StaffName, string BodyText, Microsoft.Graph.Message message)
        {
            object StaffID;
            StaffID = DLookupSQLServer("ID", "Staff", "Email = '" + Email + "'");
            if (StaffID == null)
            {
                StaffID = -1;
                LogMessage("Staff with email " + Email + " not found in SQL Server database.", message, TraceEventType.Error, "", false, LogChannels.SQLServer);
                return UnsubscribeStatus.NotFound;
            }

            string St3 = "UPDATE Staff SET Staff.NoMailing = 1, Staff.NoMailingDate = { fn now() } WHERE [EMail] = '" + Email + "'";
            if (ExecuteNonQuerySQLServer(St3, message) == null)
            {
                LogMessage("Staff " + StaffID.ToString() + ", Email " + Email + ", staff ***failed to unsubscribe.", message, TraceEventType.Error, "", false, LogChannels.SQLServer);
                return UnsubscribeStatus.Error;
            }

            var NoMailing = DLookupSQLServer("NoMailing", "Staff", "Email = '" + Email + "'");
            if (Conversions.ToBoolean(NoMailing))
            {
                LogMessage("Staff " + StaffID.ToString() + ", Email " + Email + ", staff successfully unsubscribed.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                St3 = "INSERT INTO StaffUnsubscribeRequests ( StaffID, Email, StaffName, Unsubsribed, Body ) " + "SELECT " + StaffID.ToString() + ", '" + Email + "', '" + StaffName + "', Now(), '" + BodyText + "'";
                if (Conversions.ToInteger(ExecuteNonQuery(St3, message)) >= 1)
                {
                    LogMessage("Staff " + StaffID.ToString() + ", Email " + Email + ", staff unsubscribe history updated.", message, TraceEventType.Error, "", false, LogChannels.SQLServer);
                    return UnsubscribeStatus.Unsubscribed;
                }
                else
                {
                    LogMessage("Staff " + StaffID.ToString() + ", Email " + Email + ", staff unsubscribe history update ***failed.", message, TraceEventType.Error, "", false, LogChannels.SQLServer);
                    return UnsubscribeStatus.Error;
                }
            }

            return UnsubscribeStatus.Error;
        }

        #endregion "ProcessUnsubscribeEmailsAsync"

        void UpdateReferencesSQLServer(Microsoft.Graph.Message message, string fePDFFileName, string feFileName)
        {
            try
            {
                var Conn = new SqlConnection(LocalSQLServerConnSt);
                //LogMessage("Opening reference SQL Server db connection.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                Conn.Open();
                //LogMessage("Reference SQL Server db connection opened.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                string StaffID;
                string St = "";
                if (fePDFFileName.IndexOf("-") != 0)
                {
                    SqlCommand command;
                    int i = 0;
                    int j = 0;
                    StaffID = LeftStr(fePDFFileName, fePDFFileName.IndexOf("-") - 1);
                    St = "UPDATE Staff SET Staff.LastEditableReference1Data = '" + feFileName + "', ReferenceReceived1 = { fn Now } " + "WHERE (((Staff.ID) = " + StaffID + ")) AND (Staff.LastEditableReference = " + fePDFFileName + ")";
                    try
                    {
                        command = new SqlCommand(St, Conn);
                        j = command.ExecuteNonQuery();
                        command.Dispose();
                    }
                    catch (Exception ex)
                    {
                        LogError("Error updating reference SQL Server db: " + ex.Message, message, TraceEventType.Error, "", false, LogChannels.SQLServer);
                    }

                    St = "UPDATE Staff SET Staff.LastEditableReference2Data = '" + feFileName + "', ReferenceReceived2 = { fn now } " + "WHERE (((Staff.ID) = " + StaffID + ")) AND (Staff.LastEditableReference2 = " + fePDFFileName + ")";
                    try
                    {
                        command = new SqlCommand(St, Conn);
                        j = command.ExecuteNonQuery();
                        command.Dispose();
                    }
                    catch (Exception ex)
                    {
                        LogError("Error updating reference SQL Server db: " + ex.Message, message, TraceEventType.Error, "", false, LogChannels.SQLServer);
                    }

                    Conn.Close();
                    if (i > 0)
                    {
                        LogMessage("Updated reference 1 SQL Server db.", message, TraceEventType.Information);
                        message.Subject = MSG_SAVED + message.Subject;
                        message.IsRead = true;
                        //message.Update(ConflictResolutionMode.AlwaysOverwrite);

                        graphClient.Me.Messages["{message.Id}"]
                            .Request()
                            .UpdateAsync(message);
                    }
                    else if (j > 0)
                    {
                        LogMessage("Updated reference 2 SQL Server db.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                        message.Subject = MSG_SAVED + message.Subject;
                        message.IsRead = true;
                        //message.Update(ConflictResolutionMode.AlwaysOverwrite);

                        graphClient.Me.Messages["{message.Id}"]
                           .Request()
                           .UpdateAsync(message);
                    }
                    else
                    {
                        LogMessage("No match for file " + " found in references: ", message, TraceEventType.Error, "", false, LogChannels.SQLServer);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("Error updating reference SQL Server db: " + ex.Message, message, TraceEventType.Error, "", false, LogChannels.SQLServer);
            }
        }

        void UpdateEditableTiemsheetSQLServer(Microsoft.Graph.Message message, string fePDFFileName, string feFileName)
        {
            try
            {
                string ConnSt = @"Data Source=DB-SERVER\SQL2K14;Initial Catalog=EMS_2018_HS;Persist Security Info=True;User ID=EMS_2018_HS;Password=P@ssword01";
                var Conn = new SqlConnection(ConnSt);
                //LogMessage("Opening SQL Server db connection.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                Conn.Open();
                //LogMessage("SQL Server db connection opened.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                string EventID;
                string St = "";
                if (fePDFFileName.IndexOf("-") != 0)
                {
                    EventID = LeftStr(fePDFFileName, fePDFFileName.IndexOf("-") - 1);
                    St = "UPDATE Events SET Events.LastEditableTimesheetData = '" + feFileName + "' " + "WHERE (((Events.ID)= " + EventID + "))";
                    var command = new SqlCommand(St, Conn);
                    command.ExecuteNonQuery();
                    Conn.Close();
                    LogMessage("Updated SQL Server db.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                    message.Subject = MSG_SAVED + message.Subject;
                    message.IsRead = true;
                    //message.Update(ConflictResolutionMode.AlwaysOverwrite);

                    graphClient.Me.Messages["{message.Id}"]
                        .Request()
                        .UpdateAsync(message);

                    string FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") + " Timesheet data for event " + EventID.ToString() + " saved.txt");
                    try
                    {
                        TextWriter tw = System.IO.File.CreateText(FullPath);
                        tw.Close();
                    }
                    catch (Exception ex)
                    {
                        LogError("Error writitng to log file: " + FullPath + "Exception: " + ex.Message, message, TraceEventType.Error, "", false, LogChannels.SQLServer);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("Error updating SQL Server db: " + ex.Message, message, TraceEventType.Error, "", false, LogChannels.SQLServer);
            }
        }

        void UpdateReferencesAccess(Microsoft.Graph.Message message, string fePDFFileName, string feFileName)
        {
            try
            {
                var Conn = new OleDbConnection(LocalAccessConnSt);
                //LogMessage("Opening reference db connection.", message, TraceEventType.Information);
                Conn.Open();
                //LogMessage("Reference db connection opened.", message, TraceEventType.Information);
                string StaffID;
                string St = "";
                if (fePDFFileName.IndexOf("-") != 0)
                {
                    OleDbCommand command;
                    int i = 0;
                    int j = 0;
                    StaffID = LeftStr(fePDFFileName, fePDFFileName.IndexOf("-") - 1);
                    St = "UPDATE Staff SET Staff.[Last Editable Reference 1 Data] = \"" + feFileName + "\", [Reference Received 1] = NOW() " + "WHERE (((Staff.[Staff ID]) = " + StaffID + ")) AND (Staff.[Last_Editable_Reference_1] = " + fePDFFileName + ")";
                    try
                    {
                        command = new OleDbCommand(St, Conn);
                        j = command.ExecuteNonQuery();
                        command.Dispose();
                    }
                    catch (Exception ex)
                    {
                        LogError("Error updating reference db: " + ex.Message, message, TraceEventType.Error);
                    }

                    St = "UPDATE Staff SET Staff.[Last Editable Reference 2 Data] = \"" + feFileName + "\", [Reference Received 2] = NOW() " + "WHERE (((Staff.[Staff ID]) = " + StaffID + ")) AND (Staff.[Last_Editable_Reference_2] = " + fePDFFileName + ")";
                    try
                    {
                        command = new OleDbCommand(St, Conn);
                        j = command.ExecuteNonQuery();
                        command.Dispose();
                    }
                    catch (Exception ex)
                    {
                        LogError("Error updating reference db: " + ex.Message, message, TraceEventType.Error);
                    }

                    Conn.Close();
                    if (i > 0)
                    {
                        LogMessage("Updated reference 1 db.", message, TraceEventType.Information);
                        message.Subject = MSG_SAVED + message.Subject;
                        message.IsRead = true;
                        //message.Update(ConflictResolutionMode.AlwaysOverwrite);

                        graphClient.Me.Messages["{message.Id}"]
                            .Request()
                            .UpdateAsync(message);

                    }
                    else if (j > 0)
                    {
                        LogMessage("Updated reference 2 db.", message, TraceEventType.Information);
                        message.Subject = MSG_SAVED + message.Subject;
                        message.IsRead = true;
                        //message.Update(ConflictResolutionMode.AlwaysOverwrite);

                        graphClient.Me.Messages["{message.Id}"]
                           .Request()
                           .UpdateAsync(message);
                    }
                    else
                    {
                        LogMessage("No match for file " + " found in references: ", message, TraceEventType.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("Error updating reference db: " + ex.Message, message, TraceEventType.Error);
            }
        }

        void UpdateEditableTiemsheetAccess(Microsoft.Graph.Message message, string fePDFFileName, string feFileName)
        {
            try
            {
                string ConnSt = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""S:\Events Data\Events Data.mdb"";";
                var Conn = new OleDbConnection(ConnSt);
                //LogMessage("Opening db connection.", message, TraceEventType.Information);
                Conn.Open();
                //LogMessage("Db connection opened.", message, TraceEventType.Information);
                string EventID;
                string St = "";
                if (fePDFFileName.IndexOf("-") != 0)
                {
                    EventID = LeftStr(fePDFFileName, fePDFFileName.IndexOf("-") - 1);
                    St = "UPDATE Events SET Events.[Last Editable Timesheet Data] = \"" + feFileName + "\" " + "WHERE (((Events.[Event ID])= " + EventID + "))";
                    var command = new OleDbCommand(St, Conn);
                    command.ExecuteNonQuery();
                    Conn.Close();
                    LogMessage("Updated db.", message, TraceEventType.Information);
                    message.Subject = MSG_SAVED + message.Subject;
                    message.IsRead = true;
                    //message.Update(ConflictResolutionMode.AlwaysOverwrite);

                    graphClient.Me.Messages["{message.Id}"]
                       .Request()
                       .UpdateAsync(message);

                    string FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") + " Timesheet data for event " + EventID.ToString() + " saved.txt");
                    try
                    {
                        TextWriter tw = System.IO.File.CreateText(FullPath);
                        tw.Close();
                    }
                    catch (Exception ex)
                    {
                        LogError("Error writitng to log file: " + FullPath + "Exception: " + ex.Message, message, TraceEventType.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("Error updating db: " + ex.Message, message, TraceEventType.Error);
            }
        }

        string GetPDFFileName(string feFileName, Microsoft.Graph.Message item, ref PDFType PT)
        {
            string fePDFFileName = "";
            string fePDFFileName1 = "";
            //System.Windows.Forms.Application.DoEvents();

            var r = new Regex(@"\d{1,2}\-\d{1,2}\-\d{4}", RegexOptions.IgnoreCase);
            var m = r.Match(feFileName);
            if (m.Success)
            {
                int i;
                if (feFileName.IndexOf("]") > 0)
                {
                    i = m.Index + 11;
                    fePDFFileName = Strings.Left(feFileName, i) + ".pdf";
                    LogMessage("feFileName.IndexOf(\"]\") > 0: feFileName: " + feFileName + ", fePDFFileName: " + fePDFFileName, item, TraceEventType.Information);
                }
                else
                {
                    string DateSt = Mid(feFileName, m.Index, 10);
                    i = m.Index + 10;
                    fePDFFileName = Strings.Left(feFileName, i) + ".pdf";
                    fePDFFileName1 = fePDFFileName.Replace(DateSt, "[" + DateSt + "]");
                    LogMessage("feFileName.IndexOf(\"]\") <= 0: feFileName: " + feFileName + ", fePDFFileName: " + fePDFFileName, item, TraceEventType.Information);
                }

                LogMessage("fePDFFileNamet: " + fePDFFileName + ", fePDFFileName1: " + fePDFFileName1, item, TraceEventType.Information);
                if (System.IO.File.Exists(Path.Combine(TIMESHEETS_PDF_PATH, fePDFFileName)))
                {
                    LogMessage("GetPDFFileName: Exists: " + fePDFFileName, item, TraceEventType.Information);
                    PT = PDFType.Timesheet;
                    return fePDFFileName;
                }
                else if (System.IO.File.Exists(Path.Combine(TIMESHEETS_PDF_PATH, fePDFFileName1)))
                {
                    LogMessage("GetPDFFileName: Exists: " + fePDFFileName1, item, TraceEventType.Information);
                    PT = PDFType.Timesheet;
                    return fePDFFileName1;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                // No date found
                LogMessage("GetPDFFileName: No match, feFileName: " + feFileName, item, TraceEventType.Information);
                return null;
            }

            // If feFileName.IndexOf("]") > 0 Then
            // fePDFFileName = LeftStr(feFileName, feFileName.IndexOf("]")).ToLower().Replace("_data", "") & ".pdf"
            // LogMessage("feFileName.IndexOf(""]"") > 0: feFileName: " & feFileName & ", fePDFFileName: " & fePDFFileName, Message, TraceEventType.Information)
            // Else
            // fePDFFileName = feFileName.ToLower().Replace(".fdf", ".pdf").Replace("_data", "")
            // LogMessage("feFileName.IndexOf(""]"") <= 0: feFileName: " & feFileName & ", fePDFFileName: " & fePDFFileName, Message, TraceEventType.Information)
            // Dim DateSt As String = Microsoft.VisualBasic.Right(fePDFFileName, 14)
            // fePDFFileName1 = fePDFFileName.Replace(DateSt, "[" & DateSt)
            // fePDFFileName1 = fePDFFileName1.Replace(".pdf", "].pdf")
            // LogMessage("fePDFFileNamet: " & fePDFFileName & ", fePDFFileName1: " & fePDFFileName1 & ", DateSt: " & DateSt, Message, TraceEventType.Information)
            // End If

        }

        static string LeftStr(string param, int length)
        {
            string result = param.Substring(0, length);
            return result;
        }

        static string RightStr(string param, int length)
        {
            string result = param.Substring(param.Length - length, length);
            return result;
        }

        static string Mid(string param, int startIndex, int length)
        {
            string result = param.Substring(startIndex, length);
            return result;
        }

        static string Mid(string param, int startIndex)
        {
            string result = param.Substring(startIndex);
            return result;
        }

        void LogMessage(string eventName, Microsoft.Graph.Message item, TraceEventType e, string message = "", bool CRLF = false, LogChannels channel = LogChannels.Access, bool UI = false)
        {
            if (!System.IO.Directory.Exists(SYSTEM_LOG_PATH))
            {
                System.IO.Directory.CreateDirectory(SYSTEM_LOG_PATH);
            }

            if (UI)
            {
                System.Windows.Forms.Application.DoEvents();
                this.Msg.Text = eventName;
                //System.Windows.Threading.Dispatcher.Invoke(new Action(() => Msg.Text = eventName));
                System.Windows.Forms.Application.DoEvents();
            }

            string FullPath = "";

            if (channel == LogChannels.Access)
                FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") + " " + SYSTEM_LOG_ACCESS);
            else
                FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") + " " + SYSTEM_LOG_SQLSERVER);

            TextWriter tw = System.IO.File.AppendText(FullPath);
            try
            {
                if (string.IsNullOrEmpty(eventName))
                {
                    tw.WriteLine("---");
                    tw.WriteLine(" ");
                }
                else
                {
                    tw.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + Constants.vbTab + eventName + " - " + (item == null ? "" : item.Sender.EmailAddress.Address) + " - " + (item == null ? "" : item.Subject) + " : " + message);
                }

                if (CRLF)
                {
                    tw.WriteLine(" ");
                    tw.WriteLine(" ");
                }
            }
            catch (Exception ex)
            {

                if (UI)
                {
                    Interaction.MsgBox(ex.Message);
                }
                tw.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + Constants.vbTab + eventName + " - " + (item == null ? "" : item.Sender.EmailAddress.Address) + " - " + (item == null ? "" : item.Subject) + " : " + ex.Message);
            }
            finally
            {
                tw.Close();
            }
        }

        void LogError(string eventName, Microsoft.Graph.Message item, TraceEventType e, string Message = "", bool CRLF = false, LogChannels channel = LogChannels.Access, bool UI = false)
        {
            if (!System.IO.Directory.Exists(SYSTEM_LOG_PATH))
            {
                System.IO.Directory.CreateDirectory(SYSTEM_LOG_PATH);
            }

            if (UI)
            {
                System.Windows.Forms.Application.DoEvents();
                this.Msg.Text = eventName;
                System.Windows.Forms.Application.DoEvents();
            }

            string FullPath = "";

            if (channel == LogChannels.Access)
                FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") + " ERROR - " + SYSTEM_LOG_ACCESS);
            else
                FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") + " ERROR - " + SYSTEM_LOG_SQLSERVER);


            TextWriter tw = System.IO.File.AppendText(FullPath);
            try
            {

                if (CRLF)
                {
                    tw.WriteLine(" ");
                    tw.WriteLine(" ");
                }

                // Dim tw As TextWriter = System.IO.File.AppendText(FullPath)

                if (string.IsNullOrEmpty(eventName))
                {
                    tw.WriteLine("---");
                    tw.WriteLine(" ");
                }
                else
                {
                    tw.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + Constants.vbTab + eventName + " - " + (item == null ? "" : item.From.EmailAddress.Address) + " - " + (item == null ? "" : item.Subject) + " : " + Message);
                }
            }

            // tw.Close()
            catch (Exception ex)
            {
                if (UI)
                {
                    Interaction.MsgBox(ex.Message);
                }
                tw.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + Constants.vbTab + eventName + " - " + (item == null ? "" : item.Sender.EmailAddress.Address) + " - " + (item == null ? "" : item.Subject) + " : " + ex.Message);

            }
            finally
            {
                tw.Close();
            }
        }

        private void AddNotFoundEmail(string email, bool CRLF = false, LogChannels channel = LogChannels.Access)
        {
            if (!System.IO.Directory.Exists(SYSTEM_LOG_PATH))
            {
                System.IO.Directory.CreateDirectory(SYSTEM_LOG_PATH);
            }

            System.Windows.Forms.Application.DoEvents();
            this.Msg.Text = "Adding not found email: " + email;
            System.Windows.Forms.Application.DoEvents();

            string FullPath = "";

            if (channel == LogChannels.Access)
                FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd") + " " + SYSTEM_LOG_ACCESS_NOTFOUND);
            else
                FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd") + " " + SYSTEM_LOG_SQLSERVER_NOTFOUND);

            TextWriter tw = System.IO.File.AppendText(FullPath);
            try
            {
                if (CRLF)
                {
                    tw.WriteLine(" ");
                    tw.WriteLine(" ");
                }

                if (string.IsNullOrEmpty(email))
                {
                    tw.WriteLine("---");
                    tw.WriteLine(" ");
                }
                else
                {
                    tw.WriteLine(email);
                }
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.Message);
            }
            finally
            {
                tw.Close();
            }

        }

        private async Task ProcessStaffWorkRequests(GraphServiceClient graphClient)
        {

            var Users = await graphClient
                         .Users
                         .Request()
                         .Filter("startswith(mail,'WorkInterest')")
                         .GetAsync();

            var user = Users[0];

            //var inboxMessages = await graphClient
            //     .Users[user.Id]
            //     .MailFolders.Inbox
            //     .Messages
            //     .Request()
            //     .Filter("not startswith(Subject,'Unsub: ')")
            //     .GetAsync();

            //var inboxMessages = await graphClient
            //     .Users[user.Id]
            //     .MailFolders.Inbox
            //     .Messages
            //     .Request()
            //     .Filter("not isRead")
            //     .GetAsync();

            //var inboxMessages = await graphClient
            //     .Users[user.Id]
            //     .MailFolders.Inbox
            //     .Messages
            //     .Request()
            //     .Filter("not startswith(Subject, '" + MSG_SAVED + "')").Top(100)
            //     .GetAsync();

            var inboxMessages = await graphClient
                   .Users[user.Id]
                   .MailFolders.Inbox
                   .Messages
                   .Request()
                   .Filter("isRead eq false")
                   .Top(1000)
                   //.Select("Body, Subject")
                   .GetAsync();

            foreach (Microsoft.Graph.Message x in inboxMessages)
            {
                //if (x.IsRead == false)
                //if ((x.Subject != null) && (!x.Subject.Contains(MSG_NOT_SAVED)))
                //{
                ProcessStaffWorkRequestsMessageAsync(x, user);
                //}
            }
        }

        private void ProcessStaffWorkRequestsMessageAsync(Microsoft.Graph.Message EmailMsg, User user)
        {

            //return;

            LogMessage("Message " + EmailMsg.Subject + " loaded. Ready for message body processing.", EmailMsg, TraceEventType.Information);

            // Try
            //if (Strings.InStr(EmailMsg.Body.Content, "Event Ref") != 0)
            //{
            LogMessage("Processing message body.", EmailMsg, TraceEventType.Information);
            //}
            //else
            //{
            //    return;
            //}

            var Body1 = StripHtmlTags(EmailMsg.Body.Content.Replace("<BR>", Environment.NewLine)).Split(new string[] { "\n\n" }, StringSplitOptions.RemoveEmptyEntries);

            //if ((Body1 != null) && (Body1.Count() == 7))
            if ((Body1 != null) && (EmailMsg.Body.Content.Contains("You have received a new Request to work an upcoming event")))
            {
                int StaffID = -1;
                int EventID = -1;
                int EventStaffRequiredID = -1;
                foreach (var z in Body1)
                {
                    var y = z;
                    y = y.Replace(Constants.vbLf, "");

                    if ((y.Contains("Username:")) && ((StaffID == -1) || (StaffID == 0)))
                    {
                        string StaffEmail = Strings.LTrim(Strings.RTrim(Strings.Right(y, y.Length - 9)));
                        var _StaffID = DLookupSQLServerRemote("StaffItem", "EMSSecurityUser", "Email = '" + StaffEmail.Trim() + "'");
                        if ((_StaffID == null) || (_StaffID is System.DBNull))
                        {
                            StaffID = -1;
                        }
                        else
                        {
                            StaffID = (int)_StaffID;
                        }

                    }

                    if (y.Contains("Staff ID:"))
                    {
                        int iii;
                        var SID = Strings.LTrim(Strings.RTrim(Strings.Right(y, y.Length - 9)));
                        var isNumeric = int.TryParse(SID, out iii);
                        //StaffID = DLookupSQLServerRemote("StaffID", "StaffSecurityUser", "Email = '" + StaffEmail + "'");
                        if ((SID == null) || (!isNumeric))
                        {
                            StaffID = -1;
                        }
                        else
                        {
                            StaffID = iii;
                        }
                    }

                    if (y.Contains("Event ID:"))
                    {
                        int iii;
                        var EID = Strings.LTrim(Strings.RTrim(Strings.Right(y, 6)));
                        var isNumeric = int.TryParse(EID, out iii);
                        //StaffID = DLookupSQLServerRemote("StaffID", "StaffSecurityUser", "Email = '" + StaffEmail + "'");
                        if ((EID == null) || (!isNumeric))
                        {
                            EventID = -1;
                        }
                        else
                        {
                            EventID = iii;
                        }
                    }

                    if ((y.Contains("Event Ref:")) && ((EventID == -1) || (EventID == 0) || (EventID == null)))
                    {
                        string[] lines = y.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

                        string result = lines.SingleOrDefault(l => l.StartsWith("Event Ref:"));

                        //string EventRef = Strings.LTrim(Strings.RTrim(Strings.Right(y, y.Length - 10)));
                        string EventRef = Strings.LTrim(Strings.RTrim(Strings.Right(y, 8)));
                        var _EventID = DLookupSQLServerRemote("ID", "Events", "ID = " + EventRef);
                        if (_EventID == null)
                        {
                            EventID = -1;
                        }
                        else
                        {
                            EventID = (int)_EventID;
                        }
                    }


                    if (y.Contains("Event Staff Required ID:"))
                    {
                        int iii;
                        var ESRID = Strings.LTrim(Strings.RTrim(Strings.Right(y, 7)));
                        var isNumeric = int.TryParse(ESRID, out iii);
                        //StaffID = DLookupSQLServerRemote("StaffID", "StaffSecurityUser", "Email = '" + StaffEmail + "'");
                        if ((ESRID == null) || (!isNumeric))
                        {
                            EventStaffRequiredID = 0;
                        }
                        else
                        {
                            EventStaffRequiredID = iii;
                        }
                    }

                    if ((y.Contains("Role ID:")) && (EventStaffRequiredID == -1))
                    {
                        int iii;
                        var ESRID = Strings.LTrim(Strings.RTrim(Strings.Right(y, y.Length - 8)));
                        var isNumeric = int.TryParse(ESRID, out iii);
                        //StaffID = DLookupSQLServerRemote("StaffID", "StaffSecurityUser", "Email = '" + StaffEmail + "'");
                        if ((ESRID == null) || (!isNumeric))
                        {
                            EventStaffRequiredID = 0;
                        }
                        else
                        {
                            EventStaffRequiredID = iii;
                        }
                    }
                }


                LogMessage("Staff ID " + StaffID.ToString() + " Event ID " + EventID.ToString(), EmailMsg, TraceEventType.Information);

                if ((Conversions.ToInteger(StaffID) != -1) && (Conversions.ToInteger(EventID) != -1))
                {
                    var ClientID = DLookupAccess("Client ID", "Events", "[Event ID] = " + EventID.ToString());
                    if (ClientID == null)
                    {
                        ClientID = -1;
                    }
                    else
                    {
                        var SIDNP = DLookupAccess("StaffID", "Not Prefered Staff", "StaffID = " + StaffID.ToString() + " AND [Client ID] = " + ClientID.ToString());
                        if (SIDNP == null)
                        {
                            SIDNP = -1;
                        }
                        else
                        {
                            LogMessage("Staff ID " + StaffID.ToString() + " in not Preferred list for client id " + ClientID.ToString() + ". Skipping.", EmailMsg, TraceEventType.Information);
                            return;
                        }

                        var SIDP = DLookupAccess("StaffID", "Prefered Staff", "StaffID = " + StaffID.ToString() + " AND [ClientID] = " + ClientID.ToString());
                        if (SIDP == null)
                        {
                            SIDP = -1;
                        }
                        else
                        {
                            LogMessage("Staff ID " + StaffID.ToString() + " in Preferred list for client id " + ClientID.ToString() + ". Skipping.", EmailMsg, TraceEventType.Information);
                        }

                        string St1 = "SELECT Events.[Event ID] " + "FROM Events INNER JOIN [Staff Bookings] ON Events.[Event ID] = [Staff Bookings].[Event ID] " + "WHERE (((Events.[Event ID])=" + EventID.ToString() + ") AND (([Staff Bookings].[Staff ID])=" + StaffID.ToString() + ") AND ((Events.Status)=\"Current\") AND (([Staff Bookings].[Staff Total])>0) AND ((Events.Date)<=Date()))";
                        var EID = ExecuteScalar(St1, EmailMsg);
                        if (EID == null)
                        {
                            EID = -1;
                        }

                        LogMessage("EID " + EID.ToString(), EmailMsg, TraceEventType.Information);

                        var St2 = "INSERT INTO StaffWorkInterest ( EventID, StaffID, EventStaffRequiredID, P, W, Sort ) ";
                        St2 = St2 + "SELECT " + EventID.ToString() + ", " + StaffID.ToString() + ", " + EventStaffRequiredID.ToString() + ", \"" + Microsoft.VisualBasic.Interaction.IIf(Conversions.ToInteger(SIDP) != -1, "P", "") + "\", \"" + Microsoft.VisualBasic.Interaction.IIf(Conversions.ToInteger(EID) != 0, "W", "") + "\", " + Microsoft.VisualBasic.Interaction.IIf(Conversions.ToInteger(SIDP) != 0, "1", Microsoft.VisualBasic.Interaction.IIf(Conversions.ToInteger(EID) != 0, "2", "3"));

                        if (Conversions.ToInteger(ExecuteNonQuery(St2, EmailMsg)) >= 1)
                        {
                            St2 = "UPDATE StaffWorkInterest INNER JOIN Staff ON StaffWorkInterest.StaffID = Staff.[Staff ID] SET StaffWorkInterest.Staff = [Forename] & \" \" & [Surname];";

                            if (Conversions.ToInteger(ExecuteNonQuery(St2, EmailMsg)) >= 1)
                            {
                                St2 = "UPDATE StaffWorkInterest INNER JOIN [Event Staff Required] ON StaffWorkInterest.EventStaffRequiredID = [Event Staff Required].ID SET StaffWorkInterest.Role = [Event Staff Required].[Job Type], StaffWorkInterest.Start = [Event Staff Required].[Start], StaffWorkInterest.[End] = [Event Staff Required].[End];";

                                if (Conversions.ToInteger(ExecuteNonQuery(St2, EmailMsg)) >= 1)
                                {

                                    LogMessage("Staff " + StaffID.ToString() + ", Event " + EventID.ToString() + ", request to work saved.", EmailMsg, TraceEventType.Information);

                                    string FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") + " Staff " + StaffID.ToString() + " work request data for event " + EventID.ToString() + " saved.txt");
                                    try
                                    {
                                        TextWriter tw = System.IO.File.CreateText(FullPath);
                                        tw.Close();
                                    }
                                    catch (Exception ex)
                                    {
                                        LogMessage("ProcessStaffWorkRequestsMessageAsync::Staff " + StaffID.ToString() + ", Event " + EventID.ToString() + ", exception: " + ex.Message, EmailMsg, TraceEventType.Information);
                                    }
                                }
                                else
                                {
                                    LogMessage("Staff " + StaffID.ToString() + ", Event " + EventID.ToString() + ", UPDATE StaffWorkInterest INNER JOIN [Event Staff Required] failed.", EmailMsg, TraceEventType.Information);
                                }
                            }
                            else
                            {
                                LogMessage("Staff " + StaffID.ToString() + ", Event " + EventID.ToString() + ", UPDATE StaffWorkInterest INNER JOIN Staff failed.", EmailMsg, TraceEventType.Information);
                            }
                        }
                        else
                        {
                            LogMessage("Staff " + StaffID.ToString() + ", Event " + EventID.ToString() + ", request to work save failed.", EmailMsg, TraceEventType.Information);
                        }

                        string St3 = "UPDATE Events SET Events.JobRequests = IIf(IsNull([JobRequests]),1,[JobRequests]+1) " + "WHERE Events.[Event ID] = " + EventID.ToString();
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectGreaterEqual(ExecuteNonQuery(St3, EmailMsg), 1, false)))
                        {
                            LogMessage("Staff " + StaffID.ToString() + ", Event " + EventID.ToString() + ", event staff job request counter incremented.", EmailMsg, TraceEventType.Information);
                        }

                        string BodyText = Strings.Replace("Staff ID: " + StaffID.ToString() + Constants.vbCrLf + Constants.vbCrLf + "Event ID: " + EventID.ToString() + Constants.vbCrLf + Constants.vbCrLf + "=====================================" + Constants.vbCrLf + Constants.vbCrLf + EmailMsg.Body.Content, Constants.vbCrLf, "<BR>");

                        var subject = EmailMsg.Subject.Replace(MSG_SAVED, "").Replace(MSG_NOT_SAVED, "");

                        subject = MSG_SAVED + subject;

                        //EmailMsg.Body.Content = BodyText;

                        Microsoft.Graph.ItemBody b = new Microsoft.Graph.ItemBody();

                        b.Content = BodyText;

                        EmailMsg.IsDraft = true;

                        graphClient
                            .Users[user.Id]
                            .MailFolders.Inbox
                            .Messages[$"{EmailMsg.Id}"]
                            .Request()
                            .UpdateAsync(new Microsoft.Graph.Message()
                            {
                                Subject = subject,
                                Body = b,
                                IsRead = true
                            });

                        EmailMsg.IsDraft = false;
                    }
                }
                else
                {
                    LogMessage("Message " + EmailMsg.Subject + ", exit processing as StafF ID or Event ID not found.", EmailMsg, TraceEventType.Information);
                }

                // Catch ex As Exception
                // LogError("ProcessStaffWorkRequests Exception", message, ex.Message)
                // End Try
            }
        }

        private async Task ProcessTimesheetAttachmentsAsync()
        {
            var Users = await graphClient
                      .Users
                      .Request()
                      .Filter("startswith(mail,'Timesheets')")
                      .GetAsync();

            var user = Users[0];

            //var inboxMessages = await graphClient
            //     .Users[user.Id]
            //     .MailFolders.Inbox
            //     .Messages
            //     .Request()
            //     .Filter("not startswith(Subject, '" + MSG_SAVED + "')").Top(50)
            //     .GetAsync();

            var inboxMessages = await graphClient
                 .Users[user.Id]
                 .MailFolders.Inbox
                 .Messages
                 .Request()
                 .Filter("isRead eq false")
                 .Top(1000)
                 //.Select("Body, Subject")
                 .GetAsync();

            foreach (Microsoft.Graph.Message x in inboxMessages)
            {
                //if (x.IsRead == false)
                //if ((x.Subject != null) && (!x.Subject.Contains(MSG_NOT_SAVED)))
                //{
                await ProcessStaffTimesheetsAsync(x, user);
                //}
            }
        }

        private async Task ProcessStaffTimesheetsAsync(Microsoft.Graph.Message message, User user)
        {
            LogMessage("", null, TraceEventType.Information);

            LogMessage("Message " + message.Subject + " loaded. Ready for attachment processing.", message, TraceEventType.Information);

            bool FDFAttchmentFound = false;

            if (message.HasAttachments == true)
            {
                var attachments = await graphClient
                                    .Users[user.Id]
                                    .Messages[$"{message.Id}"]
                                    .Attachments
                                    .Request()
                                    .GetAsync();

                LogMessage("Iterating through " + attachments.Count.ToString() + " attachments.", message, TraceEventType.Information);

                //foreach (var attachment in message.Attachments)
                foreach (var attachment in attachments)
                {
                    //if (attachment.ODataType == "#microsoft.graph.itemAttachment")
                    //{

                    //    var attachmentRequest = await graphClient
                    //                                .Users[user.Id]
                    //                                .MailFolders
                    //                                .Inbox
                    //                                .Messages[message.Id]
                    //                                .Attachments[attachment.Id]
                    //                                .Request()
                    //                                .Expand("microsoft.graph.itemattachment/item")
                    //                                .GetAsync();

                    //    var itemAttachment = (ItemAttachment)attachmentRequest;
                    //}


                    //var a1 = await graphClient
                    //           .Users[user.Id]
                    //           .Messages[$"{message.Id}"]
                    //           .Attachments[$"{attachment.Id}"]
                    //           .Request()
                    //           .GetAsync();

                    if (attachment is FileAttachment)
                    {
                        FileAttachment fileAttachment = attachment as FileAttachment;
                        LogMessage("Processing attachment " + fileAttachment.Name, message, TraceEventType.Information);
                        if (fileAttachment.Name.Contains(".FDF") | fileAttachment.Name.Contains(".fdf"))
                        {
                            FDFAttchmentFound = true;

                            LogMessage("File attachment name: " + fileAttachment.Name, message, TraceEventType.Information);
                            if (fileAttachment.Name.Length >= 3)
                            {
                                string feFileExtension = fileAttachment.Name.Substring(fileAttachment.Name.Length - 4, 4);
                                string feFileName = fileAttachment.Name.Replace("%20", " ");
                                LogMessage("Attachment Ext: " + feFileExtension, message, TraceEventType.Information);
                                if (feFileExtension.ToLower() == ".fdf")
                                {
                                    message.Subject = message.Subject.Replace(MSG_SAVED, "");
                                    message.Subject = message.Subject.Replace(MSG_NOT_SAVED, "");
                                    PDFType PT = default;
                                    string fePDFFileName = GetPDFFileName(feFileName, message, ref PT);
                                    if (fePDFFileName != null)
                                    {
                                        LogMessage("Saving attachment, file name: " + feFileName, message, TraceEventType.Information);

                                        //fileAttachment.Load(Path.Combine(TIMESHEETS_DATA_PATH, feFileName));

                                        SaveAttachment(fileAttachment, Path.Combine(TIMESHEETS_DATA_PATH, feFileName));

                                        LogMessage("Attachment saved.", message, TraceEventType.Information);
                                        LogMessage("Updating db...", message, TraceEventType.Information);
                                        UpdateEditableTiemsheetAccess(message, fePDFFileName, feFileName);
                                        UpdateEditableTiemsheetSQLServer(message, fePDFFileName, feFileName);

                                        SetMessageSubject(message, user, MSG_SAVED);
                                    }
                                    else if (System.IO.File.Exists(REFERENCES_PDF_PATH + fePDFFileName))
                                    {
                                        LogMessage("Saving reference attachment, file name: " + feFileName, message, TraceEventType.Information);

                                        //fileAttachment.Load(Path.Combine(REFERENCES_DATA_PATH, feFileName));

                                        SaveAttachment(fileAttachment, Path.Combine(REFERENCES_DATA_PATH, feFileName));

                                        LogMessage("Reference attachment saved.", message, TraceEventType.Information);
                                        LogMessage("Updating Reference db.", message, TraceEventType.Information);
                                        UpdateReferencesAccess(message, fePDFFileName, feFileName);
                                        UpdateReferencesSQLServer(message, fePDFFileName, feFileName);

                                        SetMessageSubject(message, user, MSG_SAVED);
                                    }
                                    else
                                    {
                                        LogMessage("Data file " + feFileName + " not saved, no matching pdf file '" + TIMESHEETS_PDF_PATH + fePDFFileName + " found.", message, TraceEventType.Error);
                                        LogError("Data file " + feFileName + " not saved, no matching pdf file '" + TIMESHEETS_PDF_PATH + fePDFFileName + " found.", message, TraceEventType.Error);

                                        SetMessageSubject(message, user, MSG_NOT_SAVED);
                                    }
                                }
                            }

                            fileAttachment = null;
                        }
                        else
                        {
                            LogMessage("Attachment " + fileAttachment.Name + " is an invalid item attachment type, skipping.", null, TraceEventType.Information);
                        }

                    }
                    else
                    {
                        LogMessage("Attachment " + attachment.Name + "is an item attachment, skipping attachment.", message, TraceEventType.Information);
                    }
                }
            }
            else
            {
                SetMessageSubject(message, user, MSG_NOT_SAVED);

                return;
            }

            if (!FDFAttchmentFound)
            {
                LogMessage("No attachment in the message valid for processing, skipping message.", message, TraceEventType.Information);

                SetMessageSubject(message, user, MSG_NOT_SAVED);

                return;
            }
        }

        private void SaveAttachment(FileAttachment fileAttachment, string p)
        {
            //using (FileStream outputFileStream = new FileStream(p, FileMode.Create))
            //{
            //    fileAttachment.ContentBytes.CopyTo(outputFileStream);

            //    fileAttachment.ContentBytes.CopyTo()
            //}

            //Buffer fileContent = new Buffer(fileAttachment.ContentBytes);

            System.IO.File.WriteAllBytes(p, fileAttachment.ContentBytes);
        }

        private void SetMessageSubject(Microsoft.Graph.Message message, User user, string v)
        {
            var subject = message.Subject.Replace(MSG_SAVED, "").Replace(MSG_NOT_SAVED, "").Replace(v, "");

            subject = v + subject;

            message.IsDraft = true;

            graphClient
                .Users[user.Id]
                .MailFolders.Inbox
                .Messages[$"{message.Id}"]
                .Request()
                .UpdateAsync(new Microsoft.Graph.Message()
                {
                    Subject = subject,
                    IsRead = true
                });

            message.IsDraft = false;
        }

        public async void Process()
        {
            this.Timer1.Enabled = false;
            CurrentTimeStamp = DateTime.Now;
            this.btnProcess.Enabled = false;
            InProcess = true;
            //ProcessTimesheetAttachments();
            //ProcessStaffWorkRequests();

            DateTime LastRun = DateTime.Now;

            if (isProcessUnsubscribeEmailsOK())
            {
                var config = LoadAppSettings();
                if (null == config)
                {
                    Console.WriteLine("Missing or invalid appsettings.json file. Please see README.md for configuration instructions.");
                    return;
                }

                graphClient = GetAuthenticatedGraphClient(config);

                //LogMessage("", null, TraceEventType.Information, "", true);


                LogMessage("Begin unsubscribe staff processing...", null, TraceEventType.Information, "", true);

                await ProcessUnsubscribeEmailsAsync(graphClient);

                //LogMessage("End unsubscribe staff processing...", null, TraceEventType.Information, "", true);
                //LogMessage(" ", null, TraceEventType.Information, "", true);


                LogMessage("Begin processing staff work requests...", null, TraceEventType.Information, "", true);

                await ProcessStaffWorkRequests(graphClient);


                LogMessage("Begin processing timesheet attachments...", null, TraceEventType.Information, "", true);

                await ProcessTimesheetAttachmentsAsync();


                //Properties.Settings.Default.UnsubscribeServiceLastRun = DateTime.Now.ToOADate();
                Properties.Settings.Default.UnsubscribeServiceLastRun = LastRun.ToOADate();

                double LastRun1 = Properties.Settings.Default.UnsubscribeServiceLastRun;

                Properties.Settings.Default.Save();
                Properties.Settings.Default.Upgrade();
                //System.Windows.Forms.Application.Restart();

                graphClient = null;
            }

            InProcess = false;
            this.btnProcess.Enabled = true;
            System.Windows.Forms.Application.DoEvents();
            this.Timer1.Enabled = true;
        }

        private bool isProcessUnsubscribeEmailsOK()
        {
            return true;

            //DateTime UnsubscribeServiceLastRun = DateTime.FromOADate(Properties.Settings.Default.UnsubscribeServiceLastRun);

            ////int UnsubscribeServiceInterval = (int)Math.Round(60000f * Properties.Settings.Default.UnsubscribeServiceTimeinMinutes);

            //if (DateTime.Now >= UnsubscribeServiceLastRun.AddMinutes(Properties.Settings.Default.UnsubscribeServiceTimeinMinutes))
            //    return true;
            //else
            //    return false;
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            this.Timer1.Enabled = false;

            if (InProcess)
            {
                return;
            }

            Process();

            Timer1.Interval = (int)Math.Round(60000f * Properties.Settings.Default.TimerTimeinMinutes);

            this.Timer1.Enabled = true;
        }

        //private static GraphServiceClient GetClient(string accessToken, IHttpProvider provider = null)
        //{
        //    var delegateAuthProvider = new DelegateAuthenticationProvider((requestMessage) =>
        //    {
        //        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

        //        return Task.FromResult(0);
        //    });

        //    var graphClient = new GraphServiceClient(delegateAuthProvider, provider ?? HttpProvider);

        //    return graphClient;
        //}

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _graphServiceClient = new GraphServiceClient(authenticationProvider);
            return _graphServiceClient;
        }

        //private static HttpClient GetAuthenticatedHTTPClient(IConfigurationRoot config)
        //{
        //    var authenticationProvider = CreateAuthorizationProvider(config);
        //    _httpClient = new HttpClient(new AuthHandler(authenticationProvider, new HttpClientHandler()));
        //    return _httpClient;
        //}

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var clientId = config["applicationId"];
            var clientSecret = config["applicationSecret"];
            var redirectUri = config["redirectUri"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

            //this specific scope means that application will default to what is defined in the application registration rather than using dynamic scopes
            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                                                    .WithAuthority(authority)
                                                    .WithRedirectUri(redirectUri)
                                                    .WithClientSecret(clientSecret)
                                                    .Build();
            return new MsalAuthenticationProvider(cca, scopes.ToArray());
        }

        public static IConfigurationRoot LoadAppSettings()
        {
            try
            {
                var config = new ConfigurationBuilder()
                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", false, true)
                .Build();

                // Validate required settings
                if (string.IsNullOrEmpty(config["applicationId"]) ||
                    string.IsNullOrEmpty(config["applicationSecret"]) ||
                    string.IsNullOrEmpty(config["redirectUri"]) ||
                    string.IsNullOrEmpty(config["tenantId"]) ||
                    string.IsNullOrEmpty(config["domain"]))
                {
                    return null;
                }

                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }

        private void frmEmailProcessor_Load(object sender, EventArgs e)
        {
            this.Timer1.Interval = (int)Math.Round(60000f * Properties.Settings.Default.TimerTimeinMinutes);
            //this.Timer1.Enabled = false;
            this.Timer1.Enabled = true;
        }

        object ExecuteScalar(string St, Microsoft.Graph.Message message = null)
        {
            var Conn = new OleDbConnection(LocalAccessConnSt);
            try
            {

                //LogMessage("Opening ExecuteScalar Access db connection.", message, TraceEventType.Information);
                Conn.Open();
                //LogMessage("ExecuteScalar Access db connection opened.", message, TraceEventType.Information);
                var Command = new OleDbCommand(St, Conn);
                var j = Command.ExecuteScalar();
                Command.Dispose();
                return j;
            }
            catch (Exception ex)
            {
                LogError("Error ExecuteScalar Access db: " + ex.Message, message, TraceEventType.Error);
            }
            finally
            {
                Conn.Close();
            }

            return null;
        }

        object ExecuteNonQuery(string St, Microsoft.Graph.Message message = null)
        {
            var Conn = new OleDbConnection(LocalAccessConnSt);
            try
            {
                //LogMessage("Opening ExecuteNonQuery Access db connection.", message, TraceEventType.Information);
                Conn.Open();
                //LogMessage("ExecuteNonQuery Access db connection opened.", message, TraceEventType.Information);
                var Command = new OleDbCommand(St, Conn);
                int j = Command.ExecuteNonQuery();
                Command.Dispose();
                return j;
            }
            catch (Exception ex)
            {
                LogError("Error ExecuteNonQuery Access db. St: " + St + ", Exception: " + ex.Message, message, TraceEventType.Error);
            }
            finally
            {
                Conn.Close();
            }

            return null;
        }

        object ExecuteNonQuerySQLServer(string St, Microsoft.Graph.Message message = null)
        {
            var Conn = new SqlConnection(LocalSQLServerConnSt);
            try
            {
                //LogMessage("Opening ExecuteNonQuery SQL Server db connection.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                Conn.Open();
                //LogMessage("ExecuteNonQuery SQL Server db connection opened.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                var Command = new SqlCommand(St, Conn);
                int j = Command.ExecuteNonQuery();
                Command.Dispose();
                return j;
            }
            catch (Exception ex)
            {
                LogError("Error ExecuteNonQuery SQL Server db. St: " + St + ", Exception: " + ex.Message, message, TraceEventType.Error);
            }
            finally
            {
                Conn.Close();
            }

            return null;
        }

        object DLookupAccess(string Value, string Table, string Condition = "", Microsoft.Graph.Message message = null)
        {
            var Conn = new OleDbConnection(LocalAccessConnSt);

            string St = "SELECT [" + Value + "] FROM [" + Table + "] " + (string.IsNullOrEmpty(Condition) ? "" : "WHERE " + Condition).Replace("[[", "[").Replace("]]", "]");

            try
            {
                //LogMessage("Opening DLookup Access db connection.", message, TraceEventType.Information);
                Conn.Open();
                //LogMessage("DLookup Access db connection opened.", message, TraceEventType.Information);
                var Command = new OleDbCommand(St, Conn);
                var j = Command.ExecuteScalar();
                Command.Dispose();
                return j;
            }
            catch (Exception ex)
            {
                LogError("Error ExecuteScalar Access db. St: " + St + ", Exception: " + ex.Message, message, TraceEventType.Error);
                return null;
            }
            finally
            {
                Conn.Close();
            }
        }

        object DLookupSQLServer(string Value, string Table, string Condition = "", Microsoft.Graph.Message message = null)
        {
            var Conn = new SqlConnection(LocalSQLServerConnSt);

            string St = "SELECT [" + Value + "] FROM [" + Table + "] " + (string.IsNullOrEmpty(Condition) ? "" : "WHERE " + Condition);

            try
            {
                //LogMessage("Opening Dlookup SQL Server db connection.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                Conn.Open();
                //LogMessage("Dlookup SQL Server db connection opened.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                var Command = new SqlCommand(St, Conn);
                // Dim j = Command.ExecuteScalar(St, Conn)
                var j = Command.ExecuteScalar();
                Command.Dispose();
                return j;
            }
            catch (Exception ex)
            {
                LogError("Error dlookup SQL Server db. St: " + St + ", Exception: " + ex.Message, message, TraceEventType.Error);
                return null;
            }

            finally
            {
                Conn.Close();
            }
        }

        object DLookupSQLServerRemote(string Value, string Table, string Condition = "", Microsoft.Graph.Message message = null)
        {
            var Conn = new SqlConnection(RemoteSQLServerConnSt);

            string St = "SELECT [" + Value + "] FROM [" + Table + "] " + (string.IsNullOrEmpty(Condition) ? "" : "WHERE " + Condition);

            try
            {
                //LogMessage("Opening Dlookup SQL Server db connection.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                Conn.Open();
                //LogMessage("Dlookup SQL Server db connection opened.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                var Command = new SqlCommand(St, Conn);
                // Dim j = Command.ExecuteScalar(St, Conn)
                var j = Command.ExecuteScalar();
                Command.Dispose();
                return j;
            }
            catch (Exception ex)
            {
                LogError("Error dlookup SQL Server db. St: " + St + ", Exception: " + ex.Message, message, TraceEventType.Error);
                return null;
            }
            finally
            {
                Conn.Close();
            }
        }

        public string StripHtmlTags(string html)
        {

            // Remove HTML tags.
            return Regex.Replace(html, "<.*?>", "");
        }
    }
}
