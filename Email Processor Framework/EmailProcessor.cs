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
using System.Net.Http.Headers;
using System.Net.Http;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Constants = Microsoft.VisualBasic.Constants;
using static Email_Processor.Utils;
using System.IO.Compression;
using System.Globalization;
using DevExpress.Data.Linq.Helpers;
using static System.Net.WebRequestMethods;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Reflection;
using static Email_Processor_Framework.EmailProcessor;
//using System.IO.Compression.FileSystem;

namespace Email_Processor_Framework
{
    static class EmailProcessor
    {
        private static GraphServiceClient _graphServiceClient;
        static private GraphServiceClient graphClient = null;
        private const string domain = "high-society.co.uk";

        static public bool InProcess = false;

        public enum PDFType
        {
            None = -1,
            Timesheet = 0,
            Reference = 1
        }

        static private SecureString ConvertToSecureString(string password)
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

        static private async Task ProcessRecruitmentEmailsAsync(GraphServiceClient graphClient)
        {
            List<QueryOption> options = new List<QueryOption>
            {
                //new QueryOption("$top", "1")
            };

            var Users = await graphClient
                         .Users
                         .Request()
                         .Filter("startswith(displayName,'Staffing')").Top(1)
                         .GetAsync();

            var user = Users[0];

            var inboxMessages = await graphClient
                      .Users[user.Id]
                      .MailFolders.Inbox
                      .Messages
                      .Request()
                      .Filter("receivedDateTime gt 2025-02-23T00:00:00Z")
                      .Top(1000)
                      .GetAsync();

            foreach (Microsoft.Graph.Message x in inboxMessages)
            {
                try
                {
                    await ProcessRecruitmentMessageAsync(x, user);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("ERROR: " + x.Sender.EmailAddress.Address);
                    MessageBox.Show(ex.Message, x.Sender.EmailAddress.Address);
                }
            }
        }

        static private async Task ProcessRecruitmentMessageAsync(Microsoft.Graph.Message message, User user)
        {
            string emailBody = message.Body.Content;
            Dictionary<string, string> applicantData = ParseEmailBody(emailBody);

            await InsertApplicantIntoDatabase(applicantData, message.ReceivedDateTime.Value.DateTime);
        }

        private static Dictionary<string, string> ParseEmailBody(string emailBody)
        {
            // Replace <br> tags with newline characters.
            string cleaned = Regex.Replace(emailBody, @"<br\s*/?>", "\n", RegexOptions.IgnoreCase);
            // Normalize Windows newlines to Unix-style.
            cleaned = cleaned.Replace("\r\n", "\n");

            // Remove any remaining HTML tags.
            cleaned = Regex.Replace(cleaned, "<.*?>", string.Empty);

            // Split the cleaned string into lines.
            string[] lines = cleaned.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);

            // Use a case-insensitive dictionary.
            Dictionary<string, string> data = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            // Separate collection for job types.
            List<string> jobTypes = new List<string>();

            foreach (string line in lines)
            {
                string trimmedLine = line.Trim();
                // Find the first colon to separate key and value.
                int colonIndex = trimmedLine.IndexOf(":");
                if (colonIndex > -1)
                {
                    string key = trimmedLine.Substring(0, colonIndex).Trim();
                    string value = trimmedLine.Substring(colonIndex + 1).Trim();
                    if (!string.IsNullOrEmpty(key))
                    {
                        // If the key is numeric, assume it's a job type.
                        if (int.TryParse(key, out _))
                        {
                            jobTypes.Add(value);
                        }
                        else
                        {
                            data[key] = value;
                        }
                    }
                }
            }

            // Aggregate job types under the "JobType" key.
            if (jobTypes.Count > 0)
            {
                data["JobType"] = string.Join(", ", jobTypes);
            }

            return data;
        }

        private static async Task InsertApplicantIntoDatabase(Dictionary<string, string> data, DateTime emailDate)
        {
            string connectionString = LocalSQLServerConnSt;
            string checkQuery = "SELECT COUNT(1) FROM Applicants WHERE Email = @Email AND PreviousExperience = @PreviousExperience";
            string insertQuery = @"INSERT INTO Applicants (Username, Forenames, Surname, DOB, Email, Tel, Mobile, Address1, Address2, Address3, Town, County, Postcode, Source, JobType, ProfessionalStatus, PreviousExperience, RefereeName1, RefereePhone1, RefereeEmail1, RefereeName2, RefereePhone2, RefereeEmail2, ApplyDate, Stage) 
                      VALUES (@Username, @Forenames, @Surname, @DOB, @Email, @Tel, @Mobile, @Address1, @Address2, @Address3, @Town, @County, @Postcode, @Source, @JobType, @ProfessionalStatus, @PreviousExperience, @RefereeName1, @RefereePhone1, @RefereeEmail1, @RefereeName2, @RefereePhone2, @RefereeEmail2, @ApplyDate, @Stage)";

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
            {
                checkCmd.Parameters.AddWithValue("@Email", data.ContainsKey("email") ? data["email"] : DBNull.Value.ToString());
                checkCmd.Parameters.AddWithValue("@PreviousExperience", data.ContainsKey("PreviousExperience") ? data["PreviousExperience"] : DBNull.Value.ToString());

                await conn.OpenAsync();
                //int exists = (int)await checkCmd.ExecuteScalarAsync();

                //if (exists > 0)
                //{
                //    return;
                //}

                using (SqlCommand insertCmd = new SqlCommand(insertQuery, conn))
                {
                    insertCmd.Parameters.AddWithValue("@Username", data.ContainsKey("email") ? data["email"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@Forenames", data.ContainsKey("forename") ? data["forename"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@Surname", data.ContainsKey("surname") ? data["surname"] : DBNull.Value.ToString());
                    //insertCmd.Parameters.AddWithValue("@DOB", data.ContainsKey("dob") ? Convert.ToDateTime(data["dob"]) : DBNull.Value);

                    insertCmd.Parameters.AddWithValue("@DOB", data.ContainsKey("dob") ? (object)Convert.ToDateTime(data["dob"]) : DBNull.Value);

                    insertCmd.Parameters.AddWithValue("@Email", data.ContainsKey("email") ? data["email"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@Tel", data.ContainsKey("tel") ? data["tel"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@Mobile", data.ContainsKey("mobile") ? data["mobile"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@Address1", data.ContainsKey("Address1") ? data["Address1"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@Address2", data.ContainsKey("Address2") ? data["Address2"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@Address3", data.ContainsKey("Address3") ? data["Address3"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@Town", data.ContainsKey("Town") ? data["Town"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@County", data.ContainsKey("County") ? data["County"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@Postcode", data.ContainsKey("Postcode") ? data["Postcode"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@Source", data.ContainsKey("Source") ? data["Source"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@JobType", data.ContainsKey("JobType") ? data["JobType"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@ProfessionalStatus", data.ContainsKey("ProfessionalStatus") ? data["ProfessionalStatus"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@PreviousExperience", data.ContainsKey("PreviousExperience") ? data["PreviousExperience"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@RefereeName1", data.ContainsKey("RefereeName1") ? data["RefereeName1"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@RefereePhone1", data.ContainsKey("RefereePhone1") ? data["RefereePhone1"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@RefereeEmail1", data.ContainsKey("RefereeEmail1") ? data["RefereeEmail1"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@RefereeName2", data.ContainsKey("RefereeName2") ? data["RefereeName2"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@RefereePhone2", data.ContainsKey("RefereePhone2") ? data["RefereePhone2"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@RefereeEmail2", data.ContainsKey("RefereeEmail2") ? data["RefereeEmail2"] : DBNull.Value.ToString());
                    insertCmd.Parameters.AddWithValue("@ApplyDate", emailDate);
                    insertCmd.Parameters.AddWithValue("@Stage", "Applicant");

                    try
                    {
                        await insertCmd.ExecuteNonQueryAsync();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine("ERROR: " + (data.ContainsKey("email") ? data["email"] : DBNull.Value.ToString()));
                        MessageBox.Show(ex.Message, data.ContainsKey("email") ? data["email"] : DBNull.Value.ToString());
                    }
                }
            }
        }

        #region "ProcessUnsubscribeEmailsAsync"

        static private async Task ProcessUnsubscribeEmailsAsync(GraphServiceClient graphClient)
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

        static private async Task ProcessUnsubscribeMessageAsync(Microsoft.Graph.Message message, User user)
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

        static private void UnsubcribeStaff(Microsoft.Graph.Message message, Microsoft.Graph.Message originalmessage, User user)
        {
            LogMessage("Processing unsubscribe request.", message, TraceEventType.Information);
            string Email = message.Sender.EmailAddress.Address;
            string StaffName = message.Sender.EmailAddress.Name;
            string BodyText = message.Body.Content.Replace("\"", "\"\"");
            //return UnsubscribeStaffAccess(Email, StaffName, BodyText, message) | UnsubscribeStaffSQLServer(Email, StaffName, BodyText, message);
            UnsubscribeStaff(Email, StaffName, BodyText, message, originalmessage, user);
            LogMessage(" ", message, TraceEventType.Information, "", true);
        }

        static private void UnsubscribeStaff(string Email, string StaffName, string BodyText, Microsoft.Graph.Message message, Microsoft.Graph.Message originalmessage, User user)
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

        static private UnsubscribeStatus UnsubscribeStaffAccess(string Email, string StaffName, string BodyText, Microsoft.Graph.Message message)
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

        static private UnsubscribeStatus UnsubscribeStaffSQLServer(string Email, string StaffName, string BodyText, Microsoft.Graph.Message message)
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

        static void UpdateReferencesSQLServer(Microsoft.Graph.Message message, string fePDFFileName, string feFileName)
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

        static void UpdateEditableTiemsheetSQLServer(Microsoft.Graph.Message message, string fePDFFileName, string feFileName)
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

        static void UpdateReferencesAccess(Microsoft.Graph.Message message, string fePDFFileName, string feFileName)
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

        static void UpdateEditableTiemsheetAccess(Microsoft.Graph.Message message, string fePDFFileName, string feFileName)
        {
            try
            {
                string ConnSt = Email_Processor_Framework.Properties.Settings.Default.DBConnection;

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

        static string GetPDFFileName(string feFileName, Microsoft.Graph.Message item, ref PDFType PT)
        {
            LogMessage($"Starting PDF search for FDF: {feFileName}", item, TraceEventType.Information);

            // --- Level 1: Exact Match Attempt ---
            LogMessage("Attempting Level 1: Exact Match", item, TraceEventType.Verbose);
            string cleanedFdfBase = CleanFdfBaseName(feFileName);
            if (!string.IsNullOrEmpty(cleanedFdfBase))
            {
                string potentialExactPdfName = cleanedFdfBase + ".pdf";
                string potentialExactPdfPath = Path.Combine(TIMESHEETS_PDF_PATH, potentialExactPdfName);

                LogMessage($"Checking for exact PDF match: {potentialExactPdfPath}", item, TraceEventType.Verbose);
                if (System.IO.File.Exists(potentialExactPdfPath))
                {
                    LogMessage($"Level 1 Success: Found exact match PDF: {potentialExactPdfName}", item, TraceEventType.Information);
                    PT = PDFType.Timesheet; // Assuming type based on path
                    return potentialExactPdfName;
                }
                else
                {
                    LogMessage($"Level 1 Failed: Exact match not found for '{potentialExactPdfName}'", item, TraceEventType.Verbose);
                }
            }
            else
            {
                LogMessage("Level 1 Skipped: Could not clean FDF filename.", item, TraceEventType.Warning);
            }

            // --- Level 2: Event ID + Date Match Attempt ---
            LogMessage("Attempting Level 2: Event ID + Date Match", item, TraceEventType.Verbose);
            string fdfEventId = ExtractEventId(feFileName);
            string fdfDateString = ExtractDateString(feFileName); // Expects dd-MM-yyyy

            if (!string.IsNullOrEmpty(fdfEventId) && !string.IsNullOrEmpty(fdfDateString))
            {
                LogMessage($"Extracted from FDF: EventID='{fdfEventId}', Date='{fdfDateString}'", item, TraceEventType.Verbose);

                string matchedPdf = FindPdfByCriteria(pdfFileName => {
                    string pdfEventId = ExtractEventId(pdfFileName);
                    string pdfDateString = ExtractDateString(pdfFileName);
                    // Match if both Event ID and Date String match
                    return pdfEventId == fdfEventId && pdfDateString == fdfDateString;
                }, TIMESHEETS_PDF_PATH, item, ref PT);

                if (matchedPdf != null)
                {
                    LogMessage($"Level 2 Success: Found match by Event ID and Date: {matchedPdf}", item, TraceEventType.Information);
                    return matchedPdf;
                }
                else
                {
                    LogMessage("Level 2 Failed: No PDF found matching both Event ID and Date.", item, TraceEventType.Verbose);
                }
            }
            else
            {
                LogMessage($"Level 2 Skipped: Could not extract both Event ID ('{fdfEventId ?? "null"}') and Date ('{fdfDateString ?? "null"}') from FDF: {feFileName}", item, TraceEventType.Verbose);
            }

            // --- Level 3: Event ID Match Attempt ---
            LogMessage("Attempting Level 3: Event ID Only Match", item, TraceEventType.Verbose);
            // Reuse fdfEventId extracted earlier
            if (!string.IsNullOrEmpty(fdfEventId))
            {
                LogMessage($"Using extracted EventID='{fdfEventId}' for Level 3 search.", item, TraceEventType.Verbose);

                string matchedPdf = FindPdfByCriteria(pdfFileName => {
                    string pdfEventId = ExtractEventId(pdfFileName);
                    // Match if Event ID matches
                    return pdfEventId == fdfEventId;
                }, TIMESHEETS_PDF_PATH, item, ref PT);

                if (matchedPdf != null)
                {
                    LogMessage($"Level 3 Success: Found match by Event ID only: {matchedPdf}", item, TraceEventType.Information);
                    return matchedPdf;
                }
                else
                {
                    LogMessage("Level 3 Failed: No PDF found matching Event ID.", item, TraceEventType.Verbose);
                }
            }
            else
            {
                // This log message might be redundant if Level 2 already logged it, but good for clarity
                LogMessage($"Level 3 Skipped: Could not extract Event ID from FDF: {feFileName}", item, TraceEventType.Verbose);
            }

            // --- All Attempts Failed ---
            LogMessage($"All matching attempts failed for FDF: {feFileName}. No corresponding PDF found in {TIMESHEETS_PDF_PATH}", item, TraceEventType.Warning);
            return null; // Indicate failure
        }

        // Helper to clean FDF filename for exact matching
        static private string CleanFdfBaseName(string fdfFileName)
        {
            if (string.IsNullOrEmpty(fdfFileName)) return null;

            string baseName = fdfFileName;

            // Remove known suffix (case-insensitive)
            if (baseName.EndsWith("_data.fdf", StringComparison.OrdinalIgnoreCase))
            {
                baseName = baseName.Substring(0, baseName.Length - "_data.fdf".Length);
            }
            else if (baseName.EndsWith(".fdf", StringComparison.OrdinalIgnoreCase))
            {
                // Fallback if only .fdf is present (shouldn't happen based on description but good practice)
                baseName = baseName.Substring(0, baseName.Length - ".fdf".Length);
            }

            // Remove trailing parenthesized numbers like (005) before the extension was removed
            // Regex: Matches optional whitespace, followed by '(', one or more digits, ')' at the end of the string.
            baseName = Regex.Replace(baseName, @"\s*\(\d+\)$", "").Trim();

            return baseName;
        }

        // Helper to extract the first number (Event ID)
        static private string ExtractEventId(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return null;
            var match = Regex.Match(fileName.Trim(), @"^\d+");
            return match.Success ? match.Value : null;
        }

        // Helper to extract the date string (dd-MM-yyyy)
        // Using a slightly more robust regex to handle potential single digits and slight variations if needed,
        // but sticking to the requested dd-MM-yyyy format primarily.
        static private string ExtractDateString(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return null;
            // Regex specifically for dd-MM-yyyy or d-M-yyyy format within the string
            // Using non-capturing group for the date pattern itself
            var match = Regex.Match(fileName, @"\b(\d{1,2}-\d{1,2}-\d{4})\b");

            // Optional: Add validation if needed (e.g., check if DateTime.TryParseExact works)
            // if (match.Success && DateTime.TryParseExact(match.Value, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
            // {
            //     return match.Value;
            // }
            return match.Success ? match.Value : null;
        }

        // Helper to find a PDF matching specific criteria by iterating through the directory
        static private string FindPdfByCriteria(Func<string, bool> criteria, string pdfDirectoryPath, Microsoft.Graph.Message item, ref PDFType pt)
        {
            try
            {
                if (!System.IO.Directory.Exists(pdfDirectoryPath))
                {
                    LogMessage($"PDF directory not found: {pdfDirectoryPath}", item, TraceEventType.Warning);
                    return null;
                }

                // Consider adjusting SearchOption if subdirectories are needed
                foreach (var pdfFilePath in System.IO.Directory.EnumerateFiles(pdfDirectoryPath, "*.pdf", SearchOption.TopDirectoryOnly))
                {
                    string pdfFileName = Path.GetFileName(pdfFilePath);
                    if (criteria(pdfFileName))
                    {
                        LogMessage($"Found matching PDF by criteria: {pdfFileName}", item, TraceEventType.Information);
                        pt = PDFType.Timesheet; // Assuming PDFs in TIMESHEETS_PDF_PATH are Timesheets
                        return pdfFileName;
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error searching for PDF files in {pdfDirectoryPath}: {ex.Message}", item, TraceEventType.Error);
                LogError($"Error searching for PDF files in {pdfDirectoryPath}: {ex.ToString()}", item, TraceEventType.Error); // Log full exception details
            }
            return null; // No match found
        }

        //static string GetPDFFileName(string feFileName, Microsoft.Graph.Message item, ref PDFType PT)
        //{
        //    string fePDFFileName = "";
        //    string fePDFFileName1 = "";
        //    //System.Windows.Forms.Application.DoEvents();

        //    var r = new Regex(@"\d{1,2}\-\d{1,2}\-\d{4}", RegexOptions.IgnoreCase);
        //    var m = r.Match(feFileName);
        //    if (m.Success)
        //    {
        //        int i;
        //        if (feFileName.IndexOf("]") > 0)
        //        {
        //            i = m.Index + 11;
        //            fePDFFileName = Strings.Left(feFileName, i) + ".pdf";
        //            LogMessage("feFileName.IndexOf(\"]\") > 0: feFileName: " + feFileName + ", fePDFFileName: " + fePDFFileName, item, TraceEventType.Information);
        //        }
        //        else
        //        {
        //            string DateSt = Mid(feFileName, m.Index, 10);
        //            i = m.Index + 10;
        //            fePDFFileName = Strings.Left(feFileName, i) + ".pdf";
        //            fePDFFileName1 = fePDFFileName.Replace(DateSt, "[" + DateSt + "]");
        //            LogMessage("feFileName.IndexOf(\"]\") <= 0: feFileName: " + feFileName + ", fePDFFileName: " + fePDFFileName, item, TraceEventType.Information);
        //        }

        //        LogMessage("fePDFFileNamet: " + fePDFFileName + ", fePDFFileName1: " + fePDFFileName1, item, TraceEventType.Information);

        //        if (System.IO.File.Exists(Path.Combine(TIMESHEETS_PDF_PATH, fePDFFileName)))
        //        {
        //            LogMessage("GetPDFFileName: Exists: " + fePDFFileName, item, TraceEventType.Information);
        //            PT = PDFType.Timesheet;
        //            return fePDFFileName;
        //        }
        //        else if (System.IO.File.Exists(Path.Combine(TIMESHEETS_PDF_PATH, fePDFFileName1)))
        //        {
        //            LogMessage("GetPDFFileName: Exists: " + fePDFFileName1, item, TraceEventType.Information);
        //            PT = PDFType.Timesheet;
        //            return fePDFFileName1;
        //        }
        //        else
        //        {
        //            return null;
        //        }
        //    }
        //    else
        //    {
        //        // No date found
        //        LogMessage("GetPDFFileName: No match, feFileName: " + feFileName, item, TraceEventType.Information);
        //        return null;
        //    }

        //    // If feFileName.IndexOf("]") > 0 Then
        //    // fePDFFileName = LeftStr(feFileName, feFileName.IndexOf("]")).ToLower().Replace("_data", "") & ".pdf"
        //    // LogMessage("feFileName.IndexOf(""]"") > 0: feFileName: " & feFileName & ", fePDFFileName: " & fePDFFileName, Message, TraceEventType.Information)
        //    // Else
        //    // fePDFFileName = feFileName.ToLower().Replace(".fdf", ".pdf").Replace("_data", "")
        //    // LogMessage("feFileName.IndexOf(""]"") <= 0: feFileName: " & feFileName & ", fePDFFileName: " & fePDFFileName, Message, TraceEventType.Information)
        //    // Dim DateSt As String = Microsoft.VisualBasic.Right(fePDFFileName, 14)
        //    // fePDFFileName1 = fePDFFileName.Replace(DateSt, "[" & DateSt)
        //    // fePDFFileName1 = fePDFFileName1.Replace(".pdf", "].pdf")
        //    // LogMessage("fePDFFileNamet: " & fePDFFileName & ", fePDFFileName1: " & fePDFFileName1 & ", DateSt: " & DateSt, Message, TraceEventType.Information)
        //    // End If

        //}

        static private void AddNotFoundEmail(string email, bool CRLF = false, LogChannels channel = LogChannels.Access)
        {
            if (!System.IO.Directory.Exists(SYSTEM_LOG_PATH))
            {
                System.IO.Directory.CreateDirectory(SYSTEM_LOG_PATH);
            }

            System.Windows.Forms.Application.DoEvents();
            //this.Msg.Text = "Adding not found email: " + email;
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

        static private async Task ProcessStaffWorkRequests(GraphServiceClient graphClient)
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

            //var inboxMessages = await graphClient
            //       .Users[user.Id]
            //       .MailFolders.Inbox
            //       .Messages
            //       .Request()
            //       .Filter("isRead eq false")
            //       .Top(1000)
            //       //.Select("Body, Subject")
            //       .GetAsync();

            var inboxMessages = await graphClient
       .Users[user.Id]
       .MailFolders.Inbox
       .Messages
       .Request()
       .Filter("isRead eq false")
       .Top(1000)
       .GetAsync();

            //var y = inboxMessages.ElementAt(0);

            //if (y != null)
            //{
            //    ProcessStaffWorkRequestsMessageAsync(y, user);
            //}


            //return;

            foreach (Microsoft.Graph.Message x in inboxMessages)
            {
                //if (x.IsRead == false)
                //if ((x.Subject != null) && (!x.Subject.Contains(MSG_NOT_SAVED)))
                //{
                ProcessStaffWorkRequestsMessageAsync(x, user);
                //}
            }
        }

        static private void ProcessStaffWorkRequestsMessageAsync(Microsoft.Graph.Message EmailMsg, User user)
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

            string Body = EmailMsg.Body.Content;

            Body = Body.Replace("<br><br>", "\r\n");
            Body = Body.Replace("<BR><BR>", "\r\n");
            Body = Body.Replace("<br>", "");
            Body = Body.Replace("<BR>", "");
            Body = StripHtmlTags(Body);

            string[] Body1 = null;

            if (EmailMsg.Body.Content.Contains("\n\n"))
            {
                Body1 = Body.Split(new string[] { "\n\n" }, StringSplitOptions.RemoveEmptyEntries);
            }

            if (EmailMsg.Body.Content.Contains("\r\n"))
            {
                Body1 = Body.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            }

            //if ((Body1 != null) && (Body1.Count() == 7))
            if ((Body1 != null) && (EmailMsg.Body.Content.Contains("You have received a new Request to work an upcoming event")))
            {
                int StaffID = -1;
                int EventID = -1;
                int EventStaffRequiredID = -1;
                bool AutoBook = false;

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

                    if (y.Contains("Auto Book:"))
                    {
                        var ESRID = Strings.LTrim(Strings.RTrim(Strings.Mid(y, 12, y.Length - 11)));

                        if (ESRID == "YES")
                        {
                            AutoBook = true;
                        }
                        else
                        {
                            AutoBook = false;
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

                        string St1 = "SELECT Events.[Event ID] " + "FROM Events INNER JOIN [Staff Bookings] ON Events.[Event ID] = [Staff Bookings].[Event ID] " + "WHERE (Events.[Event ID]=" + EventID.ToString() + ") AND ([Staff Bookings].[Staff ID]=" + StaffID.ToString() + ") AND ((Events.Status=\"Current\") OR (Events.Status=\"Provisional\")) AND ([Staff Bookings].[Staff Total]>0) AND (Events.Date<=Date())";
                        var EID = ExecuteScalar(St1, EmailMsg);
                        if (EID == null)
                        {
                            EID = -1;
                        }

                        LogMessage("EID " + EID.ToString(), EmailMsg, TraceEventType.Information);

                        var St2 = "INSERT INTO StaffWorkInterest ( EventID, StaffID, EventStaffRequiredID, P, W, Sort ) ";
                        St2 = St2 + "SELECT " + EventID.ToString() + ", " + StaffID.ToString() + ", " + EventStaffRequiredID.ToString() + ", \"" + Microsoft.VisualBasic.Interaction.IIf(Conversions.ToInteger(SIDP) != -1, "P", "") + "\", \"" + Microsoft.VisualBasic.Interaction.IIf(Conversions.ToInteger(EID) != 0, "W", "") + "\", " + Microsoft.VisualBasic.Interaction.IIf(Conversions.ToInteger(SIDP) != 0, "1", Microsoft.VisualBasic.Interaction.IIf(Conversions.ToInteger(EID) != 0, "2", "3"));

                        //St2 = St2 + "WHERE NOT EXISTS (SELECT * FROM StaffWorkInterest WHERE (EventID = " + EventID.ToString() + ") AND (StaffID = " + StaffID.ToString() + "))";

                        var zzz = DLookupAccess("ID", "StaffWorkInterest", "(EventID = " + EventID.ToString() + ") AND (StaffID = " + StaffID.ToString() + ")");

                        if (zzz == null)
                        {
                            //Insert into StaffWorkInterest table
                            if (Conversions.ToInteger(ExecuteNonQuery(St2, EmailMsg)) >= 1)
                            {
                                LogMessage("Staff " + StaffID.ToString() + ", Event " + EventID.ToString() + ", request to work saved.", EmailMsg, TraceEventType.Information);

                                St2 = "UPDATE StaffWorkInterest INNER JOIN Staff ON StaffWorkInterest.StaffID = Staff.[Staff ID] SET StaffWorkInterest.Forename1 = Staff.[Forename], StaffWorkInterest.Surname1 = Staff.[Surname], StaffWorkInterest.DOB = Staff.[DOB], StaffWorkInterest.Mobile = Staff.[Mobile], StaffWorkInterest.Email = Staff.[E-Mail] ";

                                if (Conversions.ToInteger(ExecuteNonQuery(St2, EmailMsg)) < 1)
                                {
                                    LogError("Staff " + StaffID.ToString() + ", Event " + EventID.ToString() + ", UPDATE StaffWorkInterest INNER JOIN Staff failed.", EmailMsg, TraceEventType.Information);
                                }

                                St2 = "UPDATE StaffWorkInterest INNER JOIN [Event Staff Required] ON StaffWorkInterest.EventStaffRequiredID = [Event Staff Required].ID SET StaffWorkInterest.Role = [Event Staff Required].[Job Type], StaffWorkInterest.Start = [Event Staff Required].[Start], StaffWorkInterest.[End] = [Event Staff Required].[End];";

                                if (Conversions.ToInteger(ExecuteNonQuery(St2, EmailMsg)) < 1)
                                {
                                    LogError("Staff " + StaffID.ToString() + ", Event " + EventID.ToString() + ", UPDATE StaffWorkInterest INNER JOIN [Event Staff Required] failed.", EmailMsg, TraceEventType.Information);
                                }

                                St2 = "UPDATE StaffWorkInterest INNER JOIN (Clients INNER JOIN Events ON Clients.ID = Events.[Client id]) ON StaffWorkInterest.EventID = Events.[Event ID] SET StaffWorkInterest.EventDate = [Events].[Date], StaffWorkInterest.Venue = [Events].[Event Venue], StaffWorkInterest.Client = [Clients].[Client];";

                                if (Conversions.ToInteger(ExecuteNonQuery(St2, EmailMsg)) < 1)
                                {
                                    LogError("Staff " + StaffID.ToString() + ", Event " + EventID.ToString() + ", UPDATE StaffWorkInterest INNER JOIN (Clients INNER JOIN Events ON Clients.ID = Events.[Client id]) ON StaffWorkInterest.EventID = Events.[Event ID] failed.", EmailMsg, TraceEventType.Information);
                                }

                                string FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") + " Staff " + StaffID.ToString() + " work request data for event " + EventID.ToString() + " saved.txt");

                                try
                                {
                                    TextWriter tw = System.IO.File.CreateText(FullPath);
                                    tw.Close();
                                }
                                catch (Exception ex)
                                {
                                    LogError("ProcessStaffWorkRequestsMessageAsync::Staff " + StaffID.ToString() + ", Event " + EventID.ToString() + ", exception: " + ex.Message, EmailMsg, TraceEventType.Information);
                                }
                            }
                            else
                            {
                                LogError("Staff " + StaffID.ToString() + ", Event " + EventID.ToString() + ", request to work save failed.", EmailMsg, TraceEventType.Information);
                            }
                        }

                        string St3 = "UPDATE Events SET Events.JobRequests = IIf(IsNull([JobRequests]),1,[JobRequests]+1) " + "WHERE Events.[Event ID] = " + EventID.ToString();
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectGreaterEqual(ExecuteNonQuery(St3, EmailMsg), 1, false)))
                        {
                            LogMessage("Staff " + StaffID.ToString() + ", Event " + EventID.ToString() + ", event staff job request counter incremented.", EmailMsg, TraceEventType.Information);
                        }

                        LogMessage("EventStaffRequiredID " + EventStaffRequiredID.ToString(), EmailMsg, TraceEventType.Information);

                        if (AutoBook)
                        {
                            object JobType;
                            object StartTime;
                            object EndTime;
                            object Rate;
                            object Min;
                            object Hours;

                            object StaffRateID;
                            object StaffRateAmount;
                            object StaffRateMin;

                            if ((EventStaffRequiredID > 0) && (EventID > 0) && (StaffID > 0))
                            {
                                // EventStaffRequired job type and timings
                                JobType = DLookupAccess("[Job Type]", "[Event Staff Required]", "ID = " + EventStaffRequiredID.ToString());
                                StartTime = DLookupAccess("[Start]", "[Event Staff Required]", "ID = " + EventStaffRequiredID.ToString());
                                EndTime = DLookupAccess("[End]", "[Event Staff Required]", "ID = " + EventStaffRequiredID.ToString());

                                Rate = DLookupAccess("[Rate]", "[Event Staff Required]", "ID = " + EventStaffRequiredID.ToString());
                                Min = DLookupAccess("[Min]", "[Event Staff Required]", "ID = " + EventStaffRequiredID.ToString());
                                Hours = DLookupAccess("[Hours]", "[Event Staff Required]", "ID = " + EventStaffRequiredID.ToString());

                                if (JobType == null)
                                {
                                    JobType = "";
                                }

                                if (StartTime == null)
                                {
                                    StartTime = DateTime.MinValue;
                                }

                                if (EndTime == null)
                                {
                                    EndTime = DateTime.MinValue;
                                }

                                if (Rate == null)
                                {
                                    Rate = 0;
                                }

                                if (Min == null)
                                {
                                    Min = 0;
                                }

                                if (Hours == null)
                                {
                                    Hours = 0;
                                }

                                StaffRateID = DLookupAccess("[Rate ID]", "[Staff]", "[Staff ID] = " + StaffID.ToString());
                                StaffRateMin = DLookupAccess("minrate]", "[Staff]", "[Staff ID] = " + StaffID.ToString());
                                StaffRateAmount = DLookupAccess("[Normalrate]", "[Staff]", "[Staff ID] = " + StaffID.ToString());

                                if (StaffRateID == null)
                                {
                                    StaffRateID = 0;
                                }

                                if (StaffRateMin == null)
                                {
                                    StaffRateMin = 0;
                                }

                                if (StaffRateAmount == null)
                                {
                                    StaffRateAmount = 0;
                                }

                                var St4 = "INSERT INTO [Staff Bookings] ( [Event ID], [Staff ID], Start, actend, Role, [Rate ID], Rate, [Min], acthours, minclient, clientrate, [min hours], [Staff Total], Total, [Staff Travel Percentage], SelfBooked, ModifiedBy) ";
                                St4 = St4 + "SELECT " + EventID.ToString() + ", " + StaffID.ToString() + ", #" + Convert.ToDateTime(StartTime).ToString("HH:mm") + "#, #" + Convert.ToDateTime(EndTime).ToString("HH:mm") + "#, \"" + JobType + "\", \"" + StaffRateID + "\", " + Convert.ToDouble(StaffRateAmount).ToString() + " , " + Convert.ToDouble(StaffRateMin).ToString() + ", " + Convert.ToDouble(Hours).ToString() + ", " + Min.ToString() + ", " + Rate.ToString() + ", 4, 0, 0, 70, Now(), \"Self Booked\" ";

                                //St4 = St4 + "WHERE NOT EXISTS (SELECT * FROM [Staff Bookings] WHERE ([Event ID] = " + EventID.ToString() + ") AND ([Staff ID] = " + StaffID.ToString() + "))";

                                var zzzz = DLookupAccess("ID", "[Staff Bookings]", "([Event ID] = " + EventID.ToString() + ") AND ([Staff ID] = " + StaffID.ToString() + ")");

                                if (zzzz == null)
                                {
                                    if (Conversions.ToInteger(ExecuteNonQuery(St4, EmailMsg)) >= 1)
                                    {
                                        LogMessage("Staff " + StaffID.ToString() + " booked on event " + EventID.ToString() + " for job '" + JobType + "'", EmailMsg, TraceEventType.Information);
                                    }
                                    else
                                    {
                                        LogError("Error occured executing SQL: " + St4 + ": " + EventStaffRequiredID.ToString() + " EventID: " + EventID.ToString() + " StaffID: " + StaffID.ToString(), EmailMsg, TraceEventType.Information);
                                    }
                                }
                            }
                            else
                            {
                                LogError("One or more of required values missing, can not auto book: EventStaffRequiredID: " + EventStaffRequiredID.ToString() + " EventID: " + EventID.ToString() + " StaffID: " + StaffID.ToString(), EmailMsg, TraceEventType.Information);
                            }
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

        static private async Task ProcessStudentAvailability(GraphServiceClient graphClient)
        {

            var Users = await graphClient
                         .Users
                         .Request()
                         .Filter("startswith(mail,'WorkInterest')")
                         .GetAsync();

            var user = Users[0];

            var inboxMessages = await graphClient
               .Users[user.Id]
               .MailFolders.Inbox
               .Messages
               .Request()
               .Filter("isRead eq false")
               .Top(1000)
               .GetAsync();

            foreach (Microsoft.Graph.Message x in inboxMessages)
            {
                //if (x.IsRead == false)
                //if ((x.Subject != null) && (!x.Subject.Contains(MSG_NOT_SAVED)))
                //{
                ProcessStudentAvailabilityMessageAsync(x, user);
                //}
            }
        }

        static private void ProcessStudentAvailabilityMessageAsync(Microsoft.Graph.Message EmailMsg, User user)
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

            string Body = EmailMsg.Body.Content;

            Body = Body.Replace("<br><br>", "\r\n");
            Body = Body.Replace("<BR><BR>", "\r\n");
            Body = Body.Replace("<br>", "");
            Body = Body.Replace("<BR>", "");
            Body = StripHtmlTags(Body);

            //string BodyText = Strings.Replace("Staff ID: " + StaffID.ToString() + Constants.vbCrLf + Constants.vbCrLf + "Event ID: " + EventID.ToString() + Constants.vbCrLf + Constants.vbCrLf + "=====================================" + Constants.vbCrLf + Constants.vbCrLf + EmailMsg.Body.Content, Constants.vbCrLf, "<BR>");

            // Initialize variables with default values
            string Username = string.Empty;
            string email = string.Empty;
            bool examsInWinter = false;
            string winterExamDetails = string.Empty;
            bool examsInSummer = false;
            string summerExamDetails = string.Empty;
            DateTime startOfHigherEducation = DateTime.MinValue;
            DateTime endOfHigherEducation = DateTime.MinValue;
            string educationStartEndDetails = string.Empty;

            // Split the input string by new lines and process each line
            string[] lines = Body.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                // Split each line by the first occurrence of ": " to separate key and value
                var parts = line.Split(new[] { ": " }, 2, StringSplitOptions.None);
                if (parts.Length < 2)
                {
                    continue; // Skip lines that don't have key-value pairs
                }

                var key = parts[0].Trim();
                var value = parts[1].Trim();

                switch (key)
                {
                    case "Availability Details From":
                        Username = value;
                        break;

                    case "Email":
                        email = value;
                        break;

                    case "ExamsInWinter":
                        examsInWinter = bool.TryParse(value, out bool winterBool) ? winterBool : false;
                        break;

                    case "WinterExamDetails":
                        winterExamDetails = value;
                        break;

                    case "ExamsInSummer":
                        examsInSummer = bool.TryParse(value, out bool summerBool) ? summerBool : false;
                        break;

                    case "SummerExamDetails":
                        summerExamDetails = value;
                        break;

                    case "StartOfHigherEducation":
                        startOfHigherEducation = DateTime.TryParseExact(value, "M/d/yyyy h:mm:ss tt",
                                                    CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate) ? startDate : DateTime.MinValue;
                        break;

                    case "EndOfHigherEducation":
                        endOfHigherEducation = DateTime.TryParseExact(value, "M/d/yyyy h:mm:ss tt",
                                                CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate) ? endDate : DateTime.MinValue;
                        break;

                    case "EducationStartEndDetails":
                        educationStartEndDetails = value;
                        break;
                }
            }

            // Prepare the SQL update query
            string updateQuery = $@"
            UPDATE Staff 
            SET 
                Staff.DecemberExams = {(examsInWinter ? -1 : 0)}, 
                Staff.JuneExams = {(examsInSummer ? -1 : 0)}, 
                Staff.DecemberExamsText = '{winterExamDetails}', 
                Staff.JuneExamsText = '{summerExamDetails}', 
                Staff.HigherEducationStart = #{startOfHigherEducation.ToString("MM/dd/yyyy HH:mm:ss")}#, 
                Staff.HigherEducationEnd = #{endOfHigherEducation.ToString("MM/dd/yyyy HH:mm:ss")}#,
                Staff.HigherEducationText = '{educationStartEndDetails}' 
            WHERE 
                Staff.Username = '{Username}';
        ";

            // Output the generated query
            //Console.WriteLine(updateQuery);

            ExecuteNonQuery(updateQuery);

            var subject = EmailMsg.Subject.Replace(MSG_SAVED, "").Replace(MSG_NOT_SAVED, "");

            subject = MSG_SAVED + subject;

            //EmailMsg.Body.Content = BodyText;

            Microsoft.Graph.ItemBody b = new Microsoft.Graph.ItemBody();

            b.Content = EmailMsg.Body.Content;

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

        static private async Task ProcessTimesheetAttachmentsAsync()
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

        static private async Task ProcessStaffTimesheetsAsync(Microsoft.Graph.Message message, User user)
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

        static private void SaveAttachment(FileAttachment fileAttachment, string p)
        {
            //using (FileStream outputFileStream = new FileStream(p, FileMode.Create))
            //{
            //    fileAttachment.ContentBytes.CopyTo(outputFileStream);

            //    fileAttachment.ContentBytes.CopyTo()
            //}

            //Buffer fileContent = new Buffer(fileAttachment.ContentBytes);

            System.IO.File.WriteAllBytes(p, fileAttachment.ContentBytes);
        }

        static private void SetMessageSubject(Microsoft.Graph.Message message, User user, string v)
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

        static public async Task<bool> ProcessAsync(bool clearDiskSpace = false, Email_Processor.frmEmailProcessor frmEmailProcessor = null)
        {
            if (InProcess)
            {
                //LogMessage("InProcess false", null, TraceEventType.Information);
                return false;
            }

            //if (frmEmailProcessor != null)
            //{
            //    frmEmailProcessor.Timer1.Enabled = false;
            //}

            InProcess = true;

            CurrentTimeStamp = DateTime.Now;

            //ProcessTimesheetAttachments();
            //ProcessStaffWorkRequests();

            if (clearDiskSpace)
            {
                //LogMessage("ClearDiskSpace();", null, TraceEventType.Information);
                ClearDiskSpace();
            }

            DateTime LastRun = DateTime.Now;

            //LogMessage("if (isProcessUnsubscribeEmailsOK())", null, TraceEventType.Information);

            if (isProcessUnsubscribeEmailsOK())
            {
                //LogMessage("var config = LoadAppSettings();", null, TraceEventType.Information);

                var config = LoadAppSettings();
                if (null == config)
                {
                    LogError("var config = LoadAppSettings();", null, TraceEventType.Information, "Missing or invalid appsettings.json file. Please see README.md for configuration instructions.");
                    return false;
                }

                //LogMessage("graphClient = GetAuthenticatedGraphClient(config);", null, TraceEventType.Information);

                graphClient = GetAuthenticatedGraphClient(config);

                //LogMessage("return from graphClient = GetAuthenticatedGraphClient(config);", null, TraceEventType.Information);
            }

            // Test Code Begins

            //        var call = new Call
            //        {
            //            Targets = new List<InvitationParticipantInfo>
            //{
            //    new InvitationParticipantInfo
            //    {
            //        Identity = new IdentitySet
            //        {
            //            User = new Identity
            //            {
            //                Id = "sip:07930861921",
            //                DisplayName = "Yahya Usman"
            //            }
            //        }
            //    }
            //},
            //            MediaConfig = new ServiceHostedMediaConfig(),
            //            //From = new IdentitySet
            //            //{
            //            //    User = new Identity
            //            //    {
            //            //        Id = "sip:jane@contoso.com",
            //            //        DisplayName = "Jane Doe"
            //            //    }
            //            //}
            //        };

            //        await graphClient.Communications.Calls
            //            .Request()
            //            .AddAsync(call);


            // Test code ends

            //try
            //{
            //    LogMessage("Begin recruitment staff processing...", null, TraceEventType.Information, "", true);

            //    await ProcessRecruitmentEmailsAsync(graphClient);
            //}
            //catch (Exception ex)
            //{
            //    LogError("Process::await ProcessRecruitmentEmailsAsync(graphClient);", null, TraceEventType.Error, ex.Message);
            //}


            //LogMessage("", null, TraceEventType.Information, "", true);

            try
            {
                LogMessage("Begin unsubscribe staff processing...", null, TraceEventType.Information, "", true);

                await ProcessUnsubscribeEmailsAsync(graphClient);
            }
            catch (Exception ex)
            {
                LogError("Process::await ProcessUnsubscribeEmailsAsync(graphClient);", null, TraceEventType.Error, ex.Message);
            }

            //LogMessage("End unsubscribe staff processing...", null, TraceEventType.Information, "", true);
            //LogMessage(" ", null, TraceEventType.Information, "", true);

            try
            {
                LogMessage("Begin processing staff work requests...", null, TraceEventType.Information, "", true);

                await ProcessStaffWorkRequests(graphClient);
            }
            catch (Exception ex)
            {
                LogError("Process::await ProcessStaffWorkRequests(graphClient);", null, TraceEventType.Error, ex.Message);
            }

            try
            {
                LogMessage("Begin processing student availability...", null, TraceEventType.Information, "", true);

                await ProcessStudentAvailability(graphClient);
            }
            catch (Exception ex)
            {
                LogError("Process::await ProcessStudentAvailability(graphClient);", null, TraceEventType.Error, ex.Message);
            }

            try
            {
                LogMessage("Begin processing timesheet attachments...", null, TraceEventType.Information, "", true);

                await ProcessTimesheetAttachmentsAsync();
            }
            catch (Exception ex)
            {
                LogError("Process::await ProcessTimesheetAttachmentsAsync();", null, TraceEventType.Error, ex.Message);
            }

            //Properties.Settings.Default.UnsubscribeServiceLastRun = DateTime.Now.ToOADate();
            Email_Processor_Framework.Properties.Settings.Default.UnsubscribeServiceLastRun = LastRun.ToOADate();

            double LastRun1 = Email_Processor_Framework.Properties.Settings.Default.UnsubscribeServiceLastRun;

            Email_Processor_Framework.Properties.Settings.Default.Save();
            Email_Processor_Framework.Properties.Settings.Default.Upgrade();
            //System.Windows.Forms.Application.Restart();

            graphClient = null;
            //}

            InProcess = false;

            return true;
        }

        static private bool isProcessUnsubscribeEmailsOK()
        {
            return true;

            //DateTime UnsubscribeServiceLastRun = DateTime.FromOADate(Properties.Settings.Default.UnsubscribeServiceLastRun);

            ////int UnsubscribeServiceInterval = (int)Math.Round(60000f * Properties.Settings.Default.UnsubscribeServiceTimeinMinutes);

            //if (DateTime.Now >= UnsubscribeServiceLastRun.AddMinutes(Properties.Settings.Default.UnsubscribeServiceTimeinMinutes))
            //    return true;
            //else
            //    return false;
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
            try
            {
                var authenticationProvider = CreateAuthorizationProvider(config);
                _graphServiceClient = new GraphServiceClient(authenticationProvider);
                return _graphServiceClient;
            }
            catch (Exception ex)
            {
                LogError("frmEmailProcessor::GetAuthenticatedGraphClient", null, TraceEventType.Error, ex.Message);
            }

            return null;
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
            //LogMessage("In LoadAppSettings()", null, TraceEventType.Information);

            try
            {
                //LogMessage("var config = new ConfigurationBuilder()", null, TraceEventType.Information);

                var config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", false, true)
                .Build();

                //.SetBasePath(System.IO.Directory.GetCurrentDirectory())

                //LogMessage("if (string.IsNullOrEmpty(config[\"applicationId\"]) ||", null, TraceEventType.Information);

                // Validate required settings
                if (string.IsNullOrEmpty(config["applicationId"]) ||
                    string.IsNullOrEmpty(config["applicationSecret"]) ||
                    string.IsNullOrEmpty(config["redirectUri"]) ||
                    string.IsNullOrEmpty(config["tenantId"]) ||
                    string.IsNullOrEmpty(config["domain"]))
                {
                    return null;
                }

                //LogMessage("Exiting LoadAppSettings()", null, TraceEventType.Information);

                return config;
            }
            catch (System.IO.FileNotFoundException ex1)
            {
                LogError("catch (System.IO.FileNotFoundException ex1)", null, TraceEventType.Information, ex1.Message);
                return null;
            }
            catch (Exception ex2)
            {
                LogError("catch (Exception ex2)", null, TraceEventType.Information, ex2.Message);
                return null;
            }
        }

        static object ExecuteScalar(string St, Microsoft.Graph.Message message = null)
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

        static object ExecuteNonQuery(string St, Microsoft.Graph.Message message = null)
        {
            var Conn = new OleDbConnection(LocalAccessConnSt);
            try
            {
                //LogMessage("Opening ExecuteNonQuery Access db connection.", message, TraceEventType.Information);
                Conn.Open();
                //LogMessage("ExecuteNonQuery Access db connection opened.", message, TraceEventType.Information);
                var Command = new OleDbCommand(St, Conn);

                LogMessage("ExecuteNonQuery::Executing SQL: " + St, message, TraceEventType.Information);

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

        static object ExecuteNonQuerySQLServer(string St, Microsoft.Graph.Message message = null)
        {
            var Conn = new SqlConnection(LocalSQLServerConnSt);
            try
            {
                //LogMessage("Opening ExecuteNonQuery SQL Server db connection.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                Conn.Open();
                //LogMessage("ExecuteNonQuery SQL Server db connection opened.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
                var Command = new SqlCommand(St, Conn);

                LogMessage("ExecuteNonQuerySQLServer::Executing SQL: " + St, message, TraceEventType.Information);

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

        static object DLookupAccess(string Value, string Table, string Condition = "", Microsoft.Graph.Message message = null)
        {
            var Conn = new OleDbConnection(LocalAccessConnSt);

            string St = "SELECT [" + Value + "] FROM [" + Table + "] " + (string.IsNullOrEmpty(Condition) ? "" : "WHERE " + Condition);

            St = St.Replace("[[", "[");
            St = St.Replace("]]", "]");

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

        static object DLookupSQLServer(string Value, string Table, string Condition = "", Microsoft.Graph.Message message = null)
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

        static object DLookupSQLServerRemote(string Value, string Table, string Condition = "", Microsoft.Graph.Message message = null)
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
                LogError("Error dlookup SQL Server Remote db. St: " + St + ", Exception: " + ex.Message, message, TraceEventType.Error);
                return null;
            }
            finally
            {
                Conn.Close();
            }
        }

        static public string StripHtmlTags(string html)
        {

            // Remove HTML tags.
            return Regex.Replace(html, "<.*?>", "");
        }

        static void ClearDiskSpace()
        {
            string networkDrive = @"F:\";
            //string backupFolder = @"F:\Events Data\DB_Backups";
            string backupFolder = @"Z:\OneDrive - High Society\DB_Backups - MS Access";
            //string backupFolder = @"H:\DB_Backups";

            long requiredSpace = (long)(5.0 * 1024.0 * 1024.0 * 1024.0); // 5GB in bytes

            //remove when done
            //return;


            //zip all mdb files
            ZipAndTestMdbFiles(backupFolder);


            DriveInfo drive = new DriveInfo(networkDrive);
            long freeSpace = drive.AvailableFreeSpace;

            if (freeSpace < requiredSpace)
            {
                string backupFolderPath = Path.Combine(networkDrive, backupFolder);

                //string[] mdbFiles = System.IO.Directory.EnumerateFiles(backupFolderPath, "*.mdb")
                //                    .Union(System.IO.Directory.EnumerateFiles(backupFolderPath, "*.zip"))
                //                    .ToArray();

                string[] mdbFiles = System.IO.Directory.EnumerateFiles(backupFolderPath, "*.mdb")
                        .ToArray();

                // Get all *.mdb files in the backup folder
                //string[] mdbFiles = System.IO.Directory.GetFiles(backupFolderPath, "*.mdb");

                //// Sort the files by creation time (oldest first)
                //var sortedFiles = mdbFiles.Select(file => new FileInfo(file))
                //                         .OrderBy(file => file.CreationTime)
                //                         .ToList();

                // Calculate the date one month ago
                DateTime oneMonthAgo = DateTime.Now.AddDays(-7);

                // Sort the files by creation time (oldest first) and filter files that are at least a month old
                var sortedFiles = mdbFiles.Select(file => new FileInfo(file))
                                          .Where(file => file.CreationTime <= oneMonthAgo)
                                          .OrderBy(file => file.CreationTime)
                                          .ToList();

                long spaceToFree = requiredSpace - freeSpace;

                foreach (var file in sortedFiles)
                {
                    if (spaceToFree <= 0)
                        break;

                    // Delete the file
                    System.IO.File.Delete(file.FullName);
                    spaceToFree -= file.Length;
                }

                //Console.WriteLine("Files deleted successfully.");
            }
            //else
            //{
            //    Console.WriteLine("Disk space is sufficient.");
            //}

            //Console.ReadLine();
        }

        static void ZipAndTestMdbFiles(string folderPath)
        {
            string[] mdbFiles = System.IO.Directory.GetFiles(folderPath, "*.mdb");

            foreach (string mdbFile in mdbFiles)
            {
                string zipFilePath = Path.ChangeExtension(mdbFile, ".zip");

                if (!System.IO.File.Exists(zipFilePath))
                {
                    if (CreateZipFile(mdbFile, zipFilePath))
                    {
                        if (TestZipFile(zipFilePath))
                        {
                            Console.WriteLine($"Successfully zipped and tested: {zipFilePath}");
                        }
                        else
                        {
                            Console.WriteLine($"Failed to test: {zipFilePath}");
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"Zip file already exists: {zipFilePath}");
                }
            }
        }

        static bool CreateZipFile(string sourceFile, string zipFilePath)
        {
            using (FileStream zipToOpen = new FileStream(zipFilePath, FileMode.Create))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    try
                    {
                        archive.CreateEntryFromFile(sourceFile, Path.GetFileName(sourceFile));
                    }
                    catch
                    {
                        return false;
                    }

                    return true;
                }
            }
        }

        static bool TestZipFile(string zipFilePath)
        {
            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
                {
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        using (Stream stream = entry.Open())
                        {
                            byte[] buffer = new byte[1024];
                            while (stream.Read(buffer, 0, buffer.Length) > 0) { }
                        }
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    // This class encapsulates the details of getting a token from MSAL and exposes it via the 
    // IAuthenticationProvider interface so that GraphServiceClient or AuthHandler can use it.
    // A significantly enhanced version of this class will in the future be available from
    // the GraphSDK team.  It will supports all the types of Client Application as defined by MSAL.
    public class MsalAuthenticationProvider : IAuthenticationProvider
    {
        private IConfidentialClientApplication _clientApplication;
        private string[] _scopes;

        public MsalAuthenticationProvider(IConfidentialClientApplication clientApplication, string[] scopes)
        {
            _clientApplication = clientApplication;
            _scopes = scopes;
        }

        /// <summary>
        /// Update HttpRequestMessage with credentials
        /// </summary>
        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            var token = await GetTokenAsync();
            request.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
        }

        /// <summary>
        /// Acquire Token 
        /// </summary>
        public async Task<string> GetTokenAsync()
        {
            AuthenticationResult authResult = null;
            authResult = await _clientApplication.AcquireTokenForClient(_scopes)
                                .ExecuteAsync();
            return authResult.AccessToken;
        }
    }
}

//Fill code here to insert data from message body into Applicants table with below structure
//CREATE TABLE [dbo].[Applicants] (

//    [ApplicantID][int] IDENTITY(1, 1) NOT NULL,

//    [Username] [nvarchar] (255) NULL,

//    [Title] [nvarchar] (50) NULL,

//    [Forenames] [nvarchar] (255) NULL,

//    [Surname] [nvarchar] (255) NULL,

//    [DOB] [datetime] NULL,
//	[Tel][nvarchar] (255) NULL,

//    [Mobile] [nvarchar] (255) NULL,

//    [Email] [nvarchar] (255) NULL,

//    [Address1] [nvarchar] (255) NULL,

//    [Address2] [nvarchar] (255) NULL,

//    [Address3] [nvarchar] (255) NULL,

//    [Town] [nvarchar] (255) NULL,

//    [County] [nvarchar] (255) NULL,

//    [Country] [nvarchar] (255) NULL,

//    [Postcode] [nvarchar] (255) NULL,

//    [JobType] [nvarchar] (max)NULL,
//	[CVFileName][nvarchar] (max)NULL,
//	[CVFileType][nvarchar] (255) NULL,

//    [CVFileContentType] [nvarchar] (255) NULL,

//    [CVFileSize] [int] NULL,
//	[CVFileContent][varbinary] (max)NULL,
//	[PhotoFileName][nvarchar] (max)NULL,
//	[PhotoFileType][nvarchar] (255) NULL,

//    [PhotoFileContentType] [nvarchar] (255) NULL,

//    [PhotoFileSize] [int] NULL,
//	[PhotoFileContent][varbinary] (max)NULL,
//	[Reregister][bit] NULL,
//	[Stage][nvarchar] (50) NULL,

//    [Download] [bit] NULL,
//	[DownloadDate][datetime] NULL,
//	[ApplyDate][datetime] NULL,
//	[Source][nvarchar] (max)NULL,
//	[ProfessionalStatus][nvarchar] (50) NULL,

//    [PreviousExperience] [nvarchar] (1024) NULL,

//    [RefereeName1] [nvarchar] (35) NULL,

//    [RefereePhone1] [nvarchar] (35) NULL,

//    [RefereeEmail1] [nvarchar] (35) NULL,

//    [RefereeName2] [nvarchar] (35) NULL,

//    [RefereePhone2] [nvarchar] (35) NULL,

//    [RefereeEmail2] [nvarchar] (35) NULL,

//    [CateringManagerExperience] [bit] NULL,
//	[UKDrivingLicense][bit] NULL,
// CONSTRAINT[PK_Applicants] PRIMARY KEY CLUSTERED 
//(

//    [ApplicantID] ASC
//)WITH(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON[PRIMARY]
//) ON[PRIMARY] TEXTIMAGE_ON[PRIMARY]
//GO
