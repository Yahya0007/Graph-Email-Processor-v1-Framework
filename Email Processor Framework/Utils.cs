using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Email_Processor
{
    public static class Utils
    {
        public const string TEST_TOKEN = "[TEST]";
        public const string SYSTEM_LOG_PATH = @"\\data-server\Logs\Graph Email Processor v1 Logs\";
        public const string SYSTEM_LOG_ACCESS = "GraphEmailProcessor-Access.log";
        public const string SYSTEM_LOG_SQLSERVER = "GraphEmailProcessor-SQLServer.log";
        public const string SYSTEM_LOG_ACCESS_NOTFOUND = "GraphEmailProcessor-Access-NotFound.log";
        public const string SYSTEM_LOG_SQLSERVER_NOTFOUND = "GraphEmailProcessor-SQLServer-NotFound.log";
        public const string REFERENCES_LOG = "ProcessAttachmentsv1-References.log";
        public const string UNSUBSCRIBE_LOG = "ProcessAttachmentsv1-Unsubscribe.log";
        public const string TIMESHEETS_DATA_PATH = @"\\hsserver\High Society - Documents\PDF Timesheets\Data\";
        public const string TIMESHEETS_PDF_PATH = @"\\hsserver\High Society - Documents\PDF Timesheets\";
        public const string REFERENCES_DATA_PATH = @"F:\References\Data\";
        public const string REFERENCES_PDF_PATH = @"F:\References\Sent\";
        public const string MSG_NOT_SAVED = "WARNING! Data file not saved, no matching pdf file: ";
        public const string MSG_SAVED = "SAVED: ";
        public const string MSG_NOT_AUTOBOOKED = "NOT AUTO BOOKED: ";
        public const string MSG_AUTOBOOKED = "AUTO BOOKED: ";
        //public const string RemoteSQLServerConnSt = "Data Source=91.232.125.193;Initial Catalog=HS_Staff_Portal;Persist Security Info=True;User ID=HS_Staff_Portal;Password=yMYSyZKY#9";
        public const string RemoteSQLServerConnSt = "Data Source=91.232.125.193;Initial Catalog=EMS_2018_HS;Persist Security Info=True;User ID=EMS_2018_HS;Password=P@ssword01;TrustServerCertificate=True;";
        public const string LocalSQLServerConnSt = @"Data Source=DB-SERVER;Initial Catalog=EMS_2018_HS;Persist Security Info=True;User ID=EMS_2018_HS;Password=P@ssword01;TrustServerCertificate=True;";
        public const string LocalAccessConnSt = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""\\hsserver\Events Data\Events Data\Events Data.mdb"";";

        public enum UnsubscribeStatus
        {
            Unsubscribed,
            NotFound,
            Error
        }

        public enum LogChannels
        {
            Access,
            SQLServer,
            Misc
        }

        public static DateTime CurrentTimeStamp;

        public static string LeftStr(string param, int length)
        {
            string result = param.Substring(0, length);
            return result;
        }

        public static string RightStr(string param, int length)
        {
            string result = param.Substring(param.Length - length, length);
            return result;
        }

        public static string Mid(string param, int startIndex, int length)
        {
            string result = param.Substring(startIndex, length);
            return result;
        }

        public static string Mid(string param, int startIndex)
        {
            string result = param.Substring(startIndex);
            return result;
        }

        public static void LogMessage(string eventName, Microsoft.Graph.Message item, TraceEventType e, string message = "", bool CRLF = false, LogChannels channel = LogChannels.Access, bool UI = false)
        {
            if (!System.IO.Directory.Exists(SYSTEM_LOG_PATH))
            {
                System.IO.Directory.CreateDirectory(SYSTEM_LOG_PATH);
            }

            if (UI)
            {
                System.Windows.Forms.Application.DoEvents();
                //this.Msg.Text = eventName;
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

        public static void LogError(string eventName, Microsoft.Graph.Message item, TraceEventType e, string Message = "", bool CRLF = false, LogChannels channel = LogChannels.Access, bool UI = false)
        {
            if (!System.IO.Directory.Exists(SYSTEM_LOG_PATH))
            {
                System.IO.Directory.CreateDirectory(SYSTEM_LOG_PATH);
            }

            if (UI)
            {
                System.Windows.Forms.Application.DoEvents();
                //this.Msg.Text = eventName;
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

        public static bool IsSpecificTime(TimeSpan now, TimeSpan targetTime)
        {
            // Check if the current time is exactly the specified time
            return now.Hours == targetTime.Hours && now.Minutes == targetTime.Minutes;
        }
    }
}
