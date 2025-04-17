using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Sentry;
using System.Windows.Forms;
using System.Runtime.ExceptionServices;
using System.Diagnostics;
using static Email_Processor.Utils;
using Email_Processor_Framework;
using System.Linq;

namespace Email_Processor
{
    static class Program
    {
        private static System.Windows.Forms.Timer Timer1 = new System.Windows.Forms.Timer();

        //private static HttpClient _httpClient;
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static async Task Main(string[] args)
        {
            // Init the Sentry SDK
            SentrySdk.Init(o =>
            {
                // Tells which project in Sentry to send events to:
                o.Dsn = "https://6255b28634000a79cdf27aee7a067437@o4505871099756544.ingest.sentry.io/4505962475552768";
                // When configuring for the first time, to see what the SDK is doing:
                o.Debug = true;
                // Set TracesSampleRate to 1.0 to capture 100% of transactions for performance monitoring.
                // We recommend adjusting this value in production.
                o.TracesSampleRate = 1.0;
            });
            // Configure WinForms to throw exceptions so Sentry can capture them.
            System.Windows.Forms.Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);

            //System.Windows.Forms.Application.SetHighDpiMode(System.Windows.Forms.HighDpiMode.SystemAware);
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);

            CurrentTimeStamp = DateTime.Now;

            //LogMessage("I am in!",null,TraceEventType.Information);

            // Check for the 'run' parameter in the command-line arguments.
            if (args.Contains("run"))
            {
                // Call EmailProcessor.ProcessAsync(null) and then exit the application.
                try
                {
                    //LogMessage("About to run EmailProcessor.ProcessAsync from command line", null, TraceEventType.Information);

                    await EmailProcessor.ProcessAsync();
                }
                catch (Exception ex)
                {
                    LogError("Main()::_ = EmailProcessor.ProcessAsync(null);", null, TraceEventType.Error, ex.Message);
                }
                return;
            }

            try
            {
                System.Windows.Forms.Application.Run(new frmEmailProcessor());

       
            }
            catch (Exception ex)
            {
                LogError("Main()::System.Windows.Forms.Application.Run(new frmEmailProcessor());", null, TraceEventType.Error, ex.Message);
            }
        }
    }
}
