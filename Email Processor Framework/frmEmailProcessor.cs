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
using Email_Processor_Framework;
//using System.IO.Compression.FileSystem;

namespace Email_Processor
{
    public partial class frmEmailProcessor : Form
    {
        public frmEmailProcessor()
        {
            InitializeComponent();

            // Load appsettings.json
            //var config = LoadAppSettings();
            //if (null == config)
            //{
            //    Console.WriteLine("Missing or invalid appsettings.json file. Please see README.md for configuration instructions.");
            //    return;
            //}

            //Query using Graph SDK (preferred when possible)
            //graphClient = GetAuthenticatedGraphClient(config);
        }

        private async void btnProcess_Click(object sender, EventArgs e)
        {
            //ClearDiskSpace();

            this.btnProcess.Enabled = false;
            System.Windows.Forms.Application.DoEvents();
            this.Timer1.Enabled = false;

            await EmailProcessor.ProcessAsync(true,this);

            this.Timer1.Enabled = true;
            this.btnProcess.Enabled = true;
            System.Windows.Forms.Application.DoEvents();
        }

        private void frmEmailProcessor_Load(object sender, EventArgs e)
        {
            this.Timer1.Interval = (int)Math.Round(60000f * Email_Processor_Framework.Properties.Settings.Default.TimerTimeinMinutes);
            //this.Timer1.Enabled = false;
            this.Timer1.Enabled = true;
        }

        //private async void simpleButton1_Click(object sender, EventArgs e)
        //{
        //    var config = LoadAppSettings();
        //    if (null == config)
        //    {
        //        Console.WriteLine("Missing or invalid appsettings.json file. Please see README.md for configuration instructions.");
        //        return;
        //    }

        //    graphClient = GetAuthenticatedGraphClient(config);

        //    //var graphClient = new GraphServiceClient(authProvider);

        //    var call = new Call
        //    {
        //        Targets = new List<InvitationParticipantInfo>
        //        {
        //            new InvitationParticipantInfo
        //            {
        //                Identity = new IdentitySet
        //                {
        //                    User = new Identity
        //                    {
        //                        Id = "sip:07930861921",
        //                        DisplayName = "Yahya Usman"
        //                    }
        //                }
        //            }
        //        },
        //        MediaConfig = new ServiceHostedMediaConfig(),
        //        //From = new IdentitySet
        //        //{
        //        //    User = new Identity
        //        //    {
        //        //        Id = "sip:jane@contoso.com",
        //        //        DisplayName = "Jane Doe"
        //        //    }
        //        //}
        //    };

        //    await graphClient.Communications.Calls
        //        .Request()
        //        .AddAsync(call);

        //}

        private async void Timer1_Tick_1(object sender, EventArgs e)
        {
            this.Timer1.Enabled = false;

            await EmailProcessor.ProcessAsync(true,this);

            Timer1.Interval = (int)Math.Round(60000f * Email_Processor_Framework.Properties.Settings.Default.TimerTimeinMinutes);

            this.Timer1.Enabled = true;
        }
    }
}
