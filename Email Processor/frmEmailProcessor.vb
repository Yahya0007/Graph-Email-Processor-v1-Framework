Public Partial Class frmEmailProcessor
    Inherits Form

    Friend Enum UnsubscribeStatus
        Unsubscribed
        NotFound
        [Error]
    End Enum

    Friend Enum LogChannels
        Access
        SQLServer
        Misc
    End Enum

    Private Shared _graphServiceClient As GraphServiceClient
    Private graphClient As GraphServiceClient = Nothing
    Private Const domain As String = "high-society.co.uk"
    Public Const TEST_TOKEN As String = "[TEST]"
    Public Const SYSTEM_LOG_PATH As String = "C:\Graph Email Processor v1 Logs\"
    Public Const SYSTEM_LOG_ACCESS As String = "GraphEmailProcessor-Access.log"
    Public Const SYSTEM_LOG_SQLSERVER As String = "GraphEmailProcessor-SQLServer.log"
    Public Const SYSTEM_LOG_ACCESS_NOTFOUND As String = "GraphEmailProcessor-Access-NotFound.log"
    Public Const SYSTEM_LOG_SQLSERVER_NOTFOUND As String = "GraphEmailProcessor-SQLServer-NotFound.log"
    Public Const REFERENCES_LOG As String = "ProcessAttachmentsv1-References.log"
    Public Const UNSUBSCRIBE_LOG As String = "ProcessAttachmentsv1-Unsubscribe.log"
    Public Const TIMESHEETS_DATA_PATH As String = "F:\PDF Timesheets\Data\"
    Public Const TIMESHEETS_PDF_PATH As String = "F:\PDF Timesheets\"
    Public Const REFERENCES_DATA_PATH As String = "F:\References\Data\"
    Public Const REFERENCES_PDF_PATH As String = "F:\References\Sent\"
    Public Const MSG_NOT_SAVED As String = "WARNING! Data file not saved, no matching pdf file: "
    Public Const MSG_SAVED As String = "SAVED: "
    Public Const MSG_NOT_AUTOBOOKED As String = "NOT AUTO BOOKED: "
    Public Const MSG_AUTOBOOKED As String = "AUTO BOOKED: "
    Public Const RemoteSQLServerConnSt As String = "Data Source=91.232.125.193;Initial Catalog=HS_Staff_Portal;Persist Security Info=True;User ID=HS_Staff_Portal;Password=yMYSyZKY#9"
    Public Const LocalSQLServerConnSt As String = "Data Source=DB-SERVER\SQL2K14;Initial Catalog=EMS_2018_HS;Persist Security Info=True;User ID=EMS_2018_HS;Password=P@ssword01"
    Public Const LocalAccessConnSt As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""S:\Events Data\Events Data.mdb"";"
    Public InProcess As Boolean = False

    Public Enum PDFType
        None = -1
        Timesheet = 0
        Reference = 1
    End Enum

    Public CurrentTimeStamp As Date

    Private Function ConvertToSecureString(ByVal password As String) As SecureString
        If password Is Nothing Then
            Throw New ArgumentNullException("password")
        End If

        Dim securePassword = New SecureString()

        For Each c In password
            securePassword.AppendChar(c)
        Next

        securePassword.MakeReadOnly()
        Return securePassword
    End Function

    Public Sub New()
        Me.InitializeComponent()

        ' Load appsettings.json
        Dim config = LoadAppSettings()

        If Nothing Is config Then
            Console.WriteLine("Missing or invalid appsettings.json file. Please see README.md for configuration instructions.")
            Return
        End If

        'Query using Graph SDK (preferred when possible)
        graphClient = GetAuthenticatedGraphClient(config)
    End Sub

    Private Sub btnProcess_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process()
    End Sub

#Region "ProcessUnsubscribeEmailsAsync"

    Private Async Function ProcessUnsubscribeEmailsAsync(ByVal graphClient As GraphServiceClient) As Task
        Dim options As List(Of QueryOption) = New List(Of QueryOption) From {
        'new QueryOption("$top", "1")
        }

        'var graphResult = await graphClient.Users.Request(options)
        '    .Filter("startswith(displayName,'Unsubscribe')")
        '    .GetAsync();

        'Console.WriteLine("Email Processor Result");
        'Console.WriteLine(graphResult[0].DisplayName);

        Dim Users = Await graphClient.Users.Request().Filter("startswith(displayName,'Unsubscribe')").Top(1).GetAsync()
        Dim user = Users(0)

        ' Get message from the user's inbox
        'var inboxMessages = await graphClient
        '            .Users[User.Id]
        '            .MailFolders.Inbox
        '            .Messages
        '            .Request()
        '            .Select(e => new
        '            {
        '                e.Subject,
        '                e.Body,
        '                e.Sender
        '            })
        '            .OrderBy("ReceivedDateTime/dateTime")
        '            .GetAsync();

        'var inboxMessages = await graphClient
        ' .Users[user.Id]
        ' .MailFolders.Inbox
        ' .Messages
        ' .Request()
        ' .Filter("startswith(Subject,'Unsub: ')")
        ' .OrderBy("receivedDateTime DESC")
        ' .GetAsync();

        Dim inboxMessages = Await graphClient.Users(user.Id).MailFolders.Inbox.Messages.Request().Filter("not startswith(Subject,'Unsub: ')").GetAsync()


        'var inboxMessages1 = await graphClient
        '    .Me
        '    .MailFolders.Inbox
        '    .Messages
        '    .Request()
        '    .OrderBy("")
        '    .GetAsync();


        For Each x In inboxMessages
            Await ProcessMessageAsync(x, user)
        Next
    End Function

    Private Async Function ProcessMessageAsync(ByVal message As Microsoft.Graph.Message, ByVal user As User) As Task
        'Check if message has further message attachments and process them
        If message.HasAttachments = True Then
            Dim attachments = Await graphClient.Users(user.Id).Messages($"{message.Id}").Attachments.Request().GetAsync()

            'foreach (var attachment in message.Attachments)
            For Each attachment In attachments

                If Equals(attachment.ODataType, "#microsoft.graph.itemAttachment") Then
                    Dim attachmentRequest = Await graphClient.Users(user.Id).MailFolders.Inbox.Messages(message.Id).Attachments(attachment.Id).Request().Expand("microsoft.graph.itemattachment/item").GetAsync()
                    Dim itemAttachment = CType(attachmentRequest, ItemAttachment)
                    Dim itemMessage = CType(itemAttachment.Item, Microsoft.Graph.Message)

                    If itemMessage IsNot Nothing Then
                        Await ProcessMessageAsync(itemMessage, user)
                    End If


                    'if (attachmentRequest.ODataType == "#microsoft.graph.message")
                    '{

                    '    attachmentRequest.GetType
                    '    var itemMessage = (Microsoft.Graph.Message)attachment;

                    '    if (itemMessage != null)
                    '    {
                    '        await ProcessMessageAsync(itemMessage, user);
                    '    }
                    '}

                    '                    var attachment1 = await graphClient.Me.Messages["{message-id}"].Attachments["{attachment-id}"]
                    '.Request()
                    '.GetAsync();
                    'var itemAttachment = (ItemAttachment)attachmentRequest.Result;


                End If
                'else
                '{
                '    var fileAttachment = (FileAttachment)attachment;
                '    System.IO.File.WriteAllBytes(System.IO.Path.Combine(downloadPath, fileAttachment.Name), fileAttachment.ContentBytes);
                '}
            Next
        End If

        'List<string> Emails = new List<string> { };



        If Not message.Sender.EmailAddress.Address.Contains(domain) Then
            'Emails.Add(x.Sender.EmailAddress.Address);
            UnsubcribeStaff(message, message, user)
        Else
            'x.ConversationId
            'x.ConversationIndex

            'var inboxMessages1 = graphClient
            '    .Users[User.Id]
            '    .Messages
            '    .Request()
            '    .Select(e => new
            '    {
            '        e.Subject,
            '        e.Body,
            '        e.Sender
            '    })
            '    .Filter($"startswith(ConversationId,{x.ConversationId})")
            '    .OrderBy("Recieved/dateTime")
            '    .GetAsync().Result;

            ' Get email to subscribe from body text
            Dim z = message.Body.ContentType
            message.Body.ContentType = BodyType.Text
            Dim EmailBody = message.Body.Content
            message.Body.ContentType = z
            Dim emailRegex As Regex = New Regex("\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*", RegexOptions.IgnoreCase)
            'find items that matches with our pattern
            Dim emailMatches = emailRegex.Matches(EmailBody)

            For Each emailMatch As Match In emailMatches

                If Not emailMatch.Value.Contains(domain) Then
                    'Emails.Add(emailMatch.Value);
                    Dim m = New Microsoft.Graph.Message()
                    Dim b = New ItemBody()
                    b.Content = EmailBody
                    m.Body = b
                    Dim e = New EmailAddress()
                    e.Address = emailMatch.Value
                    e.Name = emailMatch.Value
                    Dim s = New Recipient()
                    s.EmailAddress = e
                    m.Sender = s

                    'if (UnsubcribeStaff(emailMatch.Value, emailMatch.Value, EmailBody, null)== UnsubscribeStatus.Unsubscribed)
                    UnsubcribeStaff(m, message, user)
                End If
            Next
        End If

        'Finally update the message subject back into data store
        'if (x.Subject.Contains("Unsub: "))
        '{

        '    //List<QueryOption> options1 = new List<QueryOption>
        '    //{
        '    //    //new QueryOption("$top", "1")
        '    //};


        '    //await graphClient
        '    //    .Users[user.Id]
        '    //    .MailFolders.Inbox
        '    //    .Messages["{x.Id}"]
        '    //    .Request()
        '    //    .UpdateAsync(x);

        '    Update(user, x);


        '    //var inboxMessages = await graphClient
        '    // .Users[user.Id]
        '    // .MailFolders.Inbox
        '    // .Messages
        '    // .Request()
        '    // .OrderBy("receivedDateTime DESC")
        '    // .GetAsync();

        '}

    End Function

    'private void Update(User user, Microsoft.Graph.Message x)
    '{
    '    graphClient
    '    .Users[user.Id]
    '    .MailFolders.Inbox
    '    .Messages[$"{x.Id}"]
    '    .Request()
    '    .UpdateAsync(x);

    '    //Console.WriteLine("I am here!");

    '    LogMessage("Message from sender " + x.Sender.EmailAddress.Address + " with subject '" + x.Subject + "' updated successfully.", x, TraceEventType.Information);
    '}

    Private Sub UnsubcribeStaff(ByVal message As Microsoft.Graph.Message, ByVal originalmessage As Microsoft.Graph.Message, ByVal user As User)
        LogMessage("Processing unsubscribe request.", message, TraceEventType.Information)
        Dim Email = message.Sender.EmailAddress.Address
        Dim StaffName = message.Sender.EmailAddress.Name
        Dim BodyText = message.Body.Content.Replace("""", """""")
        'return UnsubscribeStaffAccess(Email, StaffName, BodyText, message) | UnsubscribeStaffSQLServer(Email, StaffName, BodyText, message);
        UnsubscribeStaff(Email, StaffName, BodyText, message, originalmessage, user)
        LogMessage(" ", message, TraceEventType.Information, "", True)
    End Sub

    Private Sub UnsubscribeStaff(ByVal Email As String, ByVal StaffName As String, ByVal BodyText As String, ByVal message As Microsoft.Graph.Message, ByVal originalmessage As Microsoft.Graph.Message, ByVal user As User)
        'return

        Dim xx = UnsubscribeStaffAccess(Email, StaffName, BodyText, message)
        Dim yy = UnsubscribeStaffSQLServer(Email, StaffName, BodyText, message)

        If xx = UnsubscribeStatus.Error Or yy = UnsubscribeStatus.Error Then
            'await graphClient.Users[user.Id].Messages["{originalmessage.Id}"]
            '    .Request()
            '    .UpdateAsync(originalmessage);

            originalmessage.Subject = "Unsub: Error: " & originalmessage.Subject.Replace("Unsub: Error: ", "").Replace("Unsub: Unsubscribed: ", "").Replace("Unsub: Not Found: ", "")
        Else

            If Not originalmessage.Subject.Contains("Unsub: Error: ") Then
                If xx = UnsubscribeStatus.Unsubscribed Or yy = UnsubscribeStatus.Unsubscribed Then
                    'await graphClient.Users[user.Id].Messages["{originalmessage.Id}"]
                    '    .Request()
                    '    .UpdateAsync(originalmessage);
                    originalmessage.Subject = "Unsub: Unsubscribed: " & originalmessage.Subject.Replace("Unsub: Unsubscribed: ", "").Replace("Unsub: Not Found: ", "")
                Else

                    If Not originalmessage.Subject.Contains("Unsub: Unsubscribed: ") Then
                        originalmessage.Subject = "Unsub: Not Found: " & originalmessage.Subject.Replace("Unsub: Not Found: ", "")
                        'await graphClient.Users[user.Id].Messages["{originalmessage.Id}"]
                        '    .Request()
                        '    .UpdateAsync(originalmessage);

                        AddNotFoundEmail(Email)
                    End If
                End If
            End If
        End If

        If originalmessage.Subject.Contains("Unsub: ") Then

            'graphClient
            ' .Users[user.Id]
            ' .MailFolders.Inbox
            ' .Messages[$"{originalmessage.Id}"]
            ' .Request()
            ' .UpdateAsync(originalmessage);

            'originalmessage.IsDraft = true;

            'graphClient
            '    .Users[user.Id]
            '    .MailFolders.Inbox
            '    .Messages[$"{originalmessage.Id}"]
            '    .Request()
            '    .UpdateAsync(originalmessage);

            'originalmessage.IsDraft = false;

            'originalmessage.IsDraft = true;

            Dim categories = New List(Of String)() From {
            }

            If originalmessage.Subject.Contains("Unsub: Error:") Then
                categories.Add("Error not unsubscribed")
            ElseIf originalmessage.Subject.Contains("Unsub: Unsubscribed:") Then
                categories.Add("Unsubscribed")
            ElseIf originalmessage.Subject.Contains("Unsub: Not Found:") Then
                categories.Add("Not Found")
            End If

            'var outlookCategory = new OutlookCategory
            '{
            '    Color = CategoryColor.Preset15
            '};

            Try
                'graphClient
                '    .Users[user.Id]
                '    .MailFolders.Inbox
                '    .Messages[$"{originalmessage.Id}"]
                '    .Request()
                '    .UpdateAsync(new Microsoft.Graph.Message()
                '    {
                '        Categories = categories
                '    });

                Dim subject = originalmessage.Subject
                originalmessage.IsDraft = True
                graphClient.Users(user.Id).MailFolders.Inbox.Messages($"{originalmessage.Id}").Request().UpdateAsync(New Microsoft.Graph.Message() With {
                    .Subject = subject
                })
                originalmessage.IsDraft = False
            Catch ex As ServiceException
                LogMessage("UnsubscribeStaff::UpdateAsync: Email message categories for staff with email " & Email & " not updated." & ex.Message, message, TraceEventType.Error)
            End Try

            'graphClient
            '    .Users[user.Id]
            '    .Outlook
            '    .MasterCategories($"{outlookCategory-id}")
            '    .Request()
            '    .UpdateAsync(originalmessage);

            'originalmessage.IsDraft = false;

            'Console.WriteLine("I am here!");

            'LogMessage("Message from sender " + originalmessage.Sender.EmailAddress.Address + " with subject '" + originalmessage.Subject + "' updated successfully.", originalmessage, TraceEventType.Information);
        End If
    End Sub

    Private Function UnsubscribeStaffAccess(ByVal Email As String, ByVal StaffName As String, ByVal BodyText As String, ByVal message As Microsoft.Graph.Message) As UnsubscribeStatus
        Dim StaffID As Object
        StaffID = DLookupAccess("Staff ID", "Staff", "[E-Mail] = """ & Email & """")

        If StaffID Is Nothing Then
            StaffID = 0
            LogMessage("Staff with email " & Email & " ***not found in Access database.", message, TraceEventType.Error)
            Return UnsubscribeStatus.NotFound
        End If

        Dim St3 = "UPDATE Staff SET Staff.NoMailing = True " & "WHERE [E-Mail] = """ & Email & """"

        If ExecuteNonQuery(St3, message) Is Nothing Then
            LogMessage("Staff " & StaffID.ToString() & ", Email " & Email & ", staff ***failed to unsubscribe.", message, TraceEventType.Error)
            Return UnsubscribeStatus.Error
        End If

        Dim NoMailing = DLookupAccess("NoMailing", "Staff", "[E-Mail] = """ & Email & """")

        If Conversions.ToBoolean(NoMailing) Then
            LogMessage("Staff " & StaffID.ToString() & ", Email " & Email & ", staff successfully unsubscribed.", message, TraceEventType.Information)
            St3 = "INSERT INTO StaffUnsubscribeRequests ( StaffID, Email, StaffName, Unsubsribed, Body ) " & "SELECT " & StaffID.ToString() & ", """ & Email & """, """ & StaffName & """, Now(), """ & BodyText & """"

            If Conversions.ToInteger(ExecuteNonQuery(St3, message)) >= 1 Then
                LogMessage("Staff " & StaffID.ToString() & ", Email " & Email & ", staff unsubscribe history updated.", message, TraceEventType.Error)
                Return UnsubscribeStatus.Unsubscribed
            Else
                LogMessage("Staff " & StaffID.ToString() & ", Email " & Email & ", staff unsubscribe history update ***failed.", message, TraceEventType.Error)
                Return UnsubscribeStatus.Error
            End If
        End If

        Return UnsubscribeStatus.Error
    End Function

    Private Function UnsubscribeStaffSQLServer(ByVal Email As String, ByVal StaffName As String, ByVal BodyText As String, ByVal message As Microsoft.Graph.Message) As UnsubscribeStatus
        Dim StaffID As Object
        StaffID = DLookupSQLServer("ID", "Staff", "Email = '" & Email & "'")

        If StaffID Is Nothing Then
            StaffID = 0
            LogMessage("Staff with email " & Email & " not found in SQL Server database.", message, TraceEventType.Error, "", False, LogChannels.SQLServer)
            Return UnsubscribeStatus.NotFound
        End If

        Dim St3 = "UPDATE Staff SET Staff.NoMailing = True " & "WHERE [E-Mail] = """ & Email & """"

        If ExecuteNonQuerySQLServer(St3, message) Is Nothing Then
            LogMessage("Staff " & StaffID.ToString() & ", Email " & Email & ", staff ***failed to unsubscribe.", message, TraceEventType.Error, "", False, LogChannels.SQLServer)
            Return UnsubscribeStatus.Error
        End If

        Dim NoMailing = DLookupSQLServer("NoMailing", "Staff", "Email = '" & Email & "'")

        If Conversions.ToBoolean(NoMailing) Then
            LogMessage("Staff " & StaffID.ToString() & ", Email " & Email & ", staff successfully unsubscribed.", message, TraceEventType.Information, "", False, LogChannels.SQLServer)
            St3 = "INSERT INTO StaffUnsubscribeRequests ( StaffID, Email, StaffName, Unsubsribed, Body ) " & "SELECT " & StaffID.ToString() & ", '" & Email & "', '" & StaffName & "', Now(), '" & BodyText & "'"

            If Conversions.ToInteger(ExecuteNonQuery(St3, message)) >= 1 Then
                LogMessage("Staff " & StaffID.ToString() & ", Email " & Email & ", staff unsubscribe history updated.", message, TraceEventType.Error, "", False, LogChannels.SQLServer)
                Return UnsubscribeStatus.Unsubscribed
            Else
                LogMessage("Staff " & StaffID.ToString() & ", Email " & Email & ", staff unsubscribe history update ***failed.", message, TraceEventType.Error, "", False, LogChannels.SQLServer)
                Return UnsubscribeStatus.Error
            End If
        End If

        Return UnsubscribeStatus.Error
    End Function

#End Region

    Private Sub UpdateReferencesSQLServer(ByVal message As Microsoft.Graph.Message, ByVal fePDFFileName As String, ByVal feFileName As String)
        Try
            Dim Conn = New SqlConnection(LocalSQLServerConnSt)
            'LogMessage("Opening reference SQL Server db connection.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
            Conn.Open()
            'LogMessage("Reference SQL Server db connection opened.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
            Dim StaffID As String
            Dim St = ""

            If fePDFFileName.IndexOf("-") <> 0 Then
                Dim command As SqlCommand
                Dim i = 0
                Dim j = 0
                StaffID = LeftStr(fePDFFileName, fePDFFileName.IndexOf("-") - 1)
                St = "UPDATE Staff SET Staff.LastEditableReference1Data = '" & feFileName & "', ReferenceReceived1 = { fn Now } " & "WHERE (((Staff.ID) = " & StaffID & ")) AND (Staff.LastEditableReference = " & fePDFFileName & ")"

                Try
                    command = New SqlCommand(St, Conn)
                    j = command.ExecuteNonQuery()
                    command.Dispose()
                Catch ex As Exception
                    LogError("Error updating reference SQL Server db: " & ex.Message, message, TraceEventType.Error, "", False, LogChannels.SQLServer)
                End Try

                St = "UPDATE Staff SET Staff.LastEditableReference2Data = '" & feFileName & "', ReferenceReceived2 = { fn now } " & "WHERE (((Staff.ID) = " & StaffID & ")) AND (Staff.LastEditableReference2 = " & fePDFFileName & ")"

                Try
                    command = New SqlCommand(St, Conn)
                    j = command.ExecuteNonQuery()
                    command.Dispose()
                Catch ex As Exception
                    LogError("Error updating reference SQL Server db: " & ex.Message, message, TraceEventType.Error, "", False, LogChannels.SQLServer)
                End Try

                Conn.Close()

                If i > 0 Then
                    LogMessage("Updated reference 1 SQL Server db.", message, TraceEventType.Information)
                    message.Subject = MSG_SAVED & message.Subject
                    message.IsRead = True
                    'message.Update(ConflictResolutionMode.AlwaysOverwrite);

                    graphClient.Me.Messages("{message.Id}").Request().UpdateAsync(message)
                ElseIf j > 0 Then
                    LogMessage("Updated reference 2 SQL Server db.", message, TraceEventType.Information, "", False, LogChannels.SQLServer)
                    message.Subject = MSG_SAVED & message.Subject
                    message.IsRead = True
                    'message.Update(ConflictResolutionMode.AlwaysOverwrite);

                    graphClient.Me.Messages("{message.Id}").Request().UpdateAsync(message)
                Else
                    LogMessage("No match for file " & " found in references: ", message, TraceEventType.Error, "", False, LogChannels.SQLServer)
                End If
            End If

        Catch ex As Exception
            LogError("Error updating reference SQL Server db: " & ex.Message, message, TraceEventType.Error, "", False, LogChannels.SQLServer)
        End Try
    End Sub

    Private Sub UpdateEditableTiemsheetSQLServer(ByVal message As Microsoft.Graph.Message, ByVal fePDFFileName As String, ByVal feFileName As String)
        Try
            Dim ConnSt = "Data Source=DB-SERVER\SQL2K14;Initial Catalog=EMS_2018_HS;Persist Security Info=True;User ID=EMS_2018_HS;Password=P@ssword01"
            Dim Conn = New SqlConnection(ConnSt)
            'LogMessage("Opening SQL Server db connection.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
            Conn.Open()
            'LogMessage("SQL Server db connection opened.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
            Dim EventID As String
            Dim St = ""

            If fePDFFileName.IndexOf("-") <> 0 Then
                EventID = LeftStr(fePDFFileName, fePDFFileName.IndexOf("-") - 1)
                St = "UPDATE Events SET Events.LastEditableTimesheetData = '" & feFileName & "' " & "WHERE (((Events.ID)= " & EventID & "))"
                Dim command = New SqlCommand(St, Conn)
                command.ExecuteNonQuery()
                Conn.Close()
                LogMessage("Updated SQL Server db.", message, TraceEventType.Information, "", False, LogChannels.SQLServer)
                message.Subject = MSG_SAVED & message.Subject
                message.IsRead = True
                'message.Update(ConflictResolutionMode.AlwaysOverwrite);

                graphClient.Me.Messages("{message.Id}").Request().UpdateAsync(message)
                Dim FullPath As String = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") & " Timesheet data for event " & EventID.ToString() & " saved.txt")

                Try
                    Dim tw As TextWriter = IO.File.CreateText(FullPath)
                    tw.Close()
                Catch ex As Exception
                    LogError("Error writitng to log file: " & FullPath & "Exception: " & ex.Message, message, TraceEventType.Error, "", False, LogChannels.SQLServer)
                End Try
            End If

        Catch ex As Exception
            LogError("Error updating SQL Server db: " & ex.Message, message, TraceEventType.Error, "", False, LogChannels.SQLServer)
        End Try
    End Sub

    Private Sub UpdateReferencesAccess(ByVal message As Microsoft.Graph.Message, ByVal fePDFFileName As String, ByVal feFileName As String)
        Try
            Dim Conn = New OleDbConnection(LocalAccessConnSt)
            'LogMessage("Opening reference db connection.", message, TraceEventType.Information);
            Conn.Open()
            'LogMessage("Reference db connection opened.", message, TraceEventType.Information);
            Dim StaffID As String
            Dim St = ""

            If fePDFFileName.IndexOf("-") <> 0 Then
                Dim command As OleDbCommand
                Dim i = 0
                Dim j = 0
                StaffID = LeftStr(fePDFFileName, fePDFFileName.IndexOf("-") - 1)
                St = "UPDATE Staff SET Staff.[Last Editable Reference 1 Data] = """ & feFileName & """, [Reference Received 1] = NOW() " & "WHERE (((Staff.[Staff ID]) = " & StaffID & ")) AND (Staff.[Last_Editable_Reference_1] = " & fePDFFileName & ")"

                Try
                    command = New OleDbCommand(St, Conn)
                    j = command.ExecuteNonQuery()
                    command.Dispose()
                Catch ex As Exception
                    LogError("Error updating reference db: " & ex.Message, message, TraceEventType.Error)
                End Try

                St = "UPDATE Staff SET Staff.[Last Editable Reference 2 Data] = """ & feFileName & """, [Reference Received 2] = NOW() " & "WHERE (((Staff.[Staff ID]) = " & StaffID & ")) AND (Staff.[Last_Editable_Reference_2] = " & fePDFFileName & ")"

                Try
                    command = New OleDbCommand(St, Conn)
                    j = command.ExecuteNonQuery()
                    command.Dispose()
                Catch ex As Exception
                    LogError("Error updating reference db: " & ex.Message, message, TraceEventType.Error)
                End Try

                Conn.Close()

                If i > 0 Then
                    LogMessage("Updated reference 1 db.", message, TraceEventType.Information)
                    message.Subject = MSG_SAVED & message.Subject
                    message.IsRead = True
                    'message.Update(ConflictResolutionMode.AlwaysOverwrite);

                    graphClient.Me.Messages("{message.Id}").Request().UpdateAsync(message)
                ElseIf j > 0 Then
                    LogMessage("Updated reference 2 db.", message, TraceEventType.Information)
                    message.Subject = MSG_SAVED & message.Subject
                    message.IsRead = True
                    'message.Update(ConflictResolutionMode.AlwaysOverwrite);

                    graphClient.Me.Messages("{message.Id}").Request().UpdateAsync(message)
                Else
                    LogMessage("No match for file " & " found in references: ", message, TraceEventType.Error)
                End If
            End If

        Catch ex As Exception
            LogError("Error updating reference db: " & ex.Message, message, TraceEventType.Error)
        End Try
    End Sub

    Private Sub UpdateEditableTiemsheetAccess(ByVal message As Microsoft.Graph.Message, ByVal fePDFFileName As String, ByVal feFileName As String)
        Try
            Dim ConnSt = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""S:\Events Data\Events Data.mdb"";"
            Dim Conn = New OleDbConnection(ConnSt)
            'LogMessage("Opening db connection.", message, TraceEventType.Information);
            Conn.Open()
            'LogMessage("Db connection opened.", message, TraceEventType.Information);
            Dim EventID As String
            Dim St = ""

            If fePDFFileName.IndexOf("-") <> 0 Then
                EventID = LeftStr(fePDFFileName, fePDFFileName.IndexOf("-") - 1)
                St = "UPDATE Events SET Events.[Last Editable Timesheet Data] = """ & feFileName & """ " & "WHERE (((Events.[Event ID])= " & EventID & "))"
                Dim command = New OleDbCommand(St, Conn)
                command.ExecuteNonQuery()
                Conn.Close()
                LogMessage("Updated db.", message, TraceEventType.Information)
                message.Subject = MSG_SAVED & message.Subject
                message.IsRead = True
                'message.Update(ConflictResolutionMode.AlwaysOverwrite);

                graphClient.Me.Messages("{message.Id}").Request().UpdateAsync(message)
                Dim FullPath As String = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") & " Timesheet data for event " & EventID.ToString() & " saved.txt")

                Try
                    Dim tw As TextWriter = IO.File.CreateText(FullPath)
                    tw.Close()
                Catch ex As Exception
                    LogError("Error writitng to log file: " & FullPath & "Exception: " & ex.Message, message, TraceEventType.Error)
                End Try
            End If

        Catch ex As Exception
            LogError("Error updating db: " & ex.Message, message, TraceEventType.Error)
        End Try
    End Sub

    Private Function GetPDFFileName(ByVal feFileName As String, ByVal item As Microsoft.Graph.Message, ByRef PT As PDFType) As String
        Dim fePDFFileName = ""
        Dim fePDFFileName1 = ""
        'System.Windows.Forms.Application.DoEvents();

        Dim r = New Regex("\d{1,2}\-\d{1,2}\-\d{4}", RegexOptions.IgnoreCase)
        Dim m = r.Match(feFileName)

        If m.Success Then
            Dim i As Integer

            If feFileName.IndexOf("]") > 0 Then
                i = m.Index + 11
                fePDFFileName = Strings.Left(feFileName, i) & ".pdf"
                LogMessage("feFileName.IndexOf(""]"") > 0: feFileName: " & feFileName & ", fePDFFileName: " & fePDFFileName, item, TraceEventType.Information)
            Else
                Dim DateSt = Mid(feFileName, m.Index, 10)
                i = m.Index + 10
                fePDFFileName = Strings.Left(feFileName, i) & ".pdf"
                fePDFFileName1 = fePDFFileName.Replace(DateSt, "[" & DateSt & "]")
                LogMessage("feFileName.IndexOf(""]"") <= 0: feFileName: " & feFileName & ", fePDFFileName: " & fePDFFileName, item, TraceEventType.Information)
            End If

            LogMessage("fePDFFileNamet: " & fePDFFileName & ", fePDFFileName1: " & fePDFFileName1, item, TraceEventType.Information)

            If IO.File.Exists(Path.Combine(TIMESHEETS_PDF_PATH, fePDFFileName)) Then
                LogMessage("GetPDFFileName: Exists: " & fePDFFileName, item, TraceEventType.Information)
                PT = PDFType.Timesheet
                Return fePDFFileName
            ElseIf IO.File.Exists(Path.Combine(TIMESHEETS_PDF_PATH, fePDFFileName1)) Then
                LogMessage("GetPDFFileName: Exists: " & fePDFFileName1, item, TraceEventType.Information)
                PT = PDFType.Timesheet
                Return fePDFFileName1
            Else
                Return Nothing
            End If
        Else
            ' No date found
            LogMessage("GetPDFFileName: No match, feFileName: " & feFileName, item, TraceEventType.Information)
            Return Nothing
        End If

        ' If feFileName.IndexOf("]") > 0 Then
        ' fePDFFileName = LeftStr(feFileName, feFileName.IndexOf("]")).ToLower().Replace("_data", "") & ".pdf"
        ' LogMessage("feFileName.IndexOf(""]"") > 0: feFileName: " & feFileName & ", fePDFFileName: " & fePDFFileName, Message, TraceEventType.Information)
        ' Else
        ' fePDFFileName = feFileName.ToLower().Replace(".fdf", ".pdf").Replace("_data", "")
        ' LogMessage("feFileName.IndexOf(""]"") <= 0: feFileName: " & feFileName & ", fePDFFileName: " & fePDFFileName, Message, TraceEventType.Information)
        ' Dim DateSt As String = Microsoft.VisualBasic.Right(fePDFFileName, 14)
        ' fePDFFileName1 = fePDFFileName.Replace(DateSt, "[" & DateSt)
        ' fePDFFileName1 = fePDFFileName1.Replace(".pdf", "].pdf")
        ' LogMessage("fePDFFileNamet: " & fePDFFileName & ", fePDFFileName1: " & fePDFFileName1 & ", DateSt: " & DateSt, Message, TraceEventType.Information)
        ' End If

    End Function

    Private Shared Function LeftStr(ByVal param As String, ByVal length As Integer) As String
        Dim result = param.Substring(0, length)
        Return result
    End Function

    Private Shared Function RightStr(ByVal param As String, ByVal length As Integer) As String
        Dim result = param.Substring(param.Length - length, length)
        Return result
    End Function

    Private Shared Function Mid(ByVal param As String, ByVal startIndex As Integer, ByVal length As Integer) As String
        Dim result = param.Substring(startIndex, length)
        Return result
    End Function

    Private Shared Function Mid(ByVal param As String, ByVal startIndex As Integer) As String
        Dim result = param.Substring(startIndex)
        Return result
    End Function

    Private Sub LogMessage(ByVal eventName As String, ByVal item As Microsoft.Graph.Message, ByVal e As TraceEventType, ByVal Optional message As String = "", ByVal Optional CRLF As Boolean = False, ByVal Optional channel As LogChannels = LogChannels.Access, ByVal Optional UI As Boolean = False)
        If Not IO.Directory.Exists(SYSTEM_LOG_PATH) Then
            IO.Directory.CreateDirectory(SYSTEM_LOG_PATH)
        End If

        If UI Then
            Windows.Forms.Application.DoEvents()
            Me.Msg.Text = eventName
            'System.Windows.Threading.Dispatcher.Invoke(new Action(() => Msg.Text = eventName));
            Windows.Forms.Application.DoEvents()
        End If

        Dim FullPath = ""

        If channel = LogChannels.Access Then
            FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") & " " & SYSTEM_LOG_ACCESS)
        Else
            FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") & " " & SYSTEM_LOG_SQLSERVER)
        End If

        Dim tw As TextWriter = IO.File.AppendText(FullPath)

        Try

            If String.IsNullOrEmpty(eventName) Then
                tw.WriteLine("---")
                tw.WriteLine(" ")
            Else
                tw.WriteLine(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & vbTab & eventName & " - " & If(item Is Nothing, "", item.Sender.EmailAddress.Address) & " - " & If(item Is Nothing, "", item.Subject) & " : " & message)
            End If

            If CRLF Then
                tw.WriteLine(" ")
                tw.WriteLine(" ")
            End If

        Catch ex As Exception

            If UI Then
                MsgBox(ex.Message)
            End If

            tw.WriteLine(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & vbTab & eventName & " - " & If(item Is Nothing, "", item.Sender.EmailAddress.Address) & " - " & If(item Is Nothing, "", item.Subject) & " : " & ex.Message)
        Finally
            tw.Close()
        End Try
    End Sub

    Private Sub LogError(ByVal eventName As String, ByVal item As Microsoft.Graph.Message, ByVal e As TraceEventType, ByVal Optional Message As String = "", ByVal Optional CRLF As Boolean = False, ByVal Optional channel As LogChannels = LogChannels.Access, ByVal Optional UI As Boolean = False)
        If Not IO.Directory.Exists(SYSTEM_LOG_PATH) Then
            IO.Directory.CreateDirectory(SYSTEM_LOG_PATH)
        End If

        If UI Then
            Windows.Forms.Application.DoEvents()
            Me.Msg.Text = eventName
            Windows.Forms.Application.DoEvents()
        End If

        Dim FullPath = ""

        If channel = LogChannels.Access Then
            FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") & " ERROR - " & SYSTEM_LOG_ACCESS)
        Else
            FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") & " ERROR - " & SYSTEM_LOG_SQLSERVER)
        End If

        Dim tw As TextWriter = IO.File.AppendText(FullPath)

        Try

            If CRLF Then
                tw.WriteLine(" ")
                tw.WriteLine(" ")
            End If

            ' Dim tw As TextWriter = System.IO.File.AppendText(FullPath)

            If String.IsNullOrEmpty(eventName) Then
                tw.WriteLine("---")
                tw.WriteLine(" ")
            Else
                tw.WriteLine(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & vbTab & eventName & " - " & If(item Is Nothing, "", item.From.EmailAddress.Address) & " - " & If(item Is Nothing, "", item.Subject) & " : " & Message)
            End If

            ' tw.Close()
        Catch ex As Exception

            If UI Then
                MsgBox(ex.Message)
            End If

            tw.WriteLine(Date.Now.ToString("yyyy/MM/dd HH:mm:ss") & vbTab & eventName & " - " & If(item Is Nothing, "", item.Sender.EmailAddress.Address) & " - " & If(item Is Nothing, "", item.Subject) & " : " & ex.Message)
        Finally
            tw.Close()
        End Try
    End Sub

    Private Sub AddNotFoundEmail(ByVal email As String, ByVal Optional CRLF As Boolean = False, ByVal Optional channel As LogChannels = LogChannels.Access)
        If Not IO.Directory.Exists(SYSTEM_LOG_PATH) Then
            IO.Directory.CreateDirectory(SYSTEM_LOG_PATH)
        End If

        Windows.Forms.Application.DoEvents()
        Me.Msg.Text = "Adding not found email: " & email
        Windows.Forms.Application.DoEvents()
        Dim FullPath = ""

        If channel = LogChannels.Access Then
            FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd") & " " & SYSTEM_LOG_ACCESS_NOTFOUND)
        Else
            FullPath = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd") & " " & SYSTEM_LOG_SQLSERVER_NOTFOUND)
        End If

        Dim tw As TextWriter = IO.File.AppendText(FullPath)

        Try

            If CRLF Then
                tw.WriteLine(" ")
                tw.WriteLine(" ")
            End If

            If String.IsNullOrEmpty(email) Then
                tw.WriteLine("---")
                tw.WriteLine(" ")
            Else
                tw.WriteLine(email)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            tw.Close()
        End Try
    End Sub

    Private Async Function ProcessStaffWorkRequests(ByVal graphClient As GraphServiceClient) As Task
        Dim Users = Await graphClient.Users.Request().Filter("startswith(mail,'workonevent')").GetAsync()
        Dim user = Users(0)

        'var inboxMessages = await graphClient
        '     .Users[user.Id]
        '     .MailFolders.Inbox
        '     .Messages
        '     .Request()
        '     .Filter("not startswith(Subject,'Unsub: ')")
        '     .GetAsync();

        'var inboxMessages = await graphClient
        '     .Users[user.Id]
        '     .MailFolders.Inbox
        '     .Messages
        '     .Request()
        '     .Filter("not isRead")
        '     .GetAsync();

        Dim inboxMessages = Await graphClient.Users(user.Id).MailFolders.Inbox.Messages.Request().Filter("not startswith(Subject,'SAVED: ')").GetAsync()

        For Each x In inboxMessages
            'if (x.IsRead == false)
            '{
            ProcessStaffWorkRequestsMessageAsync(x, user)
            '}
        Next
    End Function

    Private Sub ProcessStaffWorkRequestsMessageAsync(ByVal EmailMsg As Microsoft.Graph.Message, ByVal user As User)
        LogMessage("Message " & EmailMsg.Subject & " loaded. Ready for message body processing.", EmailMsg, TraceEventType.Information)

        ' Try
        If InStr(EmailMsg.Body.Content, "Event Ref") <> 0 Then
            LogMessage("Processing message body.", EmailMsg, TraceEventType.Information)
        Else
            Return
        End If

        Dim Body1 = StripHtmlTags(EmailMsg.Body.Content.Replace("<BR>", Environment.NewLine)).Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
        Dim StaffID As Object = -1
        Dim EventID As Object = -1

        For Each z In Body1
            Dim y = z
            y = y.Replace(vbLf, "")

            If y.Contains("Username:") Then
                Dim StaffEmail = LTrim(RTrim(Strings.Right(y, y.Length - 9)))
                StaffID = DLookupSQLServerRemote("StaffID", "StaffSecurityUser", "Email = '" & StaffEmail & "'")

                If StaffID Is Nothing Then
                    StaffID = 0
                End If
            End If

            If y.Contains("Event Ref:") Then
                Dim EventRef = LTrim(RTrim(Strings.Right(y, y.Length - 10)))
                EventID = DLookupSQLServerRemote("EventID", "UpcomingJobs", "ID = " & EventRef)

                If EventID Is Nothing Then
                    EventID = 0
                End If
            End If
        Next

        LogMessage("Staff ID " & StaffID.ToString() & " Event ID " & EventID.ToString(), EmailMsg, TraceEventType.Information)

        If Conversions.ToInteger(StaffID) <> 0 And Conversions.ToInteger(EventID) <> 0 Then
            Dim ClientID = DLookupAccess("Client ID", "Events", "[Event ID] = " & EventID.ToString())

            If ClientID Is Nothing Then
                ClientID = 0
            Else
                Dim SIDNP = DLookupAccess("StaffID", "Not Prefered Staff", "StaffID = " & StaffID.ToString() & " AND [Client ID] = " & ClientID.ToString())

                If SIDNP Is Nothing Then
                    SIDNP = 0
                Else
                    LogMessage("Staff ID " & StaffID.ToString() & " in not Preferred list for client id " & ClientID.ToString() & ". Skipping.", EmailMsg, TraceEventType.Information)
                    Return
                End If

                Dim SIDP = DLookupAccess("StaffID", "Prefered Staff", "StaffID = " & StaffID.ToString() & " AND [ClientID] = " & ClientID.ToString())

                If SIDP Is Nothing Then
                    SIDP = 0
                Else
                    LogMessage("Staff ID " & StaffID.ToString() & " in Preferred list for client id " & ClientID.ToString() & ". Skipping.", EmailMsg, TraceEventType.Information)
                End If

                Dim St1 As String = "SELECT Events.[Event ID] " & "FROM Events INNER JOIN [Staff Bookings] ON Events.[Event ID] = [Staff Bookings].[Event ID] " & "WHERE (((Events.[Event ID])=" & EventID.ToString() & ") AND (([Staff Bookings].[Staff ID])=" & StaffID.ToString() & ") AND ((Events.Status)=""Current"") AND (([Staff Bookings].[Staff Total])>0) AND ((Events.Date)<=Date()))"
                Dim EID = ExecuteScalar(St1, EmailMsg)

                If EID Is Nothing Then
                    EID = 0
                End If

                LogMessage("EID " & EID.ToString(), EmailMsg, TraceEventType.Information)
                Dim St2 = "INSERT INTO StaffWorkInterest ( EventID, StaffID, P, W, Sort ) "
                St2 = St2 & "SELECT " & EventID.ToString() & ", " & StaffID.ToString() & ", """ & IIf(Conversions.ToInteger(SIDP) <> 0, "P", "") & """, """ & IIf(Conversions.ToInteger(EID) <> 0, "W", "") & """, " & IIf(Conversions.ToInteger(SIDP) <> 0, "1", IIf(Conversions.ToInteger(EID) <> 0, "2", "3"))

                If Conversions.ToInteger(ExecuteNonQuery(St2, EmailMsg)) >= 1 Then
                    LogMessage("Staff " & StaffID.ToString() & ", Event " & EventID.ToString() & ", request to work saved.", EmailMsg, TraceEventType.Information)
                    Dim FullPath As String = Path.Combine(SYSTEM_LOG_PATH, CurrentTimeStamp.ToString("yyyy-MM-dd HH-mm-ss") & " Staff " & StaffID.ToString() & " work request data for event " & EventID.ToString() & " saved.txt")

                    Try
                        Dim tw As TextWriter = IO.File.CreateText(FullPath)
                        tw.Close()
                    Catch ex As Exception
                        LogMessage("ProcessStaffWorkRequestsMessageAsync::Staff " & StaffID.ToString() & ", Event " & EventID.ToString() & ", exception: " & ex.Message, EmailMsg, TraceEventType.Information)
                    End Try
                End If

                Dim St3 As String = "UPDATE Events SET Events.JobRequests = IIf(IsNull([JobRequests]),1,[JobRequests]+1) " & "WHERE Events.[Event ID] = " & EventID.ToString()

                If Conversions.ToBoolean(Operators.ConditionalCompareObjectGreaterEqual(ExecuteNonQuery(St3, EmailMsg), 1, False)) Then
                    LogMessage("Staff " & StaffID.ToString() & ", Event " & EventID.ToString() & ", event staff job request counter incremented.", EmailMsg, TraceEventType.Information)
                End If

                Dim BodyText As String = Replace("Staff ID: " & StaffID.ToString() & vbCrLf & vbCrLf & "Event ID: " & EventID.ToString() & vbCrLf & vbCrLf & "=====================================" & vbCrLf & vbCrLf & EmailMsg.Body.Content, vbCrLf, "<BR>")
                Dim subject = EmailMsg.Subject.Replace(MSG_SAVED, "").Replace(MSG_NOT_SAVED, "")
                subject = MSG_SAVED & subject

                'EmailMsg.Body.Content = BodyText;

                Dim b As ItemBody = New ItemBody()
                b.Content = BodyText
                EmailMsg.IsDraft = True
                graphClient.Users(user.Id).MailFolders.Inbox.Messages($"{EmailMsg.Id}").Request().UpdateAsync(New Microsoft.Graph.Message() With {
                    .Subject = subject,
                    .Body = b
                })
                EmailMsg.IsDraft = False
            End If
        End If

        ' Catch ex As Exception
        ' LogError("ProcessStaffWorkRequests Exception", message, ex.Message)
        ' End Try

    End Sub

    Private Async Function ProcessTimesheetAttachmentsAsync(ByVal graphClient As GraphServiceClient) As Task
        Dim Users = Await graphClient.Users.Request().Filter("startswith(mail,'timesheets')").GetAsync()
        Dim user = Users(0)
        Dim inboxMessages = Await graphClient.Users(user.Id).MailFolders.Inbox.Messages.Request().Filter("not startswith(Subject,'SAVED: ')").GetAsync()

        For Each x In inboxMessages
            'if (x.IsRead == false)
            '{
            Await ProcessStaffTimesheetsAsync(x, user)
            '}
        Next
    End Function

    Private Async Function ProcessStaffTimesheetsAsync(ByVal message As Microsoft.Graph.Message, ByVal user As User) As Task
        LogMessage("", Nothing, TraceEventType.Information)
        LogMessage("Message " & message.Subject & " loaded. Ready for attachment processing.", message, TraceEventType.Information)
        Dim FDFAttchmentFound = False

        If message.HasAttachments = True Then
            Dim attachments = Await graphClient.Users(user.Id).Messages($"{message.Id}").Attachments.Request().GetAsync()
            LogMessage("Iterating through " & attachments.Count.ToString() & " attachments.", message, TraceEventType.Information)

            'foreach (var attachment in message.Attachments)
            For Each attachment In attachments
                'if (attachment.ODataType == "#microsoft.graph.itemAttachment")
                '{

                '    var attachmentRequest = await graphClient
                '                                .Users[user.Id]
                '                                .MailFolders
                '                                .Inbox
                '                                .Messages[message.Id]
                '                                .Attachments[attachment.Id]
                '                                .Request()
                '                                .Expand("microsoft.graph.itemattachment/item")
                '                                .GetAsync();

                '    var itemAttachment = (ItemAttachment)attachmentRequest;
                '}


                'var a1 = await graphClient
                '           .Users[user.Id]
                '           .Messages[$"{message.Id}"]
                '           .Attachments[$"{attachment.Id}"]
                '           .Request()
                '           .GetAsync();

                If TypeOf attachment Is FileAttachment Then
                    Dim fileAttachment As FileAttachment = TryCast(attachment, FileAttachment)
                    LogMessage("Processing attachment " & fileAttachment.Name, message, TraceEventType.Information)

                    If fileAttachment.Name.Contains(".FDF") Or fileAttachment.Name.Contains(".fdf") Then
                        FDFAttchmentFound = True
                        LogMessage("File attachment name: " & fileAttachment.Name, message, TraceEventType.Information)

                        If fileAttachment.Name.Length >= 3 Then
                            Dim feFileExtension = fileAttachment.Name.Substring(fileAttachment.Name.Length - 4, 4)
                            Dim feFileName = fileAttachment.Name.Replace("%20", " ")
                            LogMessage("Attachment Ext: " & feFileExtension, message, TraceEventType.Information)

                            If Equals(feFileExtension.ToLower(), ".fdf") Then
                                message.Subject = message.Subject.Replace(MSG_SAVED, "")
                                message.Subject = message.Subject.Replace(MSG_NOT_SAVED, "")
                                Dim PT As PDFType = Nothing
                                Dim fePDFFileName = GetPDFFileName(feFileName, message, PT)

                                If Not Equals(fePDFFileName, Nothing) Then
                                    LogMessage("Saving attachment, file name: " & feFileName, message, TraceEventType.Information)

                                    'fileAttachment.Load(Path.Combine(TIMESHEETS_DATA_PATH, feFileName));

                                    SaveAttachment(fileAttachment, Path.Combine(TIMESHEETS_DATA_PATH, feFileName))
                                    LogMessage("Attachment saved.", message, TraceEventType.Information)
                                    LogMessage("Updating db...", message, TraceEventType.Information)
                                    UpdateEditableTiemsheetAccess(message, fePDFFileName, feFileName)
                                    UpdateEditableTiemsheetSQLServer(message, fePDFFileName, feFileName)
                                ElseIf IO.File.Exists(REFERENCES_PDF_PATH & fePDFFileName) Then
                                    LogMessage("Saving reference attachment, file name: " & feFileName, message, TraceEventType.Information)

                                    'fileAttachment.Load(Path.Combine(REFERENCES_DATA_PATH, feFileName));

                                    SaveAttachment(fileAttachment, Path.Combine(REFERENCES_DATA_PATH, feFileName))
                                    LogMessage("Reference attachment saved.", message, TraceEventType.Information)
                                    LogMessage("Updating Reference db.", message, TraceEventType.Information)
                                    UpdateReferencesAccess(message, fePDFFileName, feFileName)
                                    UpdateReferencesSQLServer(message, fePDFFileName, feFileName)
                                Else
                                    LogMessage("Data file " & feFileName & " not saved, no matching pdf file '" & TIMESHEETS_PDF_PATH & fePDFFileName & " found.", message, TraceEventType.Error)
                                    LogError("Data file " & feFileName & " not saved, no matching pdf file '" & TIMESHEETS_PDF_PATH & fePDFFileName & " found.", message, TraceEventType.Error)
                                    SetMessageSubject(message, user, MSG_NOT_SAVED)
                                End If
                            End If
                        End If

                        fileAttachment = Nothing
                    Else
                        LogMessage("Attachment " & fileAttachment.Name & " is an invalid item attachment type, skipping.", Nothing, TraceEventType.Information)
                    End If
                Else
                    LogMessage("Attachment " & attachment.Name & "is an item attachment, skipping attachment.", message, TraceEventType.Information)
                End If
            Next
        Else
            SetMessageSubject(message, user, MSG_NOT_SAVED)
            Return
        End If

        If Not FDFAttchmentFound Then
            LogMessage("No attachment in the message valid for processing, skipping message.", message, TraceEventType.Information)
            SetMessageSubject(message, user, MSG_NOT_SAVED)
            Return
        End If
    End Function

    Private Sub SaveAttachment(ByVal fileAttachment As FileAttachment, ByVal p As String)
        'using (FileStream outputFileStream = new FileStream(p, FileMode.Create))
        '{
        '    fileAttachment.ContentBytes.CopyTo(outputFileStream);

        '    fileAttachment.ContentBytes.CopyTo()
        '}

        Const fileContent = New Buffer(attachment.contentBytes, 'base64');

        IO.File.WriteAllBytes(p, fileContent)
    End Sub

    Private Sub SetMessageSubject(ByVal message As Microsoft.Graph.Message, ByVal user As User, ByVal v As String)
        Dim subject = message.Subject.Replace(MSG_SAVED, "").Replace(MSG_NOT_SAVED, "").Replace(v, "")
        subject = v & subject
        message.IsDraft = True
        graphClient.Users(user.Id).MailFolders.Inbox.Messages($"{message.Id}").Request().UpdateAsync(New Microsoft.Graph.Message() With {
            .Subject = subject
        })
        message.IsDraft = False
    End Sub

    Public Async Sub Process()
        Me.Timer1.Enabled = False
        CurrentTimeStamp = Date.Now
        Me.btnProcess.Enabled = False
        InProcess = True
        'ProcessTimesheetAttachments();
        'ProcessStaffWorkRequests();

        Dim LastRun = Date.Now

        If isProcessUnsubscribeEmailsOK() Then

            'LogMessage("", null, TraceEventType.Information, "", true);

            LogMessage("Begin unsubscribe staff processing...", Nothing, TraceEventType.Information, "", True)
            Await ProcessUnsubscribeEmailsAsync(graphClient)

            'LogMessage("End unsubscribe staff processing...", null, TraceEventType.Information, "", true);
            'LogMessage(" ", null, TraceEventType.Information, "", true);

            LogMessage("Begin processing staff work requests...", Nothing, TraceEventType.Information, "", True)
            Await ProcessStaffWorkRequests(graphClient)

            LogMessage("Begin processing timesheet attachments...", Nothing, TraceEventType.Information, "", True)
            Await ProcessTimesheetAttachmentsAsync(graphClient)

            'Properties.Settings.Default.UnsubscribeServiceLastRun = DateTime.Now.ToOADate();
            Email_Processor.Properties.Settings.Default.UnsubscribeServiceLastRun = LastRun.ToOADate()
            Dim LastRun1 As Double = Email_Processor.Properties.Settings.Default.UnsubscribeServiceLastRun
            Email_Processor.Properties.Settings.Default.Save()
            Email_Processor.Properties.Settings.Default.Upgrade()
            'System.Windows.Forms.Application.Restart();
        End If

        InProcess = False
        Me.btnProcess.Enabled = True
        Windows.Forms.Application.DoEvents()
        Me.Timer1.Enabled = True
    End Sub

    Private Function isProcessUnsubscribeEmailsOK() As Boolean
        Return True

        'DateTime UnsubscribeServiceLastRun = DateTime.FromOADate(Properties.Settings.Default.UnsubscribeServiceLastRun);

        '''int UnsubscribeServiceInterval = (int)Math.Round(60000f * Properties.Settings.Default.UnsubscribeServiceTimeinMinutes);

        'if (DateTime.Now >= UnsubscribeServiceLastRun.AddMinutes(Properties.Settings.Default.UnsubscribeServiceTimeinMinutes))
        '    return true;
        'else
        '    return false;
    End Function

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As EventArgs)
        Me.Timer1.Enabled = False

        If InProcess Then
            Return
        End If

        Process()
        Me.Timer1.Interval = CInt(Math.Round(60000F * Email_Processor.Properties.Settings.Default.TimerTimeinMinutes))
        Me.Timer1.Enabled = True
    End Sub

    'private static GraphServiceClient GetClient(string accessToken, IHttpProvider provider = null)
    '{
    '    var delegateAuthProvider = new DelegateAuthenticationProvider((requestMessage) =>
    '    {
    '        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

    '        return Task.FromResult(0);
    '    });

    '    var graphClient = new GraphServiceClient(delegateAuthProvider, provider ?? HttpProvider);

    '    return graphClient;
    '}

    Private Shared Function GetAuthenticatedGraphClient(ByVal config As IConfigurationRoot) As GraphServiceClient
        Dim authenticationProvider = CreateAuthorizationProvider(config)
        _graphServiceClient = New GraphServiceClient(authenticationProvider)
        Return _graphServiceClient
    End Function

    'private static HttpClient GetAuthenticatedHTTPClient(IConfigurationRoot config)
    '{
    '    var authenticationProvider = CreateAuthorizationProvider(config);
    '    _httpClient = new HttpClient(new AuthHandler(authenticationProvider, new HttpClientHandler()));
    '    return _httpClient;
    '}

    Private Shared Function CreateAuthorizationProvider(ByVal config As IConfigurationRoot) As IAuthenticationProvider
        Dim clientId = config("applicationId")
        Dim clientSecret = config("applicationSecret")
        Dim redirectUri = config("redirectUri")
        Dim authority = $"https://login.microsoftonline.com/{config("tenantId")}/v2.0"

        'this specific scope means that application will default to what is defined in the application registration rather than using dynamic scopes
        Dim scopes As List(Of String) = New List(Of String)()
        scopes.Add("https://graph.microsoft.com/.default")
        Dim cca = ConfidentialClientApplicationBuilder.Create(clientId).WithAuthority(authority).WithRedirectUri(redirectUri).WithClientSecret(clientSecret).Build()
        Return New Email_Processor.MsalAuthenticationProvider(cca, scopes.ToArray())
    End Function

    Public Shared Function LoadAppSettings() As IConfigurationRoot
        Try
            Dim config = New ConfigurationBuilder().SetBasePath(IO.Directory.GetCurrentDirectory()).AddJsonFile("appsettings.json", False, True).Build()

            ' Validate required settings
            If String.IsNullOrEmpty(config("applicationId")) OrElse String.IsNullOrEmpty(config("applicationSecret")) OrElse String.IsNullOrEmpty(config("redirectUri")) OrElse String.IsNullOrEmpty(config("tenantId")) OrElse String.IsNullOrEmpty(config("domain")) Then
                Return Nothing
            End If

            Return config
        Catch __unusedFileNotFoundException1__ As FileNotFoundException
            Return Nothing
        End Try
    End Function

    Private Sub frmEmailProcessor_Load(ByVal sender As Object, ByVal e As EventArgs)
        Me.Timer1.Interval = CInt(Math.Round(60000F * Email_Processor.Properties.Settings.Default.TimerTimeinMinutes))
        'this.Timer1.Enabled = false;
        Me.Timer1.Enabled = True
    End Sub

    Private Function ExecuteScalar(ByVal St As String, ByVal Optional message As Microsoft.Graph.Message = Nothing) As Object
        Dim Conn = New OleDbConnection(LocalAccessConnSt)

        Try

            'LogMessage("Opening ExecuteScalar Access db connection.", message, TraceEventType.Information);
            Conn.Open()
            'LogMessage("ExecuteScalar Access db connection opened.", message, TraceEventType.Information);
            Dim Command = New OleDbCommand(St, Conn)
            Dim j = Command.ExecuteScalar()
            Command.Dispose()
            Return j
        Catch ex As Exception
            LogError("Error ExecuteScalar Access db: " & ex.Message, message, TraceEventType.Error)
        Finally
            Conn.Close()
        End Try

        Return Nothing
    End Function

    Private Function ExecuteNonQuery(ByVal St As String, ByVal Optional message As Microsoft.Graph.Message = Nothing) As Object
        Dim Conn = New OleDbConnection(LocalAccessConnSt)

        Try
            'LogMessage("Opening ExecuteNonQuery Access db connection.", message, TraceEventType.Information);
            Conn.Open()
            'LogMessage("ExecuteNonQuery Access db connection opened.", message, TraceEventType.Information);
            Dim Command = New OleDbCommand(St, Conn)
            Dim j As Integer = Command.ExecuteNonQuery()
            Command.Dispose()
            Return j
        Catch ex As Exception
            LogError("Error ExecuteNonQuery Access db. St: " & St & ", Exception: " & ex.Message, message, TraceEventType.Error)
        Finally
            Conn.Close()
        End Try

        Return Nothing
    End Function

    Private Function ExecuteNonQuerySQLServer(ByVal St As String, ByVal Optional message As Microsoft.Graph.Message = Nothing) As Object
        Dim Conn = New SqlConnection(LocalSQLServerConnSt)

        Try
            'LogMessage("Opening ExecuteNonQuery SQL Server db connection.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
            Conn.Open()
            'LogMessage("ExecuteNonQuery SQL Server db connection opened.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
            Dim Command = New SqlCommand(St, Conn)
            Dim j As Integer = Command.ExecuteNonQuery()
            Command.Dispose()
            Return j
        Catch ex As Exception
            LogError("Error ExecuteNonQuery SQL Server db. St: " & St & ", Exception: " & ex.Message, message, TraceEventType.Error)
        Finally
            Conn.Close()
        End Try

        Return Nothing
    End Function

    Private Function DLookupAccess(ByVal Value As String, ByVal Table As String, ByVal Optional Condition As String = "", ByVal Optional message As Microsoft.Graph.Message = Nothing) As Object
        Dim Conn = New OleDbConnection(LocalAccessConnSt)
        Dim St = "SELECT [" & Value & "] FROM [" & Table & "] " & If(String.IsNullOrEmpty(Condition), "", "WHERE " & Condition).Replace("[[", "[").Replace("]]", "]")

        Try
            'LogMessage("Opening DLookup Access db connection.", message, TraceEventType.Information);
            Conn.Open()
            'LogMessage("DLookup Access db connection opened.", message, TraceEventType.Information);
            Dim Command = New OleDbCommand(St, Conn)
            Dim j = Command.ExecuteScalar()
            Command.Dispose()
            Return j
        Catch ex As Exception
            LogError("Error ExecuteScalar Access db. St: " & St & ", Exception: " & ex.Message, message, TraceEventType.Error)
            Return Nothing
        Finally
            Conn.Close()
        End Try
    End Function

    Private Function DLookupSQLServer(ByVal Value As String, ByVal Table As String, ByVal Optional Condition As String = "", ByVal Optional message As Microsoft.Graph.Message = Nothing) As Object
        Dim Conn = New SqlConnection(LocalSQLServerConnSt)
        Dim St = "SELECT [" & Value & "] FROM [" & Table & "] " & If(String.IsNullOrEmpty(Condition), "", "WHERE " & Condition)

        Try
            'LogMessage("Opening Dlookup SQL Server db connection.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
            Conn.Open()
            'LogMessage("Dlookup SQL Server db connection opened.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
            Dim Command = New SqlCommand(St, Conn)
            ' Dim j = Command.ExecuteScalar(St, Conn)
            Dim j = Command.ExecuteScalar()
            Command.Dispose()
            Return j
        Catch ex As Exception
            LogError("Error dlookup SQL Server db. St: " & St & ", Exception: " & ex.Message, message, TraceEventType.Error)
            Return Nothing
        Finally
            Conn.Close()
        End Try
    End Function

    Private Function DLookupSQLServerRemote(ByVal Value As String, ByVal Table As String, ByVal Optional Condition As String = "", ByVal Optional message As Microsoft.Graph.Message = Nothing) As Object
        Dim Conn = New SqlConnection(RemoteSQLServerConnSt)
        Dim St = "SELECT [" & Value & "] FROM [" & Table & "] " & If(String.IsNullOrEmpty(Condition), "", "WHERE " & Condition)

        Try
            'LogMessage("Opening Dlookup SQL Server db connection.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
            Conn.Open()
            'LogMessage("Dlookup SQL Server db connection opened.", message, TraceEventType.Information, "", false, LogChannels.SQLServer);
            Dim Command = New SqlCommand(St, Conn)
            ' Dim j = Command.ExecuteScalar(St, Conn)
            Dim j = Command.ExecuteScalar()
            Command.Dispose()
            Return j
        Catch ex As Exception
            LogError("Error dlookup SQL Server db. St: " & St & ", Exception: " & ex.Message, message, TraceEventType.Error)
            Return Nothing
        Finally
            Conn.Close()
        End Try
    End Function

    Public Function StripHtmlTags(ByVal html As String) As String

        ' Remove HTML tags.
        Return Regex.Replace(html, "<.*?>", "")
    End Function
End Class
