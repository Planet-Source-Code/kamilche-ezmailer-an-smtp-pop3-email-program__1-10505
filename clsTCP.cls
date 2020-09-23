VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const CurrentModule As String = "clsTCP"
Private WithEvents mTCP As Winsock
Attribute mTCP.VB_VarHelpID = -1
Private WithEvents mTCP2 As Winsock
Attribute mTCP2.VB_VarHelpID = -1
Private WithEvents prgBar As ProgressBar
Attribute prgBar.VB_VarHelpID = -1
Private Response As String
Private Response2 As String
Private ErrResponse As String
Private ErrResponse2 As String
Private TCPError As Boolean
Private TCPError2 As Boolean
Private GetMsg As Boolean
Private Buffer As String
Private nb As Long
Private TotalBytes As Long
Private DoneSending As Boolean
Private TotalBytesSent As Long

Private Sub Class_Initialize()
    Set mTCP = frmMain.Winsock1
    On Error GoTo Err_Init
    Set mTCP2 = frmMain.Winsock2
    Set prgBar = frmMain.ProgressBar1
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Class_Initialize", Err.Number, Err.Description
End Sub

Private Sub Class_Terminate()
    Set mTCP = Nothing
    On Error GoTo Err_Init
    Set mTCP2 = Nothing
    Set prgBar = Nothing
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Class_Terminate", Err.Number, Err.Description
End Sub

Friend Function SendEmail(ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String, Optional CCName As String, Optional Boundary As String) As Boolean
'Sends email via winsock FTP commands. Can't handle IMAP, has to be SMTP server.
    Dim DateNow As String, First As String, Second As String, Third As String
    Dim Fourth As String, Fifth As String, Sixth As String, Seventh As String
    Dim Eighth As String, Ninth As String, CC As String
    
    SendEmail = True
    
    TCPError = False
    
    On Error GoTo Err_Init
    Screen.MousePointer = vbHourglass
    Status 2, "Sending email to " & ToName & ", please wait..."
    
    prgBar.Visible = True
    prgBar.Value = 0
    
    mTCP.LocalPort = 0
 
    If mTCP.State = sckClosed Then ' Check to see if socket is closed
        DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
        First = "mail from:" + Chr(32) + User.Email + vbCrLf ' Get who's sending E-CN address
        Second = "rcpt to:" + Chr(32) + ToEmailAddress + vbCrLf ' Get who CN is going to
        If Len(CCName) > 0 Then
            CC = "rcpt to:" + Chr(32) + CCName + vbCrLf
        End If
        Third = "Date:" + Chr(32) + DateNow + vbCrLf ' Date when being sent
        Fourth = "From: " & vbQuote & User.Name & vbQuote & " <" & User.Email & ">" & vbCrLf  ' Who's Sending
        Fifth = "To: " & vbQuote & ToName & vbQuote & " <" & ToEmailAddress & ">" & vbCrLf   ' Who it going to
        Sixth = "Subject:" + Chr(32) + EmailSubject + vbCrLf ' Subject of E-CN
        Seventh = EmailBodyOfMessage + vbCrLf ' E-CN message body
        Ninth = "X-Mailer: EZMailer" + vbCrLf ' What program sent the e-CN, customize this
        If Len(Boundary) > 0 Then
            'there's an attachment.
            Seventh = "This is a multi-part message in MIME format." & vbCrLf & Seventh
            Ninth = Ninth & "MIME-Version: 1.0" & vbCrLf & "Content-Type: multipart/mixed; boundary=" & vbQuote & Boundary & vbQuote & vbCrLf
        End If
        Eighth = Fourth + Third + Ninth + Fifth + Sixth  ' Combine for proper SMTP sending
     
        mTCP.RemoteHost = User.SMTP
        mTCP.RemotePort = 25
        mTCP.Protocol = sckTCPProtocol ' Set protocol for sending     mTCP.RemoteHost = MailServerName ' Set the server address     mTCP.RemotePort = 25 ' Set the SMTP Port
        mTCP.Connect ' Start connection
     
        WaitFor ("220")
        If TCPError = True Then GoTo TCPError
     
        Status 2, "Connecting...."
     
        mTCP.SendData ("HELO " & mTCP.LocalHostName & vbCrLf)
        
        WaitFor ("250")
        If TCPError = True Then GoTo TCPError
     
        Status 2, "Connected"
        
        mTCP.SendData (First)
     
        Status 2, "Sending Email"
    
        WaitFor ("250")
        If TCPError = True Then GoTo TCPError
     
        mTCP.SendData (Second)
     
        WaitFor ("250")
        If TCPError = True Then GoTo TCPError
     
        If Len(CC) > 0 Then
            mTCP.SendData (CC)
            WaitFor ("250")
            If TCPError = True Then GoTo TCPError
        End If
            
        mTCP.SendData ("data" + vbCrLf)
     
        WaitFor ("354")
        If TCPError = True Then GoTo TCPError
     
        mTCP.SendData (Eighth + vbCrLf)
        
        'Send the bulk of the message.
        DoneSending = False
        TotalBytesSent = 0
        TotalBytes = Len(Seventh)
        mTCP.SendData (Seventh + vbCrLf)
        WaitForDoneSending
        
        'Send the 'end of message' indicator.
        mTCP.SendData ("." + vbCrLf)
     
        WaitFor ("250")
     
        mTCP.SendData ("quit" + vbCrLf)
     
        Status 2, "Disconnecting"
        
        WaitFor ("221")
        If TCPError = True Then GoTo TCPError
     
        mTCP.Close
        
        Status 2, "Email send is complete."
        prgBar.Visible = False
        
        Screen.MousePointer = vbDefault
    End If
    
    Exit Function
    
TCPError:
    SendEmail = False
    mTCP.Close
    Screen.MousePointer = vbDefault
    MsgBox ErrResponse & vbCrLf, vbCritical, "TCP Error"
    prgBar.Visible = False
    Exit Function
    
Err_Init:
    SendEmail = False
    mTCP.Close
    If Err.Number = 10049 Then
        MsgBox "Please connect to the internet first!"
    Else
        HandleError CurrentModule, "SendEmail", Err.Number, Err.Description
    End If
    prgBar.Visible = False
    
End Function

Private Sub WaitFor(ResponseCode As String)
'Called by 'sendmail' to wait for the correct response code.
    Dim Start As Long, Tmr As Long
    On Error GoTo Err_Init
    Start = Timer
    While Len(Response) = 0
        Tmr = Timer - Start
        DoEvents
        If Tmr > 20 Then
            TCPError = True
            ErrResponse = "SMTP service error, timed out while waiting for response"
            Status 2, ErrResponse
            GoTo Done
        End If
    Wend
    If Len(Response) > 2 Then
        If Left$(Response, 3) <> ResponseCode Then
            TCPError = True
            ErrResponse = "Code should have been: " + ResponseCode & vbCrLf & "Code recieved: " + Response
            Status 2, ErrResponse
        End If
    End If
Done:
    Response = ""
    Exit Sub
Err_Init:
    HandleError CurrentModule, "WaitFor", Err.Number, Err.Description
End Sub

Private Function WaitForDoneSending()
'Called by 'sendmail' to wait until the message body is sent.
'Uses progress and completion codes from mTCP_SendProgress and mTCP_SendComplete
    Dim Start As Long, Tmr As Long, LastB As Long
    On Error GoTo Err_Init
    Start = Timer
    LastB = 0
    Do
        DoEvents
        If DoneSending = True Then
            Exit Do
        End If
        If TotalBytesSent > LastB Then
            LastB = TotalBytesSent
            Start = Timer 'reset the timer
        End If
        Tmr = Timer - Start
        If Tmr > 60 Then
            TCPError = True
            ErrResponse = "SMTP service error, timed out while sending the message."
            Status 2, ErrResponse
            GoTo Done
        End If
    Loop
Done:
    Response = ""
    Exit Function
Err_Init:
    HandleError CurrentModule, "WaitForDoneSending", Err.Number, Err.Description
End Function

Private Function WaitForChar(ByVal Which As String) As String
'Called by 'getmail' to wait for the correct response code.
    Dim Start As Long, Tmr As Long
    On Error GoTo Err_Init
    Start = Timer
    While Len(Response2) = 0
        Tmr = Timer - Start
        DoEvents
        If Tmr > 20 Then
            TCPError2 = True
            ErrResponse2 = "POP3 service error, timed out while waiting for response2"
            Status 2, ErrResponse2
            GoTo Done
        End If
    Wend
    If Left$(Response2, 1) <> Which Then
        TCPError2 = True
        ErrResponse2 = InStr(1, Response2, Which) & " Code should have been: +. Code recieved: " + Response2
        Status 2, ErrResponse2
    End If
Done:
    frmMain.txtMail.SelStart = Len(frmMain.txtMail.Text)
    frmMain.txtMail.SelText = Response2
    frmMain.txtMail.SelStart = Len(frmMain.txtMail.Text)
    WaitForChar = Response2
    Response2 = ""
    Exit Function
Err_Init:
    HandleError CurrentModule, "WaitForChar", Err.Number, Err.Description
End Function

Private Function WaitForMsg() As String
'Called by 'getmail' to wait for the complete message.
    Dim Start As Long, Tmr As Long
    On Error GoTo Err_Init
    Start = Timer
    Do
        DoEvents
        If Right$(Buffer, 5) = vbCrLf & "." & vbCrLf Then
            Exit Do
        End If
    Loop
    WaitForMsg = Buffer
    Buffer = ""
Done:
    Exit Function
Err_Init:
    HandleError CurrentModule, "WaitForMsg", Err.Number, Err.Description
End Function

Private Sub mTCP_DataArrival(ByVal bytesTotal As Long)
'Receives incoming commands from the TCP control
    On Error GoTo Err_Init
    mTCP.GetData Response
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mTCP_DataArrival", Err.Number, Err.Description
End Sub

Private Sub mTCP_SendComplete()
    DoneSending = True
End Sub

Private Sub mTCP_SendProgress(ByVal BytesSent As Long, ByVal bytesRemaining As Long)
    Dim i As Long
    TotalBytesSent = TotalBytesSent + BytesSent
    If TotalBytes > 0 Then
        i = Int(100 * TotalBytesSent / TotalBytes)
        If i > 100 Then i = 100
        prgBar.Value = i
    End If
End Sub

Private Sub mTCP2_Connect()
    Status 2, "Connected to " & mTCP2.RemoteHost
    On Error GoTo Err_Init
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mTCP2_Connect", Err.Number, Err.Description
End Sub

Private Sub mTCP2_DataArrival(ByVal bytesTotal As Long)
'Receives incoming commands from the TCP control
    On Error GoTo Err_Init
    Dim s As String, i As Long
    If GetMsg = True Then
        nb = nb + mTCP2.BytesReceived
        mTCP2.GetData s
        Buffer = Buffer & s
        i = Int(100 * nb / TotalBytes)
        If i > 100 Then i = 100
        prgBar.Value = i
    Else
        mTCP2.GetData Response2
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mTCP2_DataArrival", Err.Number, Err.Description
End Sub

Friend Sub GetMail()

    Dim s() As String, i As Long, max As Long, Stat As String
    Dim msg As String
    On Error GoTo Err_Init
    
    frmMain.txtMail.Text = ""

    Status 2, "Retrieving mail."
    With mTCP2
        .RemoteHost = User.POP3
        .RemotePort = 110
        .Protocol = sckTCPProtocol
        .Connect
    End With
    WaitForChar "+"
    If TCPError2 = True Then GoTo TCPError
    
    mTCP2.SendData "USER " & LeftPart(User.Email) & vbCrLf
    WaitForChar "+"
    If TCPError2 = True Then GoTo TCPError
    
    mTCP2.SendData "PASS " & User.Password & vbCrLf
    WaitForChar "+"
    If TCPError2 = True Then GoTo TCPError
    
    mTCP2.SendData "STAT" & vbCrLf
    Stat = WaitForChar("+")
    If TCPError2 = True Then GoTo TCPError
    
    s = Split(Stat, " ")
    max = CLng(s(1))
    
    TotalBytes = CLng(s(2))
    nb = 0
    
    prgBar.Visible = True
    prgBar.Value = 0
    
    GetMsg = True
    
    For i = 1 To max
        Buffer = ""
        mTCP2.SendData "RETR " & i & vbCrLf
        msg = WaitForMsg
        If TCPError2 = True Then GoTo TCPError
        If DB.SaveMail(msg) Then
            'delete the message
            GetMsg = False
            mTCP2.SendData "DELE " & i & vbCrLf
            WaitForChar "+"
            If TCPError2 = True Then GoTo TCPError
            GetMsg = True
        Else
            MsgBox "Error in saving the mail, so I left it on the server. Please contact Kamilche.", vbInformation
            Exit For
        End If
    Next i
    
    GetMsg = False
    
    mTCP2.SendData "QUIT" & vbCrLf
    WaitForChar "+"
    If TCPError2 = True Then GoTo TCPError
    
    mTCP2.Close
    Screen.MousePointer = vbDefault
    Status 2, max & " messages received."
    prgBar.Visible = False
    
    Exit Sub
    
TCPError:
    GetMsg = False
    mTCP2.Close
    Screen.MousePointer = vbDefault
    MsgBox ErrResponse2 & vbCrLf, vbCritical, "TCP Error"
    Exit Sub
    
Err_Init:
    GetMsg = False
    mTCP2.Close
    If Err.Number = 10049 Then
        MsgBox "Please connect to the internet first!"
    Else
        HandleError CurrentModule, "GetMail", Err.Number, Err.Description
    End If
End Sub

Private Function LeftPart(ByVal s As String) As String
    On Error GoTo Err_Init
    Dim c As Long
    c = InStr(1, s, "@")
    LeftPart = Left$(s, c - 1)
    Exit Function

Err_Init:
    HandleError CurrentModule, "LeftPart", Err.Number, Err.Description
End Function
