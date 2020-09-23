VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmMain 
   Caption         =   "EZMailer"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   602
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "F:\Code\EZMailer\mail.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Mail"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2820
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   465
      Top             =   4140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   7155
      TabIndex        =   7
      Top             =   5070
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer tmrSize 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2085
      Top             =   3840
   End
   Begin VB.Timer tmrClock 
      Interval        =   1000
      Left            =   2970
      Top             =   3900
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1455
      Top             =   3795
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00C00000&
      Height          =   2430
      Index           =   2
      Left            =   2580
      ScaleHeight     =   158
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   302
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2460
      Width           =   4590
      Begin RichTextLib.RichTextBox txtMail 
         Height          =   1740
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   3069
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":0000
      End
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H0000C000&
      Height          =   1290
      Index           =   1
      Left            =   1800
      ScaleHeight     =   82
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   477
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   7215
      Begin TrueDBGrid60.TDBGrid tdbMail 
         Bindings        =   "frmMain.frx":00C9
         DragIcon        =   "frmMain.frx":00DD
         Height          =   1215
         Left            =   0
         OleObjectBlob   =   "frmMain.frx":051F
         TabIndex        =   1
         Top             =   15
         Width           =   7215
      End
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H000000FF&
      Height          =   1860
      Index           =   0
      Left            =   120
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   1515
      Begin VB.ListBox lstFoldersNames 
         Height          =   540
         IntegralHeight  =   0   'False
         Left            =   -90
         TabIndex        =   9
         Top             =   1290
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.ListBox lstFoldersData 
         Height          =   540
         IntegralHeight  =   0   'False
         Left            =   15
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.ListBox lstFolders 
         Height          =   540
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2100
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5010
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9816
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2910
            MinWidth        =   2910
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   1305
      Top             =   4470
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileReply 
         Caption         =   "&Reply"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuForward 
         Caption         =   "&Forward"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsGetMail 
         Caption         =   "&Get Mail"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuOptionsSendMail 
         Caption         =   "&Send Mail"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save A&ttachments"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsAccount 
         Caption         =   "Edit Account &Info"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFileRouting 
         Caption         =   "Edit R&ules"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuAddressBook 
         Caption         =   "Edit Address &Book"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileShortcut 
         Caption         =   "Show S&hortcuts"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuProgrammer 
         Caption         =   "&Programmer"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const CurrentModule As String = "frmMain"
Private Sizing As String
Private Const Margin As Long = 2
Private mX As Long
Private mY As Long
Private Const GWL_STYLE = -16&
Private Const TVM_SETBKCOLOR = 4381&
Private Const TVM_GETBKCOLOR = 4383&
Private Const TVS_HASLINES = 2&
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Startup As Boolean

Private Sub Form_Load()
    On Error GoTo Err_Init
    Dim i As Long

    'Set the initial parameters
    Status 2, "Welcome to EZMailer!"
    Status 3, "User: " & User.Name
    ProgressBar1.Visible = False
    frmMain.pic(0).MousePointer = vbArrow
    frmMain.pic(1).MousePointer = vbArrow
    frmMain.pic(2).MousePointer = vbArrow
    frmMain.StatusBar1.MousePointer = vbArrow
    
    'Set the colors
    lstFolders.AddItem "Inbox"
    lstFolders.AddItem "Outbox"
    lstFolders.AddItem "Sent Items"
    lstFolders.BackColor = Settings.Color(1)
    tdbMail.BackColor = Settings.Color(2)
    tdbMail.DeadAreaBackColor = Settings.Color(2)
    txtMail.BackColor = Settings.Color(3)
    
    'Load the folders and messages
    DB.LoadFolders
    'Navigate to the correct folder
    frmMain.lstFolders.ListIndex = 0
    DB.LoadMessages
    'Apply the filtering rules
    DB.LoadRules
    
    Show
    lstFolders.SetFocus
    
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_Load", Err.Number, Err.Description
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Init
    If X < pic(1).Left Then
        Sizing = "v"
        tmrSize.Enabled = True
    ElseIf Y < StatusBar1.Top Then
        Sizing = "h"
        tmrSize.Enabled = True
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_MouseDown", Err.Number, Err.Description
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Init
    If Sizing = "" Then
        If X < pic(1).Left Then
            frmMain.MousePointer = vbSizeWE
        ElseIf Y < StatusBar1.Top Then
            frmMain.MousePointer = vbSizeNS
        Else
            frmMain.MousePointer = vbArrow
        End If
    End If
    mX = X
    mY = Y
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_MouseMove", Err.Number, Err.Description
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Init
    Sizing = ""
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_MouseUp", Err.Number, Err.Description
End Sub

Private Sub mnuAddressBook_Click()
    On Error GoTo Err_Init
    frmAddresses.Show vbModal
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuAddressBook_Click", Err.Number, Err.Description
End Sub

'--------------------------------------------------------
'Menu routines
'--------------------------------------------------------

Private Sub mnuFileExit_Click()
    On Error GoTo Err_Init
    ShutDown
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuFileExit_Click", Err.Number, Err.Description
End Sub

'--------------------------------------------------------
'Form sizing routines
'--------------------------------------------------------

Private Sub Form_Resize()
    On Error Resume Next
    Dim w As Long, h As Long
    If ScaleWidth < 100 Then Width = 100 * Screen.TwipsPerPixelX
    If ScaleHeight < 100 Then Height = 100 * Screen.TwipsPerPixelY
    w = ScaleWidth
    h = ScaleHeight - StatusBar1.Height '- Toolbar1.Height
    'Resize the parent controls.
    pic(0).Move 0, 0, (w * Settings.VerticalDivider) - Margin, h
    pic(1).Move pic(0).Width + Margin, pic(0).Top, w - pic(0).Width - Margin, h * Settings.HorizontalDivider - Margin
    pic(2).Move pic(1).Left, pic(1).Height + pic(1).Top + Margin, pic(1).Width, h - pic(1).Height - Margin
    'Move the progress bar control
    ProgressBar1.Move w - ProgressBar1.Width - 20, pic(2).Top + pic(2).Height + 4
End Sub

Private Sub mnuFileNew_Click()
    On Error GoTo Err_Init
    frmMail.Show
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuFileNew_Click", Err.Number, Err.Description
End Sub

Private Sub mnuFileReply_Click()
    On Error GoTo Err_Init
    DB.Reply
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuFileReply_Click", Err.Number, Err.Description
End Sub

Private Sub mnuFileRouting_Click()
    On Error GoTo Err_Init
    frmRules.Show vbModal
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuFileRouting_Click", Err.Number, Err.Description
End Sub

Private Sub mnuFileSave_Click()
    On Error GoTo Err_Init
    DB.SaveAttachments
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuFileSave_Click", Err.Number, Err.Description
End Sub

Private Sub mnuFileShortcut_Click()
    On Error GoTo Err_Init
    MsgBox "Shortcuts:" & vbCrLf & _
     vbCrLf & _
    "<ENTER> - Reply to current message" & vbCrLf & _
    "<DEL> - Delete current message" & vbCrLf & _
    "<ALT-CLICK> - Change color of current area" & vbCrLf & _
    "<CTRL-CLICK> - Display message + headers with default text processor" & vbCrLf & _
    "<CTRL-G> - Get Mail" & vbCrLf & _
    "<CTRL-S> - Send Mail" & vbCrLf & _
    "<CTRL-P> - Preferences" & vbCrLf & _
    "<CTRL-T> - Save Attachments" & vbCrLf & _
    "<CTRL-H> - Display Shortcuts" & vbCrLf & _
    "<CTRL-Q> - Quit"
    
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuFileShortcut_Click", Err.Number, Err.Description
End Sub

Private Sub mnuForward_Click()
    On Error GoTo Err_Init
    DB.Forward
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuForward_Click", Err.Number, Err.Description
End Sub

Private Sub mnuOptionsAccount_Click()
    On Error GoTo Err_Init
    frmUser.Show vbModal
    DB.EditUser
    Status 3, "User: " & User.Name
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuOptionsAccount_Click", Err.Number, Err.Description
End Sub

Private Sub mnuOptionsGetMail_Click()
    On Error GoTo Err_Init
    If Receiving = True Then
        Exit Sub
    End If
    Receiving = True
    Screen.MousePointer = vbHourglass
    frmMain.Enabled = False
    CN.GetMail
    DB.ApplyRules
    Receiving = False
    frmMain.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuOptionsGetMail_Click", Err.Number, Err.Description
End Sub

Private Sub mnuOptionsSendMail_Click()
    On Error GoTo Err_Init
    If Receiving = True Then
        Exit Sub
    End If
    Receiving = True
    Screen.MousePointer = vbHourglass
    frmMain.Enabled = False
    DB.SendMail
    Receiving = False
    frmMain.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuOptionsSendMail_Click", Err.Number, Err.Description
End Sub

Private Sub mnuProgrammer_Click()
    On Error GoTo Err_Init
    DB.Programmer
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mnuProgrammer_Click", Err.Number, Err.Description
End Sub

Private Sub pic_Resize(Index As Integer)
    On Error GoTo Err_Init
    If Index = 0 Then
        lstFolders.Move -2, -2, pic(0).Width, pic(0).Height
    ElseIf Index = 1 Then
        tdbMail.Move -2, -2, pic(1).Width, pic(1).Height
    ElseIf Index = 2 Then
        txtMail.Move -4, -4, pic(2).Width + Margin + 1, pic(2).Height + Margin '- 15
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "pic_Resize", Err.Number, Err.Description
End Sub

Private Sub tdbMail_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Init
    If KeyCode = 13 Then
        DB.Reply
    ElseIf KeyCode = vbKeyDelete Then
        DB.DeleteMessage
        Status 2, "Message deleted."
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "tdbMail_KeyUp", Err.Number, Err.Description
End Sub

Private Sub tmrClock_Timer()
    On Error GoTo Err_Init
    Status 1, Format(Now, "MM/DD/YY HH:MM:SS")
    Exit Sub

Err_Init:
    HandleError CurrentModule, "tmrClock_Timer", Err.Number, Err.Description
End Sub

Private Sub tmrSize_Timer()
    On Error GoTo Err_Init
    If Sizing > "" Then
        If Sizing = "v" Then
            'move ns
            Settings.VerticalDivider = mX / ScaleWidth
        Else
            'move ew
            Settings.HorizontalDivider = mY / ScaleHeight
        End If
        Form_Resize
    Else
        tmrSize.Enabled = False
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "tmrSize_Timer", Err.Number, Err.Description
End Sub

'--------------------------------------------------------
'Color setting routines
'--------------------------------------------------------

Private Sub tdbMail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Init

   Dim c As Long
   Dim ShiftDown, AltDown, CtrlDown
   ShiftDown = (Shift And vbShiftMask) > 0
   AltDown = (Shift And vbAltMask) > 0
   CtrlDown = (Shift And vbCtrlMask) > 0
       
    If AltDown = True And X > 270 Then
        c = GetColor
        If c >= 0 Then
            Settings.Color(2) = c
            tdbMail.BackColor = c
            tdbMail.DeadAreaBackColor = c
        End If
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "tdbMail_MouseDown", Err.Number, Err.Description
End Sub

Private Sub lstFolders_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Init
    Dim c As Long
    Dim ShiftDown, AltDown, CtrlDown
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    If AltDown = True Then
        c = GetColor
        If c >= 0 Then
            Settings.Color(1) = c
            lstFolders.BackColor = c
        End If
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "lstFolders_MouseDown", Err.Number, Err.Description
End Sub

Private Sub txtMail_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Init
    If KeyCode = 13 Then DB.Reply
    Exit Sub

Err_Init:
    HandleError CurrentModule, "txtMail_KeyUp", Err.Number, Err.Description
End Sub

Private Sub txtMail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Init
    Dim c As Long
    Dim ShiftDown, AltDown, CtrlDown
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    If AltDown = True Then
        c = GetColor
        If c >= 0 Then
            Settings.Color(3) = c
            txtMail.BackColor = c
        End If
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "txtMail_MouseDown", Err.Number, Err.Description
End Sub

Private Function GetColor() As Long
    On Error GoTo Err_Init
    With CommonDialog1
        .CancelError = True
        .ShowColor
        GetColor = .Color
    End With
    Exit Function
Err_Init:
    If Err.Number = 32755 Then
        'user cancelled
    Else
        MsgBox Err.Number & " - " & Err.Description
    End If
    GetColor = -1
End Function

Private Sub txtMail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Init
    Dim CtrlDown As Boolean
    CtrlDown = (Shift And vbCtrlMask) > 0
    If CtrlDown Then
        'save as a text
        DB.ExportMail
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "txtMail_MouseUp", Err.Number, Err.Description
End Sub
