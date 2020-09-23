VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMail 
   Caption         =   "Send Mail"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboAttachments 
      Height          =   315
      Left            =   2940
      TabIndex        =   5
      Top             =   960
      Width           =   5835
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAttach 
      Caption         =   "Attach File:"
      Height          =   375
      Left            =   1875
      TabIndex        =   4
      Top             =   930
      Width           =   975
   End
   Begin RichTextLib.RichTextBox txtMessage 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4471
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMail.frx":0000
   End
   Begin VB.ComboBox cboTo 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   7950
      TabIndex        =   3
      Top             =   6570
      Width           =   855
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   "Message:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Subject:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const CurrentModule As String = "frmMail"
Private Const m As Long = 10
Private MailSent As Boolean

'Auto fillin stuff
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_FINDSTRING = &H14C
Private Const CB_ERR = (-1)
Private m_bEditFromCode As Boolean
Private MatchingEntryIndex As Long
Private KeyStrokeCount As Long
Private CurrentAttachment As Long

Private Sub cboAttachments_Click()
    CurrentAttachment = cboAttachments.ListIndex
End Sub

Private Sub cboAttachments_DragDrop(Source As Control, X As Single, Y As Single)
    MsgBox "you dropped a file onto the attachments box"
End Sub

Private Sub cboAttachments_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    If KeyCode = vbKeyDelete Then
        'Delete the current attachment
        i = CurrentAttachment
        If i > -1 Then
            cboAttachments.RemoveItem i
            If cboAttachments.ListIndex = -1 And cboAttachments.ListCount > 0 Then
                cboAttachments.ListIndex = 0
            End If
        End If
    End If
End Sub

Private Sub cboTo_Change()
    On Error GoTo Err_Init
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String

    'Prevent processing as a result of changes from code
    If m_bEditFromCode Then
        m_bEditFromCode = False
        Exit Sub
    End If
    With cboTo
        'Lookup list item matching text so far
        strPartial = .Text
        i = SendMessage(.hWnd, CB_FINDSTRING, -1, ByVal strPartial)
        'If match found, append unmatched characters
        If i <> CB_ERR Then
            'Get full text of matching list item
            strTotal = .List(i)
            MatchingEntryIndex = i
            'Compute number of unmatched characters
            j = Len(strTotal) - Len(strPartial)
            '
            If j <> 0 Then
                'Append unmatched characters to string
                m_bEditFromCode = True
                .SelText = Right$(strTotal, j)
                'Select unmatched characters
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    Exit Sub

Err_Init:
    HandleError CurrentModule, "cboTo_Change", Err.Number, Err.Description
End Sub

Private Sub cboTo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Init
    Select Case KeyCode
        Case vbKeyDelete
            m_bEditFromCode = True
        Case vbKeyBack
            m_bEditFromCode = True
    End Select
    KeyStrokeCount = KeyStrokeCount + 1
    Exit Sub

Err_Init:
    HandleError CurrentModule, "cboTo_KeyDown", Err.Number, Err.Description
End Sub

Private Sub cboTo_LostFocus()
    On Error GoTo Err_Init
    Dim s As String
    If StrComp(cboTo.Text, cboTo.List(MatchingEntryIndex), vbTextCompare) = 0 Then
        Exit Sub
    End If
    If KeyStrokeCount < 3 Then
        'They didn't type it - don't add it.
        Exit Sub
    End If
    s = Replace(cboTo.Text, "'", vbQuote)
    Status 2, "Adding new entry to address list"
    If DB.AddTo(s) = True Then
        DB.LoadTo cboTo
        cboTo.Text = s
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "cboTo_LostFocus", Err.Number, Err.Description
End Sub

Private Sub cmdAttach_Click()
    Dim FileToEncode As String, s() As String, i As Long
    FileToEncode = GetFileName("AllFiles|*.*")
    If FileToEncode = "" Then
        MsgBox "Process cancelled!"
        Exit Sub
    End If
    cboAttachments.AddItem FileToEncode
    If cboAttachments.ListIndex = -1 Then
        cboAttachments.ListIndex = 0
    End If
End Sub

Private Function ConvertAttachmentOld(ByVal FileToEncode As String, ByVal TheBoundary As String) As String
    Dim dummy As Long, i As Long
    Dim EncodedFile As String
    Dim EncodingApp As String
    EncodingApp = App.Title
    SetEncodingApplication EncodingApp
    EncodedFile = App.Path & "\OUTATTACH.MIM"
    dummy = EncodeFile(FileToEncode, EncodedFile, TheBoundary, MIME_TYPE, 0)
    FinishAttachments EncodedFile
    ConvertAttachmentOld = EncodedFile
End Function

Private Sub cmdSend_Click()
    On Error GoTo Err_Init
    Dim m As typeMail, i As Long, TheBoundary As String
    Dim EncodingApp As String, EncodedFile As String
    Dim dummy As Long, FileNo As Integer
    Screen.MousePointer = vbHourglass
    EncodingApp = App.Title
    SetEncodingApplication EncodingApp
    EncodedFile = App.Path & "\OUTATTACH.MIM"
    TheBoundary = "--Bound"
    With m
        m.Subject = txtSubject.Text
        m.To = cboTo.Text
        If cboAttachments.ListCount > 0 Then
            'Erase old file
            FileNo = FreeFile
            Open EncodedFile For Output As #FileNo
            Close #FileNo
            'Create new file to hold attachments
            m.Boundary = TheBoundary
            m.Body = "--" & m.Boundary & vbCrLf & _
              "Content-Type: text/plain; charset=us-ascii" & vbCrLf & _
              "Content-Transfer-Encoding: quoted-printable" & vbCrLf & _
              vbCrLf & LineBreak(txtMessage.Text, 80) & vbCrLf & vbCrLf & "--" & m.Boundary
            For i = 0 To cboAttachments.ListCount - 1
                dummy = EncodeFile(cboAttachments.List(i), EncodedFile, TheBoundary, MIME_TYPE, 1)
            Next i
            FinishAttachments EncodedFile
            'Tack new file onto end of message body.
            m.Body = m.Body & LoadFile(EncodedFile)
        Else
            If txtMessage.Tag = "Preformatted" Then
                'don't format
                m.Body = txtMessage.Text
            Else
                m.Body = LineBreak(txtMessage.Text, 80)
            End If
            If InStr(1, m.Body, "----Bound" & vbCrLf, vbTextCompare) = 1 Then
                m.Boundary = TheBoundary
            End If
        End If
        DB.StoreDraft m
        MailSent = True
        Unload Me
    End With
    Screen.MousePointer = vbDefault
    Exit Sub

Err_Init:
    HandleError CurrentModule, "cmdSend_Click", Err.Number, Err.Description
End Sub

Private Function LoadFile(ByVal FileName As String) As String
    On Error GoTo Err_Init
    Dim FileNo As Integer, l As Long
    FileNo = FreeFile
    l = FileLen(FileName)
    Open FileName For Input As #FileNo
    LoadFile = Input(l, #FileNo)
    Close #FileNo
    l = InStr(3, LoadFile, vbCrLf, vbTextCompare)
    LoadFile = Right$(LoadFile, Len(LoadFile) - l + 1)
    Exit Function

Err_Init:
    HandleError CurrentModule, "LoadFile", Err.Number, Err.Description
End Function

Private Sub Form_Load()
    On Error GoTo Err_Init
    cboAttachments.Clear
    txtMessage.Tag = "" 'Clear any 'preformatted' messages lurking
    'CommonDialog1.Flags = cdlOFNAllowMultiselect
    MailSent = False
    LoadTo
    KeyStrokeCount = 0
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_Load", Err.Number, Err.Description
End Sub

Private Sub Form_Resize()
    On Error GoTo Err_Init
    Dim w As Long, h As Long
    If frmMail.WindowState = vbMinimized Then
        Exit Sub
    End If
    w = ScaleWidth
    h = ScaleHeight
    cmdSend.Left = w - cmdSend.Width - m
    cmdSend.Top = h - m - cmdSend.Height
    cboTo.Width = w - m - cboTo.Left
    txtSubject.Width = cboTo.Width
    cboAttachments.Width = w - m - cboAttachments.Left
    txtMessage.Move m, txtMessage.Top, w - txtMessage.Left - m, h - txtMessage.Top - m * 2 - cmdSend.Height
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_Resize", Err.Number, Err.Description
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim RetVal As Long
    If Len(txtMessage.Text) > 0 And MailSent = False Then
        RetVal = MsgBox("Cancel message?", vbOKCancel)
        If RetVal = vbOK Then
            'ok
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub txtMessage_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Init
    Dim CtrlDown As Boolean
    CtrlDown = (Shift And vbCtrlMask) > 0
    If CtrlDown = True And KeyCode = 13 Then
        cmdSend_Click
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "txtMessage_KeyUp", Err.Number, Err.Description
End Sub

Private Sub LoadTo()
    On Error GoTo Err_Init
    DB.LoadTo frmMail.cboTo
    Exit Sub

Err_Init:
    HandleError CurrentModule, "LoadTo", Err.Number, Err.Description
End Sub
