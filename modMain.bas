Attribute VB_Name = "modMain"
Option Explicit
Const CurrentModule As String = "modMain"

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Public Const PROMPT_NONE As Long = 1
Public Const MIME_TYPE As Long = 2

Public Settings As clsSettings
Public CN As clsTCP
Public DB As clsDatabase
Public Const vbQuote = """"
Public Receiving As Boolean

Public Type typeUser
    ID As Long
    Name As String
    Email As String
    Password As String
    SMTP As String
    POP3 As String
End Type

Public Type typeMail
    ID As Long
    Folder As Long
    From As String
    To As String
    Subject As String
    Date As Date
    Read As Boolean
    Header As String
    Body As String
    Boundary As String
    Attachments As Long
End Type

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Type typeRule
    PartID As Long
    FindPhrase As String
    FolderID As Long
End Type

Public Rules() As typeRule

Public User As typeUser

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function GetNumFilesToDecode Lib "DECENC32.DLL" (ByVal strInFile As String) As Long
Public Declare Function GetEncodedFile Lib "DECENC32.DLL" (ByVal strOutFile As String, ByVal nIndex As Long) As Long
Public Declare Function DecodeFile Lib "DECENC32.DLL" (ByVal strInFile As String, ByVal strOutFile As String, ByVal nPrompts As Long) As Long
Public Declare Sub SetEncodingApplication Lib "DECENC32.DLL" (ByVal strInFile As String)
Public Declare Function EncodeFile Lib "DECENC32.DLL" (ByVal SourceFile As String, ByVal EncodedFile As String, ByVal strBoundary As String, ByVal CodeOption As Long, ByVal xAppend As Long) As Long
Public Declare Sub FinishAttachments Lib "DECENC32.DLL" (ByVal strFileOut As String)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Sub Main()
    On Error GoTo Err_Init
    Set Settings = New clsSettings
    Set DB = New clsDatabase
    Set CN = New clsTCP
    frmLogon.Show vbModal
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Main", Err.Number, Err.Description
End Sub

Public Sub ShutDown()
    On Error GoTo Err_Init
    Set CN = Nothing
    Set DB = Nothing
    Set Settings = Nothing
    Unload frmMain
    End
    Exit Sub

Err_Init:
    HandleError CurrentModule, "ShutDown", Err.Number, Err.Description
End Sub
 
Public Sub Status(ByVal i As Long, ByVal s As String)
    On Error GoTo Err_Init
    frmMain.StatusBar1.Panels(i).Text = s
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Status", Err.Number, Err.Description
End Sub

Public Sub HandleError(ByVal TheMod As String, ByVal TheSub As String, ByVal ErrNo As Long, ByVal ErrDescription As String)
    MsgBox "Module:   " & TheMod & vbCrLf & _
           "Function: " & TheSub & vbCrLf & _
           "Error #:    " & ErrNo & vbCrLf & vbCrLf & _
    ErrDescription, _
    vbCritical, "Error"
End Sub

Public Sub SaveFile(ByVal s As String, ByVal FileName As String)
    On Error GoTo Err_Init
    Dim FileNo As Integer
    FileNo = FreeFile
    Open FileName For Output As #FileNo
    Print #FileNo, s
    Close #FileNo
    Exit Sub

Err_Init:
    HandleError CurrentModule, "SaveFile", Err.Number, Err.Description
End Sub

Public Sub OpenFile(ByVal FileName As String)
    On Error GoTo Err_Init
    Call ShellExecute(0&, vbNullString, FileName, vbNullString, vbNullString, vbNormalFocus)
    Exit Sub

Err_Init:
    HandleError CurrentModule, "OpenFile", Err.Number, Err.Description
End Sub

Public Function GetFolderName() As String
    On Error GoTo Err_Init
    'Opens a Treeview control that displays
    '     the directories in a computer
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = "This is the title"


    With tBrowseInfo
        .hWndOwner = 0 'Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)


    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    End If
    
    GetFolderName = sBuffer
    Exit Function

Err_Init:
    HandleError CurrentModule, "GetFolderName", Err.Number, Err.Description
End Function

Public Function GetFileName(ByVal Filter As String) As String
    On Error GoTo Err_Init
    With frmMail
        .CommonDialog1.Filter = "AllFiles|*.*"
        '.CommonDialog1.
        .CommonDialog1.CancelError = True
        .CommonDialog1.ShowOpen
        GetFileName = .CommonDialog1.FileName
    End With
    Exit Function

Err_Init:
    If Err.Number = 32755 Then
        'user cancelled
    Else
        HandleError CurrentModule, "GetFileName", Err.Number, Err.Description
    End If
End Function

Public Function LeftPart(ByVal s As String) As String
    On Error GoTo Err_Init
    Dim c As Long
    c = InStr(1, s, "@")
    If c = 0 Then
        LeftPart = s
    Else
        LeftPart = Left$(s, c - 1)
    End If
    Exit Function

Err_Init:
    HandleError CurrentModule, "LeftPart", Err.Number, Err.Description
End Function

Public Function RightPart(ByVal s As String) As String
    On Error GoTo Err_Init
    Dim c As Long
    c = InStr(1, s, "@")
    If c = 0 Then
        RightPart = s
    Else
        RightPart = Right$(s, Len(s) - c)
    End If
    Exit Function

Err_Init:
    HandleError CurrentModule, "RightPart", Err.Number, Err.Description
End Function

Public Function LineBreak(ByVal s As String, ByVal LineLength As Long) As String
'This routine inserts line breaks at the desired column.
'Note that it will only break words up on spaces - if a line is too long
'AND has no spaces, it won't be broken.

    Dim c As Long, t As Long, LastLineBreak As Long, LastSpace As Long
    
    On Error GoTo Err_Init
    
    'Insert line breaks
    Do
        c = c + 1 'run of characters without spaces
        t = t + 1 'all characters
        If t > Len(s) Then
            Exit Do
        End If
        'Grab the last line break and space characters.
        If Mid$(s, t, 2) = vbCrLf Then
            LastLineBreak = t
            c = 0
        ElseIf Mid$(s, t, 1) = " " Then
            LastSpace = t
        End If
        'Is the line too long?
        If c > LineLength Then
            'the line's too long
            If LastSpace > (t - LineLength) Then
                'break at the last space found on the prior line
                Mid(s, LastSpace, 1) = Chr$(1)
            Else
                'don't break
            End If
            c = 0
        End If
    Loop
    
    'Replace all chr$(1)'s with vbcrlf's
    Do While InStr(1, s, Chr$(1), vbTextCompare) > 0
        s = Replace(s, Chr$(1), vbCrLf)
    Loop
    
    'Assign to function and return.
    LineBreak = s
    Exit Function
    
Err_Init:
    HandleError CurrentModule, "LineBreak", Err.Number, Err.Description
    Resume Next
End Function


