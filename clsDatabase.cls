VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const CurrentModule As String = "clsDatabase"
Private rsTemp As Recordset
Private rsTemp2 As Recordset
Private DB As Database
Private SQL As String
Private WithEvents mLstFolders As ListBox
Attribute mLstFolders.VB_VarHelpID = -1
Private mLstFoldersData As ListBox
Private mLstFoldersNames As ListBox
Private WithEvents mTDBMail As TDBGrid
Attribute mTDBMail.VB_VarHelpID = -1
Private mData As Data
Private mTxtMail As RichTextBox
Private SortColumn As Long
Private mSortOrder As Integer
Private mSortString As String

Private Sub Class_Initialize()
    On Error GoTo Err_Init
    Dim FileName As String
    FileName = App.Path & "\mail.mdb"
    If Dir(FileName, vbNormal Or vbArchive) = "" Then
        If Dir(App.Path & "\blank.mdb") > "" Then
            Name App.Path & "\blank.mdb" As FileName
        End If
    End If
    Set DB = OpenDatabase(App.Path & "\mail.mdb")
    Set mLstFolders = frmMain.lstFolders
    Set mLstFoldersData = frmMain.lstFoldersData
    Set mLstFoldersNames = frmMain.lstFoldersNames
    Set mTDBMail = frmMain.tdbMail
    Set mTxtMail = frmMain.txtMail
    Set mData = frmMain.Data1
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Class_Initialize", Err.Number, Err.Description
End Sub

Private Sub Class_Terminate()
    On Error GoTo Err_Init
    Set DB = Nothing
    Set mLstFolders = Nothing
    Set mLstFoldersData = Nothing
    Set mLstFoldersNames = Nothing
    Set mTDBMail = Nothing
    Set mTxtMail = Nothing
    Set mData = Nothing
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Class_Terminate", Err.Number, Err.Description
End Sub

Friend Sub AddUser(User As typeUser)
    On Error GoTo Err_Init
    With User
        SQL = "SELECT * FROM USERS"
        Set rsTemp = DB.OpenRecordset(SQL)
        rsTemp.AddNew
        rsTemp!Name = .Name
        rsTemp!Email = .Email
        rsTemp!Password = .Password
        rsTemp!POP3 = .POP3
        rsTemp!SMTP = .SMTP
        rsTemp.Update
        rsTemp.Close
        Set rsTemp = Nothing
    End With
    Exit Sub

Err_Init:
    HandleError CurrentModule, "AddUser", Err.Number, Err.Description
End Sub

Friend Sub EditUser()
    On Error GoTo Err_Init
    With User
        SQL = "SELECT * FROM USERS WHERE ID = " & .ID
        Set rsTemp = DB.OpenRecordset(SQL)
        rsTemp.Edit
        rsTemp!Name = .Name
        rsTemp!Email = .Email
        rsTemp!Password = .Password
        rsTemp!POP3 = .POP3
        rsTemp!SMTP = .SMTP
        rsTemp.Update
        rsTemp.Close
        Set rsTemp = Nothing
    End With
    Exit Sub

Err_Init:
    HandleError CurrentModule, "EditUser", Err.Number, Err.Description
End Sub

Friend Property Get GetUserList() As typeUser()
    On Error GoTo Err_Init
    Dim ctr As Long, UserList() As typeUser
    ReDim UserList(0 To 0)
    SQL = "SELECT * FROM Users"
    Set rsTemp = DB.OpenRecordset(SQL)
    If rsTemp.RecordCount = 0 Then
        rsTemp.Close
        Set rsTemp = Nothing
        GetUserList = UserList
        Exit Property
    End If
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        ReDim Preserve UserList(0 To ctr)
        With UserList(ctr)
            .ID = rsTemp!ID
            .Name = rsTemp!Name
            .Email = rsTemp!Email
            .Password = rsTemp!Password
            .SMTP = rsTemp!SMTP
            .POP3 = rsTemp!POP3
        End With
        ctr = ctr + 1
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    GetUserList = UserList
    Exit Property

Err_Init:
    HandleError CurrentModule, "GetUserList", Err.Number, Err.Description
End Property

Friend Function SaveMail(ByVal s As String) As Boolean
    'Saves the mail to the database
    Dim mail As typeMail, sDate As String, FileName As String
    Dim NumAttachments
    Dim c As Long, c2 As Long
    On Error GoTo Err_Init
    SaveMail = True
    If Left$(s, 1) = "+" Then
        c = InStr(1, s, vbCrLf)
        s = Right$(s, Len(s) - c - 1)
    End If
    FileName = App.Path & "\temp.txt"
    SaveFile s, FileName
    NumAttachments = GetNumFilesToDecode(FileName)
    c = InStr(1, s, vbCrLf & vbCrLf)
    With mail
        .Attachments = NumAttachments
        .Header = Left$(s, c - 1)
        .Body = Right$(s, Len(s) - (c + 3))
        .From = GetHeaderBit("From:", .Header)
        .To = GetHeaderBit("To:", .Header)
        .Subject = GetHeaderBit("Subject:", .Header)
        sDate = GetHeaderBit("Date:", .Header)
        If Mid$(sDate, 4, 1) = "," Then
            sDate = Right$(sDate, Len(sDate) - 5)
        End If
        If Len(sDate) > 23 Then
            sDate = Left$(sDate, 20)
        End If
        If Len(sDate) = 0 Then
            .Date = Now
        Else
            On Error Resume Next
            .Date = CDate(sDate)
            If Err.Number = 13 Then
                'non-standard date.
                'don't fart with it - just slap today's date on there. :-P
                .Date = Now
            End If
            On Error GoTo Err_Init
        End If
        Set rsTemp = DB.OpenRecordset("Mail")
        rsTemp.AddNew
        rsTemp!Folder = 1
        rsTemp!UserID = User.ID
        rsTemp!From = .From
        rsTemp!To = .To
        rsTemp!Subject = .Subject
        rsTemp!Date = .Date
        rsTemp!UnRead = True
        rsTemp!Attachments = .Attachments
        rsTemp!Header = .Header
        rsTemp!Body = .Body
        rsTemp.Update
        rsTemp.Close
        Set rsTemp = Nothing
    End With
    Exit Function
Err_Init:
    SaveMail = False
End Function

Friend Function GetHeaderBit(ByVal FindPhrase As String, ByVal Header As String) As String
    On Error GoTo Err_Init
    Dim c As Long, c2 As Long
    Header = vbCrLf & Header & vbCrLf 'in case 'FROM' is in the first position
    c = InStr(1, Header, vbCrLf & FindPhrase)
    c2 = InStr(c + 2, Header, vbCrLf)
    If c = 0 Or c2 = 0 Then
        'skip it
        GetHeaderBit = ""
    Else
        GetHeaderBit = Mid$(Header, c + Len(FindPhrase) + 3, c2 - c - Len(FindPhrase) - 3)
    End If
    Exit Function

Err_Init:
    HandleError CurrentModule, "GetHeaderBit", Err.Number, Err.Description
End Function

Friend Sub LoadFolders()
    On Error GoTo Err_Init
    Dim i As Long, UnRead As Long, UnReadString As String
    mLstFolders.Clear
    mLstFoldersData.Clear
    mLstFoldersNames.Clear
    SQL = "SELECT * FROM Folders"
    Set rsTemp = DB.OpenRecordset(SQL)
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        'Add the folder to the list
        'Refresh the folders, etc.
        mLstFolders.AddItem rsTemp!Name & UnReadString
        mLstFoldersData.AddItem rsTemp!ID
        mLstFoldersNames.AddItem rsTemp!Name
        'Calculate how many unread messages there are in the folder.
        UnRead = RefreshUnreadCounts(i)
        If UnRead > 0 Then
            UnReadString = " (" & UnRead & ")"
        Else
            UnReadString = ""
        End If
        mLstFolders.List(i) = rsTemp!Name & UnReadString
        'Move to the next record
        i = i + 1
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    mData.DatabaseName = App.Path & "\mail.mdb"
    
    Exit Sub

Err_Init:
    HandleError CurrentModule, "LoadFolders", Err.Number, Err.Description
    Resume Next
End Sub

Friend Sub LoadMessages()
    On Error GoTo Err_Init
    Dim Folder As Long
    '4=from,5=to
    If mSortString = "" Then
        mSortString = " ORDER BY DATE DESC"
        SortColumn = 7
    End If
    Folder = mLstFoldersData.List(mLstFolders.ListIndex)
    SQL = "SELECT * FROM Mail WHERE USERID = " & User.ID & " AND FOLDER = " & Folder & mSortString
    mData.RecordSource = SQL
    mData.Refresh
    mTDBMail.Refresh
    If Folder = 2 Or Folder = 3 Then
        mTDBMail.Columns(4).Visible = False
        mTDBMail.Columns(5).Visible = True
    Else
        mTDBMail.Columns(4).Visible = True
        mTDBMail.Columns(5).Visible = False
    End If
    Exit Sub
Err_Init:
    If Err.Number = 3261 Then
        MsgBox "Unable to start EZMailer, as the database is open by another program."
        ShutDown
    End If
End Sub

Private Sub mLstFolders_Click()
    On Error GoTo Err_Init
    mLstFolders_MouseUp 1, 0, 0, 0
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mLstFolders_Click", Err.Number, Err.Description
End Sub

Private Sub mLstFolders_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Init
    Dim RetVal As Long, FolderName As String, FolderID As Long, i As Long
    If KeyCode = vbKeyDelete Then
        i = mLstFolders.ListIndex
        FolderName = mLstFolders.List(i)
        FolderID = mLstFoldersData.List(i)
        If FolderID < 5 Then
            MsgBox "You can't delete the first 4 folders!", vbInformation
            Exit Sub
        End If
        RetVal = MsgBox("Delete folder " & FolderName & " and all messages within?", vbOKCancel)
        If RetVal = vbCancel Then
            Exit Sub
        End If
        'Delete all messages
        SQL = "DELETE * FROM MAIL WHERE FOLDERID = " & FolderID
        DB.Execute SQL
        'Delete the folder
        SQL = "DELETE * FROM FOLDERS WHERE FOLDERID = " & FolderID
        DB.Execute SQL
        mLstFolders.ListIndex = 0
        LoadMessages
        Status 2, "Folder " & FolderName & " was deleted."
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mLstFolders_KeyUp", Err.Number, Err.Description
End Sub

Private Sub mLstFolders_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Init
    Dim s As String
    If Button = 2 Then
        'right click
        s = InputBox("Enter new folder name:", "New Folder", "New Folder")
        If s = "" Then
            Exit Sub
        End If
        'Add the folder
        SQL = "INSERT INTO Folders (Name) VALUES ('" & s & "');"
        DB.Execute SQL
        LoadFolders
        LoadMessages
    Else
        'left click
        LoadMessages
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mLstFolders_MouseUp", Err.Number, Err.Description
End Sub

Private Sub mTDBMail_DragCell(ByVal SplitIndex As Integer, RowBookmark As Variant, ByVal ColIndex As Integer)
    On Error GoTo Err_Init
    ' Set the current cell to the one being dragged
    If ColIndex = -1 Then
        'dragging multiple rows
    Else
        'dragging single row - select it
        mTDBMail.Col = ColIndex
        mTDBMail.Bookmark = RowBookmark
    End If
    
    ' Set up drag operation, such as creating visual effects by
    ' highlighting the cell or row being dragged.
    
    ' Use VB manual drag support (put mtdbmail into drag mode)
    mTDBMail.Drag vbBeginDrag
    
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mTDBMail_DragCell", Err.Number, Err.Description
End Sub

Private Sub mLstFolders_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    On Error GoTo Err_Init
    Dim i As Long
    Const LB_SETSEL = &H185
    Const LB_SETCURSEL = &H186
     
    i = ListRowCalc(mLstFolders, Y)
    
    mLstFolders.SetFocus
    
    'single-select list
    'wParam = = 0-based index
    'lParam = not used, must be 0
    Call SendMessage(mLstFolders.hWnd, LB_SETCURSEL, i, ByVal 0&)
     
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mLstFolders_DragOver", Err.Number, Err.Description
End Sub

Private Sub mLstFolders_DragDrop(Source As Control, X As Single, Y As Single)
    On Error GoTo Err_Init
    'Executed when the item is dropped on the lstFolder control
    Dim i As Long, ctr As Long, s As String
    Dim Bookmark As Variant

    'Determine which folder the messages were dragged on to
    i = ListRowCalc(mLstFolders, Y)
    
    'if no records are selected, physically select the first one.
    If mTDBMail.SelBookmarks.Count = 0 Then
        mTDBMail.SelBookmarks.Add mTDBMail.GetBookmark(0)
    End If
    
    'apply the changes to all the selected records.
    With mData.Recordset
        For Each Bookmark In mTDBMail.SelBookmarks
            mTDBMail.Bookmark = Bookmark
            .Edit
            !Folder = frmMain.lstFoldersData.List(i)
            .Update
            ctr = ctr + 1
        Next Bookmark
    End With
    
    mData.Refresh
    
    If ctr = 1 Then
        s = "Message was moved to folder " & mLstFolders.List(i)
    Else
        s = ctr & " messages were moved to folder " & mLstFolders.List(i)
        LoadFolders
    End If
    
    Status 2, s
    
    mTDBMail.Drag vbEndDrag
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mLstFolders_DragDrop", Err.Number, Err.Description
End Sub

Private Function ListRowCalc(lstTemp As Control, ByVal Y As Single) As Integer
    On Error GoTo Err_Init
    Const LB_GETITEMHEIGHT = &H1A1
    'Determines the height of each item in ListBox control in pixels
    Dim ItemHeight As Integer
    ItemHeight = SendMessage(lstTemp.hWnd, LB_GETITEMHEIGHT, 0, 0)
    ListRowCalc = min(((Y / Screen.TwipsPerPixelY) \ ItemHeight) + _
                  lstTemp.TopIndex, lstTemp.ListCount - 1)
    Exit Function

Err_Init:
    HandleError CurrentModule, "ListRowCalc", Err.Number, Err.Description
End Function

Private Function min(X As Integer, Y As Integer) As Integer
    On Error GoTo Err_Init
    If X > Y Then min = Y Else min = X
    Exit Function

Err_Init:
    HandleError CurrentModule, "min", Err.Number, Err.Description
End Function

Private Sub mTDBMail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If StrComp(Screen.ActiveControl.Name, "mtdbmail", vbTextCompare) = 0 Then
    On Error GoTo Err_Init
        'ok
    Else
        mTDBMail.SetFocus
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mTDBMail_MouseUp", Err.Number, Err.Description
End Sub

Private Sub mTDBMail_OLEStartDrag(ByVal Data As TrueDBGrid60.DataObject, AllowedEffects As Long)
    MsgBox "start drag"
    On Error GoTo Err_Init
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mTDBMail_OLEStartDrag", Err.Number, Err.Description
End Sub

Private Sub mTDBMail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Err_Init
    Dim s As String, i As Long, Cnt As Long
    If Receiving Then
        Exit Sub
        'don't process any unread flags while receiving
    End If
    
    mTxtMail.Text = mData.Recordset!Body
    
    If mData.Recordset!UnRead = True Then
        mData.Recordset.Edit
        mData.Recordset.UnRead = False
        mData.Recordset.Update
        i = mLstFolders.ListIndex
        Cnt = RefreshUnreadCounts(i)
        If Cnt = 0 Then
            mLstFolders.List(i) = mLstFoldersNames.List(i)
        Else
            mLstFolders.List(i) = mLstFoldersNames.List(i) & " (" & Cnt & ")"
        End If
    End If
    Exit Sub
Err_Init:
    If Err.Number = 3021 Then
        mTxtMail.Text = ""
    Else
        HandleError CurrentModule, "mTDBMail_RowColChange", Err.Number, Err.Description
    End If
End Sub

Friend Sub ExportMail()
    On Error GoTo Err_Init
    Dim FileName As String
    FileName = App.Path & "\body.txt"
    SaveFile mData.Recordset!Header & vbCrLf & vbCrLf & mData.Recordset!Body, FileName
    OpenFile FileName
    Exit Sub

Err_Init:
    HandleError CurrentModule, "ExportMail", Err.Number, Err.Description
End Sub

Friend Sub SaveAttachments()
    On Error GoTo Err_Init
    Dim FileName As String, NumAttachments As Long, FolderName As String
    Dim dummy As Long, X As Long, s() As String, c As Long
    Dim DecodedFileName As String * 256, FileList As String
    FileName = App.Path & "\temp.txt"
    SaveFile mData.Recordset!Header & vbCrLf & vbCrLf & mData.Recordset!Body, FileName
    NumAttachments = GetNumFilesToDecode(FileName)
    If NumAttachments = 0 Then
        MsgBox "No attachments to save!", vbInformation
        Exit Sub
    End If
    FolderName = GetFolderName()
    If FolderName = "" Then
        FolderName = "c:\windows\desktop"
        If Dir(FolderName, vbDirectory) = "" Then
            FolderName = App.Path & "\Attachments"
            If Dir(FolderName, vbDirectory) = "" Then
                MkDir FolderName
            End If
        Else
            FolderName = FolderName & "\EZMailer " & Format(Now, "MM.DD.YY HH.MM.SS")
            If Dir(FolderName, vbDirectory) = "" Then
                MkDir FolderName
            End If
        End If
    End If
    FileName = FolderName & "\ORIGMSG.TXT"
    FileCopy App.Path & "\temp.txt", FileName
    'For some bizarre reason, I have to execute this routine in the NEW
    'directory, else it will put the files in the wrong place!
    NumAttachments = GetNumFilesToDecode(FileName)
    'Create list of files to decode.
    For X = 0 To NumAttachments - 1
        dummy = GetEncodedFile(DecodedFileName, X)
        ReDim Preserve s(0 To X)
        c = InStr(1, DecodedFileName, Chr$(0))
        s(X) = Left$(DecodedFileName, c - 1)
        FileList = FileList & s(X)
        If X < NumAttachments - 1 Then
            FileList = FileList & ";"
        End If
    Next X
    Status 2, X & " file(s) saved to " & FolderName ' Files attached: " & FileList
    dummy = DecodeFile(FileName, FileList, PROMPT_NONE)
    Exit Sub

Err_Init:
    HandleError CurrentModule, "SaveAttachments", Err.Number, Err.Description
End Sub

Friend Sub Reply()
    On Error GoTo Err_Init
    If mData.Recordset.RecordCount = 0 Then
        'skip it
    Else
        frmMail.cboTo.Text = mData.Recordset!From
        frmMail.txtSubject.Text = "Re: " & mData.Recordset!Subject
        frmMail.Show
        frmMail.txtMessage.SetFocus
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Reply", Err.Number, Err.Description
End Sub

Friend Sub Forward()
    On Error GoTo Err_Init
    If mData.Recordset.RecordCount = 0 Then
        'skip it
    Else
        frmMail.cboTo.Text = ""
        frmMail.txtSubject.Text = mData.Recordset!Subject
        frmMail.txtMessage.Text = mData.Recordset!Body
        frmMail.txtMessage.Tag = "Preformatted"
        frmMail.Show
        frmMail.cboTo.SetFocus
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Forward", Err.Number, Err.Description
End Sub

Friend Sub StoreDraft(m As typeMail)
    On Error GoTo Err_Init
    Dim ToEmail As String, ToName As String, s As String, t As String
    Dim s2() As String, max As Long
    s = m.To
    Do While InStr(1, s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s2 = Split(s, " ")
    max = UBound(s2, 1)
    If max = 0 Then
        ToName = vbQuote & s2(max) & vbQuote
    Else
        If InStr(1, s2(max), "@") > 0 Then
            'well ok then
        Else
            'switch them around!
            s = Right$(s, Len(s) - Len(s2(0)) - 1)
            s = s & " " & s2(0)
            s2 = Split(s, " ")
        End If
        ToName = Left$(s, Len(s) - Len(s2(max)) - 1)
    End If
    If InStr(1, ToName, vbQuote) = 0 Then
        ToName = vbQuote & ToName & vbQuote
    End If
    ToEmail = s2(max)
    If InStr(1, ToEmail, "<") = 0 Then
        ToEmail = "<" & ToEmail & ">"
    End If
    s = ToName & " " & ToEmail
    m.To = s
    
    With mData.Recordset
        .AddNew
        !UserID = User.ID
        !Folder = 2
        !To = m.To
        !Subject = m.Subject
        !Body = m.Body
        !Boundary = m.Boundary & ""
        .Update
        mData.Refresh
    End With
    Exit Sub

Err_Init:
    HandleError CurrentModule, "StoreDraft", Err.Number, Err.Description
End Sub

Friend Sub DeleteMessage()
    On Error GoTo Err_Init
    Dim Bookmark As Variant
    
    'if no records are selected, physically select the first one.
    If mTDBMail.SelBookmarks.Count = 0 Then
        mTDBMail.SelBookmarks.Add mTDBMail.GetBookmark(0)
    End If
    
    'apply the changes to all the selected records.
    With mData.Recordset
        For Each Bookmark In mTDBMail.SelBookmarks
            mTDBMail.Bookmark = Bookmark
            mData.Recordset.Edit
            If mData.Recordset!Folder = 4 Then
                'truly delete it
                mData.Recordset.Delete
            Else
                'just move it to 'deleted items'
                mData.Recordset!Folder = 4
                mData.Recordset.Update
            End If
        Next Bookmark
    End With
    
    mData.Refresh
    Exit Sub

Err_Init:
    HandleError CurrentModule, "DeleteMessage", Err.Number, Err.Description
End Sub

Friend Sub SendMail()
    On Error GoTo Err_Init
    Dim s As String, s2() As String, ToName As String, ToEmail As String
    Dim max As Long, Num As Long, B As String
    mLstFolders.ListIndex = 1
    LoadMessages
    If mData.Recordset.RecordCount = 0 Then
        Status 2, "No messages waiting to be sent."
    Else
        Num = mData.Recordset.RecordCount
        With mData.Recordset
            .MoveFirst
            Do While Not .EOF
                s = !To
                s2 = Split(s, " ")
                max = UBound(s2, 1)
                ToEmail = s2(max)
                ToEmail = Mid$(ToEmail, 2, Len(ToEmail) - 2)
                ToName = Left$(s, Len(s) - Len(s2(max)) - 1)
                ToName = Mid$(ToName, 2, Len(ToName) - 2)
                If CN.SendEmail(ToName, ToEmail, !Subject, !Body, , !Boundary & "") = True Then
                    'Place in the 'sent items' folder
                    .Edit
                    !Folder = 3
                    .Update
                Else
                    'Keep in the 'outbox' folder.
                    Status 2, "Error sending message."
                End If
                .MoveNext
            Loop
        End With
        mData.Refresh
        Status 2, Num & " messages sent."
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "SendMail", Err.Number, Err.Description
End Sub

Private Sub mTDBMail_HeadClick(ByVal ColIndex As Integer)
    'sort by the desired column
    On Error GoTo Err_Init
    Screen.MousePointer = vbHourglass
    mTDBMail.MousePointer = vbHourglass
    SortOrder ColIndex
    mTDBMail.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    Exit Sub

Err_Init:
    HandleError CurrentModule, "mTDBMail_HeadClick", Err.Number, Err.Description
End Sub

Private Sub SortOrder(ByVal Index As Integer)
'Change the sort order of the displayed records.
    Dim s As String
    On Error GoTo Err_Init
    If Index = SortColumn Then
        'same field - just swap sort direction
        mSortOrder = IIf(mSortOrder = 1, 0, 1)
    Else
        'different field
        SortColumn = Index
        mSortOrder = 0
    End If
    If mSortOrder = 0 Then
        s = " DESC"
    Else
        s = " ASC"
    End If
    mSortString = " ORDER BY " & mTDBMail.Columns(Index).Name & s
    LoadMessages
    Status 2, "Sorted by " & SortColumn & " " & mTDBMail.Columns(Index).Caption & s
    Exit Sub
Err_Init:
    HandleError CurrentModule, "SortOrder", Err.Number, Err.Description
    Resume Next
End Sub

Friend Sub LoadValueItems(ByVal Table As String, VS As ValueItems)
    On Error GoTo Err_Init
    Dim v As ValueItem
    SQL = "SELECT * FROM " & Table
    Set rsTemp = DB.OpenRecordset(SQL)
    rsTemp.MoveFirst
    VS.Clear
    Do While Not rsTemp.EOF
        Set v = New ValueItem
        v.Value = rsTemp!ID
        v.DisplayValue = rsTemp!Name
        VS.Add v
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    VS.Translate = True
    Exit Sub

Err_Init:
    HandleError CurrentModule, "LoadValueItems", Err.Number, Err.Description
End Sub

Friend Sub LoadRules()
    On Error GoTo Err_Init
    Dim ctr As Long
    ReDim Rules(0 To 0)
    SQL = "SELECT * FROM RULES"
    Set rsTemp = DB.OpenRecordset(SQL)
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    End If
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        ReDim Preserve Rules(0 To ctr)
        With Rules(ctr)
            .PartID = rsTemp!PartID
            .FolderID = rsTemp!FolderID
            .FindPhrase = rsTemp!FindPhrase
        End With
        ctr = ctr + 1
        rsTemp.MoveNext
    Loop
    Exit Sub

Err_Init:
    HandleError CurrentModule, "LoadRules", Err.Number, Err.Description
End Sub

Friend Sub ApplyRules()
    On Error GoTo Err_Init
    Dim i As Long, max As Long, s As String
    If Rules(i).PartID = 0 Then
        'no rules specified
        Exit Sub
    End If
    max = UBound(Rules, 1)
    'Navigate to the correct folder
    mLstFolders.ListIndex = 0
    LoadMessages
    'If no messages in the folder exit early
    If mData.Recordset.RecordCount = 0 Then
        Exit Sub
    End If
    
    With mData.Recordset
        .MoveFirst
        Do While Not .EOF
            For i = 0 To max
                Select Case Rules(i).PartID
                    Case 1: s = !Subject
                    Case 2: s = !From
                    Case 3: s = !To
                    Case 4: s = !Body
                    Case Else: s = ""
                End Select
                If InStr(1, s, Rules(i).FindPhrase, vbTextCompare) > 0 Then
                    'found it - so route it
                    .Edit
                    !Folder = Rules(i).FolderID
                    .Update
                    'exit early - do no more checks
                    Exit For
                End If
            Next i
            .MoveNext
        Loop
    End With
    LoadFolders
    mLstFolders.ListIndex = 0
    LoadMessages
    Exit Sub

Err_Init:
    HandleError CurrentModule, "ApplyRules", Err.Number, Err.Description
End Sub

Friend Sub LoadTo(cboBox As ComboBox)
    On Error GoTo Err_Init
    SQL = "SELECT NAME FROM TO"
    Set rsTemp = DB.OpenRecordset(SQL)
    If rsTemp.RecordCount = 0 Then
        Exit Sub
    End If
    rsTemp.MoveFirst
    cboBox.Clear
    Do While Not rsTemp.EOF
        cboBox.AddItem rsTemp!Name
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Sub

Err_Init:
    HandleError CurrentModule, "LoadTo", Err.Number, Err.Description
End Sub

Friend Function AddTo(ByVal s As String) As Boolean
    'Extract out the email part
    On Error GoTo Err_Init
    Dim c As Long, s2() As String, i As Long, FoundIt As Boolean
    Dim Email As String
    s2 = Split(s, " ")
    For i = 0 To UBound(s2, 1)
        c = InStr(1, s2(i), "@", vbTextCompare)
        If c > 1 Then
            'found it
            FoundIt = True
            Email = s2(i)
            If Left$(Email, 1) = "<" Then
                Email = Mid$(Email, 2, Len(Email) - 2)
            End If
            Exit For
        End If
    Next i
    If FoundIt Then
        Set rsTemp = DB.OpenRecordset("SELECT * FROM TO WHERE EMAIL = '" & Email & "'")
        If rsTemp.RecordCount = 0 Then
            AddTo = True
            SQL = "INSERT INTO To (Name, Email) VALUES ('" & s & "','" & Email & "');"
            DB.Execute SQL
        Else
            AddTo = False
        End If
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        AddTo = False
    End If
    Exit Function
Err_Init:
    HandleError CurrentModule, "AddTo", Err.Number, Err.Description
    Resume Next
End Function

Friend Sub DeleteUser(ByVal i As Long)
    On Error GoTo Err_Init
    SQL = "DELETE * FROM MAIL WHERE USERID = " & i
    DB.Execute SQL
    SQL = "DELETE * FROM USERS WHERE ID = " & i
    DB.Execute SQL
    Exit Sub

Err_Init:
    HandleError CurrentModule, "DeleteUser", Err.Number, Err.Description
End Sub

Friend Sub Programmer()
    On Error GoTo Err_Init
    Dim bit As String, s As String
    s = mData.Recordset!Header
    bit = GetHeaderBit("Subject:", s)
    MsgBox bit
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Programmer", Err.Number, Err.Description
End Sub

Friend Function RefreshUnreadCounts(ByVal Folder As Long) As Long
    On Error GoTo Err_Init
    Dim Cnt As Long
    SQL = "SELECT COUNT(*) FROM Mail WHERE USERID = " & User.ID & " AND FOLDER = " & mLstFoldersData.List(Folder) & " AND UNREAD=True"
    Set rsTemp2 = DB.OpenRecordset(SQL)
    RefreshUnreadCounts = rsTemp2.Fields(0)
    rsTemp2.Close
    Set rsTemp2 = Nothing
    Exit Function

Err_Init:
    HandleError CurrentModule, "RefreshUnreadCounts", Err.Number, Err.Description
End Function
