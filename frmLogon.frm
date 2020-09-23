VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logon"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNew 
      Caption         =   "New User"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox lstUsers 
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   4890
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Choose User"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const CurrentModule As String = "frmLogon"
Private Users() As typeUser

Private Sub CancelButton_Click()
    On Error GoTo Err_Init
    End
    Exit Sub

Err_Init:
    HandleError CurrentModule, "CancelButton_Click", Err.Number, Err.Description
End Sub

Private Sub cmdNew_Click()
    On Error GoTo Err_Init
    frmUser.Show vbModal
    DB.AddUser User
    RefreshNames
    Exit Sub

Err_Init:
    HandleError CurrentModule, "cmdNew_Click", Err.Number, Err.Description
End Sub

Private Sub Form_Load()
    On Error GoTo Err_Init
    RefreshNames
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_Load", Err.Number, Err.Description
End Sub

Private Sub RefreshNames()
    On Error GoTo Err_Init
    Dim i As Long
    lstUsers.Clear
    Users = DB.GetUserList
    If UBound(Users, 1) = 0 Then
        If Users(0).Email = "" Then
            'No users have been entered
            cmdNew_Click
            Exit Sub
        End If
    End If
    For i = LBound(Users, 1) To UBound(Users, 1)
        lstUsers.AddItem Users(i).Name
    Next i
    lstUsers.ListIndex = 0
    Exit Sub

Err_Init:
    HandleError CurrentModule, "RefreshNames", Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Err_Init
    If User.Email = "" Then
        End
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_Unload", Err.Number, Err.Description
End Sub

Private Sub lstUsers_Click()
    On Error GoTo Err_Init
    User = Users(lstUsers.ListIndex)
    Exit Sub

Err_Init:
    HandleError CurrentModule, "lstUsers_Click", Err.Number, Err.Description
End Sub

Private Sub lstUsers_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Init
    Dim i As Long, RetVal As Long
    i = lstUsers.ListIndex
    If KeyCode = vbKeyDelete Then
        RetVal = MsgBox("Delete user " & User.Name & " and all their messages?", vbOKCancel)
        If RetVal = vbCancel Then
            Exit Sub
        End If
        DB.DeleteUser Users(i).ID
        RefreshNames
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "lstUsers_KeyUp", Err.Number, Err.Description
End Sub

Private Sub OKButton_Click()
    On Error GoTo Err_Init
    Unload Me
    frmMain.Show
    Exit Sub

Err_Init:
    HandleError CurrentModule, "OKButton_Click", Err.Number, Err.Description
End Sub
