VERSION 5.00
Begin VB.Form frmUser 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Done"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtPOP3 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox txtSMTP 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "POP3 Server:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   2310
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "SMTP Server:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Email:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   870
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "User Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const CurrentModule As String = "frmUser"

Private Sub cmdOK_Click()
    On Error GoTo Err_Init
    If frmUser.txtEmail = "" Then
        MsgBox "Please enter the user information before continuing.", vbCritical
    Else
        With frmUser
            User.Name = .txtName
            User.Password = .txtPassword
            User.Email = .txtEmail
            User.SMTP = .txtSMTP
            User.POP3 = .txtPOP3
        End With
        Unload Me
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "cmdOK_Click", Err.Number, Err.Description
End Sub

Private Sub Form_Load()
    On Error GoTo Err_Init
    With frmUser
        .txtName = User.Name
        .txtPassword = User.Password
        .txtEmail = User.Email
        .txtSMTP = User.SMTP
        .txtPOP3 = User.POP3
    End With
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_Load", Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmUser.txtEmail = "" Then
        End
    End If
End Sub
