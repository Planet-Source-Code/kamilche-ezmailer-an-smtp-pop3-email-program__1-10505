VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmAddresses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Addresses"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "F:\Code\EZMailer\mail.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "To"
      Top             =   6840
      Visible         =   0   'False
      Width           =   4140
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "frmAddresses.frx":0000
      Height          =   6135
      Left            =   120
      OleObjectBlob   =   "frmAddresses.frx":0014
      TabIndex        =   0
      Top             =   840
      Width           =   9135
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAddresses.frx":270A
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "frmAddresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const CurrentModule As String = "frmAddresses"

Private Sub Form_Load()
    On Error GoTo Err_Init
    Data1.DatabaseName = App.Path & "\mail.mdb"
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_Load", Err.Number, Err.Description
End Sub
