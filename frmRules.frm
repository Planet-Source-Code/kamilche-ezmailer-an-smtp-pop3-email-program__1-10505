VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmRules 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                                                                          "
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "F:\Code\EZMailer\mail.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Rules"
      Top             =   6960
      Visible         =   0   'False
      Width           =   4140
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "frmRules.frx":0000
      Height          =   6015
      Left            =   240
      OleObjectBlob   =   "frmRules.frx":0014
      TabIndex        =   1
      Top             =   1080
      Width           =   9135
   End
   Begin VB.Label Label1 
      Caption         =   $"frmRules.frx":332A
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const CurrentModule As String = "frmRules"

Private Sub Form_Load()
    On Error GoTo Err_Init
    Dim v As ValueItem, rsTemp As Recordset, SQL As String
    Data1.DatabaseName = App.Path & "\mail.mdb"
    'Load value items
    DB.LoadValueItems "Parts", TDBGrid1.Columns(1).ValueItems
    DB.LoadValueItems "Folders", TDBGrid1.Columns(3).ValueItems
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_Load", Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Err_Init
    DB.LoadRules
    DB.ApplyRules
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Form_Unload", Err.Number, Err.Description
End Sub

