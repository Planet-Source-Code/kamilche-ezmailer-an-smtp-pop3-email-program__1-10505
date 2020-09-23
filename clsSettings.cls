VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const CurrentModule As String = "clsSettings"
Private t As String
Private p As String
Private mHorizontalDivider As Double
Private mVerticalDivider As Double
Private mColor(1 To 3) As Long

Friend Property Get HorizontalDivider() As Double
    On Error GoTo Err_Init
    HorizontalDivider = mHorizontalDivider
    Exit Property

Err_Init:
    HandleError CurrentModule, "HorizontalDivider", Err.Number, Err.Description
End Property

Friend Property Let HorizontalDivider(v As Double)
    On Error GoTo Err_Init
    If v < 0.05 Then
        v = 0.05
    End If
    mHorizontalDivider = v
    Exit Property

Err_Init:
    HandleError CurrentModule, "HorizontalDivider", Err.Number, Err.Description
End Property

Friend Property Get VerticalDivider() As Double
    On Error GoTo Err_Init
    VerticalDivider = mVerticalDivider
    Exit Property

Err_Init:
    HandleError CurrentModule, "VerticalDivider", Err.Number, Err.Description
End Property

Friend Property Let VerticalDivider(v As Double)
    On Error GoTo Err_Init
    If v < 0.05 Then
        v = 0.05
    End If
    mVerticalDivider = v
    Exit Property

Err_Init:
    HandleError CurrentModule, "VerticalDivider", Err.Number, Err.Description
End Property

Friend Property Get Color(ByVal Which As Long) As Long
    On Error GoTo Err_Init
    Color = mColor(Which)
    Exit Property

Err_Init:
    HandleError CurrentModule, "Color", Err.Number, Err.Description
End Property

Friend Property Let Color(ByVal Which As Long, ByVal c As Long)
    On Error GoTo Err_Init
    mColor(Which) = c
    Exit Property

Err_Init:
    HandleError CurrentModule, "Color", Err.Number, Err.Description
End Property

Private Sub Class_Initialize()
    t = App.Title
    On Error GoTo Err_Init
    p = "Preferences"
    mHorizontalDivider = GetSetting(t, p, "HorizontalDivider", 0.43)
    mVerticalDivider = GetSetting(t, p, "VerticalDivider", 0.177)
    mColor(1) = GetSetting(t, p, "Color1", 8421504)
    mColor(2) = GetSetting(t, p, "Color2", 4227200)
    mColor(3) = GetSetting(t, p, "Color3", 12632256)
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Class_Initialize", Err.Number, Err.Description
End Sub

Private Sub Class_Terminate()
    SaveSetting t, p, "HorizontalDivider", mHorizontalDivider
    On Error GoTo Err_Init
    SaveSetting t, p, "VerticalDivider", mVerticalDivider
    SaveSetting t, p, "Color1", mColor(1)
    SaveSetting t, p, "Color2", mColor(2)
    SaveSetting t, p, "Color3", mColor(3)
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Class_Terminate", Err.Number, Err.Description

End Sub

