VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmMain 
   Caption         =   "Keyboard Monitor"
   ClientHeight    =   1305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2520
   OleObjectBlob   =   "FrmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Update(KeyCode As MSForms.ReturnInteger, Pressed As Boolean)
    Dim ctl As Control
    For Each ctl In Me.Controls
        If ctl.Tag = CStr(KeyCode) Then
            Dim chk As Object
            Set chk = ctl
            chk.Value = Pressed
        End If
    Next ctl
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Update KeyCode, True
End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Update KeyCode, False
End Sub
