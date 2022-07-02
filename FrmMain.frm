VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmMain 
   Caption         =   "Keyboard Monitor"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2475
   OleObjectBlob   =   "FrmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long
Private Pressed_At(32 To 40) As Long


Private Sub PostDuration(dt As Long)
    If Len(Me.txtD1.Text) Then Me.txtD2.Text = Me.txtD1.Text  'make room
    Me.txtD1.Text = dt
End Sub

Private Sub Update(KeyCode As MSForms.ReturnInteger, Pressed As Boolean)
    Dim ctl As Control
    For Each ctl In Me.Controls
        If ctl.Tag = CStr(KeyCode) Then
            Dim chk As Object
            Set chk = ctl
            chk.Value = Pressed
            If Pressed Then
                'ignore repeated keys, otherwise the delay becomes meaningless
                If Pressed_At(KeyCode) = 0 Then Pressed_At(KeyCode) = GetTickCount
            Else
                PostDuration GetTickCount - Pressed_At(KeyCode)
                Pressed_At(KeyCode) = 0
            End If
        End If
    Next ctl
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Update KeyCode, True
End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Update KeyCode, False
End Sub
