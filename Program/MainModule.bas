Attribute VB_Name = "MainModule"
Option Explicit

Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As KeyCodeConstants) As Integer
Declare Function GetKeyboardState Lib "user32.dll" (pbKeyState As Byte) As Long
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


Sub Main()
    On Error Resume Next
    frmMain.ShowForm
End Sub
