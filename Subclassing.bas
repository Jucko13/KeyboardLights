Attribute VB_Name = "Subclassing"
'Form1 code
'Option Explicit

'Private Sub Form_Load()
'Hook Me.hWnd, True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'Hook Me.hWnd, False
'End Sub

'BAS module code
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Long

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
    ByVal idHook As Long, _
    ByVal lpfn As Long, _
    ByVal hmod As Long, _
    ByVal dwThreadId As Long _
) As Long
 
Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
    ByVal hHook As Long _
) As Long
 
Private Declare Function CallNextHookEx Lib "user32" ( _
    ByVal hHook As Long, _
    ByVal ncode As Long, _
    ByVal wParam As Long, _
    lParam As Any _
) As Long
 
 
Private Const WM_KEYUP As Long = &H101
Private Const WM_KEYDOWN As Long = &H100


Public Const WH_KEYBOARD_LL As Long = 13
Private Const HC_ACTION As Long = 0
Private Const HC_NOREMOVE As Long = 3

Private Const VK_VOLUMEDOWN As Long = 174
Private Const VK_VOLUMEUP As Long = 175
Private Const VK_VOLUMEMUTE As Long = 173

Private Const VK_SHIFT = &H10
Private Const VK_CONTROL = &H11
Private Const VK_MENU = &H12

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type


Private hHook As Long
Public IsHooked As Boolean
 
Public Sub SetKeyboardHook()
    If IsHooked Then
        MsgBox "Don't hook WH_KEYBOARD_LL twice or you will be unable to unhook it."
    Else
        hHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
        IsHooked = True
    End If
End Sub
 
Public Sub RemoveKeyboardHook()
    UnhookWindowsHookEx hHook
    IsHooked = False
End Sub
 
Public Function LowLevelKeyboardProc(ByVal uCode As Long, ByVal wParam As Long, lParam As KBDLLHOOKSTRUCT) As Long
    Dim i As Long
    Dim newBrightness As Long
    Dim Gain As Long
    
    Const myOwnPC As Boolean = True
    
    If uCode >= 0 Then
        Select Case uCode
            Case HC_ACTION
            'Debug.Print Hex(lParam.vkCode)
                If (lParam.vkCode = VK_VOLUMEUP Or lParam.vkCode = VK_VOLUMEDOWN) Then
                    If GetAsyncKeyState(VK_CONTROL) Then
                        If wParam = WM_KEYUP Then
                            
                            If lParam.vkCode = VK_VOLUMEUP Then
                                Gain = 10
                            Else
                                Gain = -10
                            End If
                            
                            For i = 0 To MonitorCount - 1
                                newBrightness = Monitors(i).lBrightnessCurrent
                                If myOwnPC Then
                                    newBrightness = newBrightness + Gain
                                    If newBrightness >= 100 And Gain > 0 Then GoTo next_i
                                    If newBrightness <= 0 And Gain < 5 Then GoTo next_i
                                Else
                                    If newBrightness >= Monitors(i).lBrightnessMaximum And Gain > 0 Then GoTo next_i
                                    If newBrightness <= Monitors(i).lBrightnessMinimum And Gain < 5 Then GoTo next_i
                                    
                                    newBrightness = newBrightness + Gain
                                    
                                    If newBrightness > Monitors(i).lBrightnessMaximum Then
                                        newBrightness = Monitors(i).lBrightnessMaximum
                                    ElseIf newBrightness < Monitors(i).lBrightnessMinimum Then
                                        newBrightness = Monitors(i).lBrightnessMinimum
                                    End If
                                End If
                                
                                Monitors(i).lBrightnessCurrent = newBrightness
                                
                                SetMonitorBrightness Monitors(i).pPhysicalInfo(0).hPhysicalMonitor, newBrightness
                                DoEvents
next_i:
                            Next i
                        End If
                        
                        LowLevelKeyboardProc = 1
                        Exit Function
                    End If
                    
                ElseIf lParam.vkCode = VK_VOLUMEMUTE Then
                    If GetAsyncKeyState(VK_CONTROL) Then
                        If wParam = WM_KEYUP Then
                            For i = 0 To MonitorCount - 1
                                'Debug.Print Monitors(i).lPowerModeCurrent
                                'If Monitors(i).lPowerModeCurrent = 1 Then
                                    SetVCPFeature Monitors(i).pPhysicalInfo(0).hPhysicalMonitor, CByte(&HD6), 5
                                    'PrintError
                                'Else
                                '    SetVCPFeature Monitors(i).pPhysicalInfo(0).hPhysicalMonitor, CByte(&HD6), 1
                                    'PrintError
                                'End If
                                
                                
                                GetVCPFeatureAndVCPFeatureReply Monitors(i).pPhysicalInfo(0).hPhysicalMonitor, &HD6, Monitors(i).lPowerModeSetting, Monitors(i).lPowerModeCurrent, Monitors(i).lPowerModeMaximum
                                'PrintError
                            Next i
                        End If
                        
                        LowLevelKeyboardProc = 1
                        Exit Function
                    End If
                End If
            Case HC_NOREMOVE
                'The message has not been removed from the message queue
        End Select
    End If
        
    LowLevelKeyboardProc = CallNextHookEx(hHook, uCode, wParam, lParam)
End Function


