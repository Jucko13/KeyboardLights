Attribute VB_Name = "DisplaySettings"
Option Explicit

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Declare Function FormatMessage Lib "Kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long


Public Const MONITORINFOF_PRIMARY = &H1
Public Const MONITOR_DEFAULTTONEAREST = &H2
Public Const MONITOR_DEFAULTTONULL = &H0
Public Const MONITOR_DEFAULTTOPRIMARY = &H1

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type


Public Type POINT
    X As Long
    y As Long
End Type


Public Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Public Declare Function MonitorFromPoint Lib "user32.dll" (ByVal X As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
Public Declare Function MonitorFromRect Lib "user32.dll" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
Public Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Public Declare Function EnumDisplayMonitors Lib "user32.dll" (ByVal hdc As Long, ByRef lprcClip As Any, ByVal lpfnEnum As Long, ByVal dwData As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long


Public Declare Function SetMonitorBrightness Lib "Dxva2.dll" (ByVal hMonitor As Long, ByVal dwNewBrightness As Long) As Long
Public Declare Function GetMonitorBrightness Lib "Dxva2.dll" (ByVal hMonitor As Long, ByRef pdwMinimumBrightness As Long, ByRef pdwCurrentBrightness As Long, ByRef pdwMaximumBrightness As Long) As Long

Public Declare Function SetMonitorContrast Lib "Dxva2.dll" (ByVal hMonitor As Long, ByVal dwNewContrast As Long) As Long
Public Declare Function GetMonitorContrast Lib "Dxva2.dll" (ByVal hMonitor As Long, ByRef pdwMinimumContrast As Long, ByRef pdwCurrentContrast As Long, ByRef pdwMaximumContrast As Long) As Long


Public Declare Function SetMonitorDisplayAreaSize Lib "Dxva2.dll" (ByVal hMonitor As Long, ByVal sizeType As MC_SIZE_TYPE, ByVal dwNewDisplayAreaWidthOrHeight As Long) As Long
Public Declare Function GetMonitorDisplayAreaSize Lib "Dxva2.dll" (ByVal hMonitor As Long, ByVal sizeType As MC_SIZE_TYPE, ByRef pdwMinimumWidthOrHeight As Long, ByRef pdwCurrentWidthOrHeight As Long, ByRef pdwMaximumWidthOrHeight As Long) As Long

Enum LPMC_VCP_CODE_TYPE
    MC_MOMENTARY = 0
    MC_SET_PARAMETER = 1
End Enum

Public Declare Function SetVCPFeature Lib "Dxva2.dll" (ByVal hMonitor As Long, ByVal bVCPCode As Byte, ByVal dwNewValue As Long) As Long
Public Declare Function GetVCPFeatureAndVCPFeatureReply Lib "Dxva2.dll" (ByVal hMonitor As Long, ByVal bVCPCode As Byte, ByRef pvct As LPMC_VCP_CODE_TYPE, ByRef pdwCurrentValue As Long, ByRef pdwMaximumValue As Long) As Long

Public Declare Function DestroyPhysicalMonitor Lib "Dxva2.dll" (ByVal hMonitor As Long) As Long




'Type PHYSICAL_MONITOR
'    lHandleHigh As Long
'    lHandleLow As Long
'    sDescription(0 To 255) As Byte
'End Type

Enum MC_SIZE_TYPE
    gWidth = 0
    gHeight = 1
End Enum


Type PHYSICAL_MONITOR
    hPhysicalMonitor As Long
    szPhysicalMonitorDescription(0 To 255) As Byte
End Type

Public Declare Function GetNumberOfPhysicalMonitorsFromHMONITOR Lib "Dxva2.dll" (ByVal hMonitor As Long, ByRef pdwNumberOfPhysicalMonitors As Long) As Long
Public Declare Function GetPhysicalMonitorsFromHMONITOR Lib "Dxva2.dll" (ByVal hMonitor As Long, ByVal dwPhysicalMonitorArraySize As Long, ByRef pPhysicalMonitorArray As PHYSICAL_MONITOR) As Long 'PHYSICAL_MONITOR


Public Type MonitorStats
    lBrightnessCurrent As Long
    lBrightnessMinimum As Long
    lBrightnessMaximum As Long
    
    lContrastCurrent As Long
    lContrastMinimum As Long
    lContrastMaximum As Long
    
    lScreenWidthCurrent As Long
    lScreenWidthMinimum As Long
    lScreenWidthMaximum As Long
    
    lScreenHeightCurrent As Long
    lScreenHeightMinimum As Long
    lScreenHeightMaximum As Long
    
    lPowerModeCurrent As Long
    lPowerModeMaximum As Long
    lPowerModeSetting As LPMC_VCP_CODE_TYPE
    
    lHandle As Long
    lMonitorCount As Long
    lDescription As String * 128
    pPhysicalInfo() As PHYSICAL_MONITOR
    lAllInfo As MONITORINFO
    
End Type

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)



Global Monitors() As MonitorStats
Global MonitorBrightness() As Long
Global MonitorCount As Long

Private Function ByteArrayToString(Bytes() As Byte) As String
    Dim iUnicode As Long, i As Long, j As Long
    
    On Error Resume Next
    i = UBound(Bytes)
    
    If (i < 1) Then
        'ANSI, just convert to unicode and return
        ByteArrayToString = StrConv(Bytes, vbUnicode)
        Exit Function
    End If
    i = i + 1
    
    'Examine the first two bytes
    CopyMemory iUnicode, Bytes(0), 2
    
    If iUnicode = Bytes(0) Then 'Unicode
        'Account for terminating null
        If (i Mod 2) Then i = i - 1
        'Set up a buffer to recieve the string
        ByteArrayToString = String$(i / 2, 0)
        'Copy to string
        CopyMemory ByVal StrPtr(ByteArrayToString), Bytes(0), i
    Else 'ANSI
        ByteArrayToString = StrConv(Bytes, vbUnicode)
    End If
                    
End Function


Public Sub ScanForMonitors()
    ReDim Monitors(0) As MonitorStats
    MonitorCount = 0
    
    EnumDisplayMonitors ByVal 0&, ByVal 0&, AddressOf MonitorEnumProc, ByVal 0&
End Sub

Public Sub CloseAllMonitorHandles()
    Dim i As Long
    
    For i = 0 To MonitorCount - 1
        '
        Debug.Print DestroyPhysicalMonitor(Monitors(i).pPhysicalInfo(0).hPhysicalMonitor)
        'PrintError
    Next i
    
    
    
End Sub

'Public Sub SetMonitorBrightness(nValue As Long)
'    If nValue < 0 Or nValue > 100 Then Exit Sub
'
'
'
'End Sub


Public Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As RECT, ByVal dwData As Long) As Long
    Dim i As Long
    ReDim Preserve Monitors(0 To MonitorCount) As MonitorStats
   
    Monitors(MonitorCount).lHandle = hMonitor
    Monitors(MonitorCount).lAllInfo.cbSize = Len(Monitors(MonitorCount).lAllInfo)
    
    GetMonitorInfo hMonitor, Monitors(MonitorCount).lAllInfo
    
    Dim lSettingmode As LPMC_VCP_CODE_TYPE
    Dim lCurrent As Long
    Dim lMaximum As Long

    GetNumberOfPhysicalMonitorsFromHMONITOR hMonitor, Monitors(MonitorCount).lMonitorCount
    
    If Monitors(MonitorCount).lMonitorCount > 0 Then
        ReDim Monitors(MonitorCount).pPhysicalInfo(0 To Monitors(MonitorCount).lMonitorCount - 1)
        
        GetPhysicalMonitorsFromHMONITOR hMonitor, Monitors(MonitorCount).lMonitorCount, Monitors(MonitorCount).pPhysicalInfo(0)
        
        Monitors(MonitorCount).lDescription = Replace(ByteArrayToString(Monitors(MonitorCount).pPhysicalInfo(0).szPhysicalMonitorDescription), Chr(0), "")

        For i = 0 To Monitors(MonitorCount).lMonitorCount - 1
            GetMonitorBrightness Monitors(MonitorCount).pPhysicalInfo(i).hPhysicalMonitor, Monitors(MonitorCount).lBrightnessMinimum, Monitors(MonitorCount).lBrightnessCurrent, Monitors(MonitorCount).lBrightnessMaximum
            'PrintError
            
            GetMonitorContrast Monitors(MonitorCount).pPhysicalInfo(i).hPhysicalMonitor, Monitors(MonitorCount).lContrastMinimum, Monitors(MonitorCount).lContrastCurrent, Monitors(MonitorCount).lContrastMaximum
            
            GetVCPFeatureAndVCPFeatureReply Monitors(MonitorCount).pPhysicalInfo(i).hPhysicalMonitor, &HD6, Monitors(MonitorCount).lPowerModeSetting, Monitors(MonitorCount).lPowerModeCurrent, Monitors(MonitorCount).lPowerModeMaximum
            
            'GetMonitorDisplayAreaSize Monitors(MonitorCount).pPhysicalInfo(i).hPhysicalMonitor, gWidth, Monitors(MonitorCount).lScreenWidthMinimum, Monitors(MonitorCount).lScreenWidthCurrent, Monitors(MonitorCount).lScreenWidthMaximum
            'GetMonitorDisplayAreaSize Monitors(MonitorCount).pPhysicalInfo(i).hPhysicalMonitor, gHeight, Monitors(MonitorCount).lScreenHeightMinimum, Monitors(MonitorCount).lScreenHeightCurrent, Monitors(MonitorCount).lScreenHeightMaximum
        Next i
        
    End If
    
    'GetMonitorBrightness hMonitor, j, k, l
    
    MonitorEnumProc = 1
    MonitorCount = MonitorCount + 1
    
    
'''
'''    Dim MI As MONITORINFO, R As RECT
'''    Debug.Print "Moitor handle: " + CStr(hMonitor)
'''    'initialize the MONITORINFO structure
'''    MI.cbSize = Len(MI)
'''    'Get the monitor information of the specified monitor
'''    GetMonitorInfo hMonitor, MI
'''    'write some information on teh debug window
'''    Debug.Print "Monitor Width/Height: " + CStr(MI.rcMonitor.Right - MI.rcMonitor.Left) + "x" + CStr(MI.rcMonitor.Bottom - MI.rcMonitor.Top)
'''    Debug.Print "Primary monitor: " + CStr(CBool(MI.dwFlags = MONITORINFOF_PRIMARY))
'''    'check whether Form1 is located on this monitor
'''    If MonitorFromWindow(frmMain.hwnd, MONITOR_DEFAULTTONEAREST) = hMonitor Then
'''        Debug.Print "Form1 is located on this monitor"
'''    End If
'''    'heck whether the point (0, 0) lies within the bounds of this monitor
'''    If MonitorFromPoint(0, 0, MONITOR_DEFAULTTONEAREST) = hMonitor Then
'''        Debug.Print "The point (0, 0) lies wihthin the range of this monitor..."
'''    End If
'''    'check whether Form1 is located on this monitor
'''    GetWindowRect frmMain.hwnd, R
'''    If MonitorFromRect(R, MONITOR_DEFAULTTONEAREST) = hMonitor Then
'''        Debug.Print "The rectangle of Form1 lies within this monitor"
'''    End If
'''    Debug.Print ""
'''    'Continue enumeration

    
End Function


Sub PrintError()
    Dim lngStatus As Long
    Dim lngErrorCode As Long
    Dim strMessage As String
    
    lngErrorCode = Err.LastDllError
    
    strMessage = Space$(512)
    
    lngStatus = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0, lngErrorCode, 0, strMessage, Len(strMessage), 0)
    
    Debug.Print "Error " & lngErrorCode & ": " & strMessage
End Sub
