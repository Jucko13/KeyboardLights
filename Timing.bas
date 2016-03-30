Attribute VB_Name = "Timing"
Option Explicit

Declare Function QueryPerformanceCounter Lib "Kernel32" _
                                 (X As Currency) As Boolean
Declare Function QueryPerformanceFrequency Lib "Kernel32" _
                                 (X As Currency) As Boolean
 
Type TimerData
    StartCount As Currency
    StopCount As Currency
    Overhead As Currency
    Frequency As Currency
End Type
 
'returns the number of seconds elapsed between StopCount and StartCount
'Assumes that InitPerformanceTimer has been called
Function ElapsedTime(MyData As TimerData) As Single
 With MyData
    ElapsedTime = (.StopCount - .StartCount - .Overhead) / .Frequency 'seconds
 End With
End Function
 
'Initializes the data structure so that the ElapsedTime function can be called
'Returns false if the High-resolution timer is not supported
Function InitPerformanceTimer(MyData As TimerData) As Boolean
    Dim Ctr1 As Currency, Ctr2 As Currency
    With MyData
      If QueryPerformanceCounter(Ctr1) Then
          QueryPerformanceCounter Ctr2
          QueryPerformanceFrequency .Frequency
          '1/Freq*10000 is resolution of timer in seconds
          'Debug.Print "QueryPerformanceCounter minimum resolution: 1/" & _
                      .Frequency * 10000; " seconds"
          'This overhead is present for each call to API
          .Overhead = Ctr2 - Ctr1
          'Debug.Print "API Overhead: "; .Overhead / .Frequency; "seconds"
          InitPerformanceTimer = True
        Else
          'Debug.Print "High-resolution counter not supported."
          InitPerformanceTimer = False
        End If
    End With
End Function

