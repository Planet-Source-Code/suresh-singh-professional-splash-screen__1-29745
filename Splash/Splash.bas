Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Function DelayTime(SecToDelay As Single)
Dim MarkTime As Single
MarkTime = GetTickCount
Do While GetTickCount < MarkTime + SecToDelay * 1000 'Convert into Seconds
DoEvents
Loop
End Function

