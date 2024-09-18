VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "timeStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function timeStrToMinute(ByVal timeStr As String) As Integer
    Dim timeValue As Date
    Dim totalMinutes As Integer
    
    timeValue = CDate(timeStr)
    totalMinutes = Hour(timeValue) * 60 + Minute(timeValue)
    
    timeStrToMinute = totalMinutes
End Function

Public Function MinuteToTimeStr(ByVal m As Integer) As String
    Dim timeStr As String
    
    timeStr = Format(Application.WorksheetFunction.RoundDown(m / 60, 0), "00") & ":" & Format(Application.WorksheetFunction.RoundDown(m Mod 60, 0), "00")
    
    MinuteToTimeStr = timeStr
End Function


