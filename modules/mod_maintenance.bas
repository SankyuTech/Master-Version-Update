Attribute VB_Name = "Module4"
Option Explicit

Private Const LOCALE_SDATE = &H1F
Private Const LOCALE_STIMEFORMAT = &H1003

Private Const WM_SETTINGCHANGE = &H1A

Private Const HWND_BROADCAST = &HFFFF&

Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Function SetDateTime() As Boolean
'On Error Resume Next
Dim dwLCID As Long

dwLCID = GetSystemDefaultLCID()

If SetLocaleInfo(dwLCID, LOCALE_SDATE, "yyyy-MM-dd") = False Then
SetDateTime = False
Exit Function
End If

If SetLocaleInfo(dwLCID, LOCALE_STIMEFORMAT, "HH:mm:ss") = False Then
SetDateTime = False
Exit Function
End If

PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0

SetDateTime = True

End Function
