Attribute VB_Name = "Module57"
'Working with registry declarations and constants
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Const ERROR_SUCCESS = 0&
Private Const APINULL = 0&
Private Const HKEY_LOCAL_MACHINE = &H80000002
'Working with wininet.dll declarations and constants
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Long, ByVal dwReserved As Long) As Long 'Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long 'this function used with IE4
'Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long 'this function used with IE4
Private Const INTERNET_CONNECTION_MODEM = &H1&
Private Const INTERNET_CONNECTION_LAN = &H2&
Private Const INTERNET_CONNECTION_PROXY = &H4&
Private Const INTERNET_RAS_INSTALLED = &H10&
Private Const INTERNET_CONNECTION_OFFLINE = &H20&
Private Const INTERNET_CONNECTION_CONFIGURED = &H40&
'Declares for direct ping
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInet As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long

Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Dim checkType As Integer
Dim remMsg(2) As String
Sub check_internet_dev(Optional ByRef ConnectionInfo As Long, Optional ByRef sConnectionName As String)
'On Error GoTo logging:
Dim dwFlags As Long
Dim sNameBuf As String, Msg As String
Dim lPos As Long
sNameBuf = String$(513, 0)

G_TOKEN_PASS = G_TOKEN_PASS_DEFAULT

If InternetGetConnectedStateEx(dwFlags, sNameBuf, 512, 0&) Then
    'Call generate_dev_pass
'online mode
Else
    G_TOKEN_PASS = G_TOKEN_PASS_DEFAULT
'offline mode
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Module1 : check_internet_dev" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
