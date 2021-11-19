Attribute VB_Name = "mod_internet_conn"
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
Sub Frm3_check_internet(Optional ByRef ConnectionInfo As Long, Optional ByRef sConnectionName As String)
'On Error Resume Next
Dim dwFlags As Long
Dim sNameBuf As String, Msg As String
Dim lPos As Long
sNameBuf = String$(513, 0)
If InternetGetConnectedStateEx(dwFlags, sNameBuf, 512, 0&) Then
'   lPos = InStr(sNameBuf, vbNullChar)
'   If lPos > 0 Then
'     sConnectionName = Left$(sNameBuf, lPos - 1)
'   Else
'     sConnectionName = ""
'   End If
'   msg = "Your computer is connected to Internet" & vbCrLf & "Connection Name: " & sConnectionName
'   If (dwFlags And INTERNET_CONNECTION_LAN) Then
'       msg = msg & vbCrLf & "Connection use LAN"
'   ElseIf lFlags And INTERNET_CONNECTION_MODEM Then
'       msg = msg & vbCrLf & "Connection use modem"
'   End If
'   If lFlags And INTERNET_CONNECTION_PROXY Then msg = msg & vbCrLf & "Connection use Proxy"
'   If lFlags And INTERNET_RAS_INSTALLED Then
'      msg = msg & vbCrLf & "RAS INSTALLED"
'   Else
'      msg = msg & vbCrLf & "RAS NOT INSTALLED"
'   End If
'   If lFlags And INTERNET_CONNECTION_OFFLINE Then
'      msg = msg & vbCrLf & "You are OFFLINE"
'   Else
'      msg = msg & vbCrLf & "You are ONLINE"
'   End If
'   If lFlags And INTERNET_CONNECTION_CONFIGURED Then
'      msg = msg & vbCrLf & "Your connection is Configured"
'   Else
'      msg = msg & vbCrLf & "Your connection is not Configured"
'   End If
Else
    Msg = "Komputer ini tidak disambungkan dengan internet." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sila pastikan internet connection anda."
    
    MsgBox Msg, vbExclamation, "Internet Connection"
    
    Exit Sub
End If
End Sub
Sub check_internet_connection_main(Optional ByRef ConnectionInfo As Long, Optional ByRef sConnectionName As String)
'On Error Resume Next
Dim dwFlags As Long
Dim sNameBuf As String, Msg As String
Dim lPos As Long
sNameBuf = String$(513, 0)
LM_CONN = 0 '0 : Offline , 1 : Online

If InternetGetConnectedStateEx(dwFlags, sNameBuf, 512, 0&) Then
'   lPos = InStr(sNameBuf, vbNullChar)
'   If lPos > 0 Then
'     sConnectionName = Left$(sNameBuf, lPos - 1)
'   Else
'     sConnectionName = ""
'   End If
'   msg = "Your computer is connected to Internet" & vbCrLf & "Connection Name: " & sConnectionName
'   If (dwFlags And INTERNET_CONNECTION_LAN) Then
'       msg = msg & vbCrLf & "Connection use LAN"
'   ElseIf lFlags And INTERNET_CONNECTION_MODEM Then
'       msg = msg & vbCrLf & "Connection use modem"
'   End If
'   If lFlags And INTERNET_CONNECTION_PROXY Then msg = msg & vbCrLf & "Connection use Proxy"
'   If lFlags And INTERNET_RAS_INSTALLED Then
'      msg = msg & vbCrLf & "RAS INSTALLED"
'   Else
'      msg = msg & vbCrLf & "RAS NOT INSTALLED"
'   End If
'   If lFlags And INTERNET_CONNECTION_OFFLINE Then
'      msg = msg & vbCrLf & "You are OFFLINE"
'   Else
'      msg = msg & vbCrLf & "You are ONLINE"
'   End If
'   If lFlags And INTERNET_CONNECTION_CONFIGURED Then
'      msg = msg & vbCrLf & "Your connection is Configured"
'   Else
'      msg = msg & vbCrLf & "Your connection is not Configured"
'   End If

    LM_CONN = 1 '0 : Offline , 1 : Online
    
Else

    LM_CONN = 0 '0 : Offline , 1 : Online
    
End If

    'LM_CONN = 1
    
If G_SYSTEM_TYPE = "ONLINE" Then

    If LM_CONN = 1 Then '0 : Offline , 1 : Online
    
        MDI_frm1.L17_Text = "ONLINE"
        MDI_frm1.Image1.Visible = True
        MDI_frm1.Image3.Visible = False
        
    Else
    
        MDI_frm1.L17_Text = "OFFLINE"
        MDI_frm1.L18_Text = "0"
        MDI_frm1.L19_Text = "0"
        MDI_frm1.L22_Text = "0"
        MDI_frm1.Image3.Visible = True
        MDI_frm1.Image1.Visible = False
        
    End If
    
Else
    
    'If LM_CONN = 1 Then '0 : Offline , 1 : Online
    '    MDI_frm1.L17_Text = "ONLINE"
    '    MDI_frm1.Image1.Visible = True
    '    MDI_frm1.Image2.Visible = False
    'Else
        MDI_frm1.L17_Text = "OFFLINE"
        MDI_frm1.Image3.Visible = True
        MDI_frm1.Image1.Visible = False
    'End If
    
    MDI_frm1.L18_Text = "0"
    MDI_frm1.L19_Text = "0"
    MDI_frm1.L22_Text = "0"
        
End If
End Sub
Sub check_internet_interval(Optional ByRef ConnectionInfo As Long, Optional ByRef sConnectionName As String)
'On Error Resume Next
Dim dwFlags As Long
Dim sNameBuf As String, Msg As String
Dim lPos As Long
sNameBuf = String$(513, 0)
LM_CONN = 0 '0 : Offline , 1 : Online

If InternetGetConnectedStateEx(dwFlags, sNameBuf, 512, 0&) Then
'   lPos = InStr(sNameBuf, vbNullChar)
'   If lPos > 0 Then
'     sConnectionName = Left$(sNameBuf, lPos - 1)
'   Else
'     sConnectionName = ""
'   End If
'   msg = "Your computer is connected to Internet" & vbCrLf & "Connection Name: " & sConnectionName
'   If (dwFlags And INTERNET_CONNECTION_LAN) Then
'       msg = msg & vbCrLf & "Connection use LAN"
'   ElseIf lFlags And INTERNET_CONNECTION_MODEM Then
'       msg = msg & vbCrLf & "Connection use modem"
'   End If
'   If lFlags And INTERNET_CONNECTION_PROXY Then msg = msg & vbCrLf & "Connection use Proxy"
'   If lFlags And INTERNET_RAS_INSTALLED Then
'      msg = msg & vbCrLf & "RAS INSTALLED"
'   Else
'      msg = msg & vbCrLf & "RAS NOT INSTALLED"
'   End If
'   If lFlags And INTERNET_CONNECTION_OFFLINE Then
'      msg = msg & vbCrLf & "You are OFFLINE"
'   Else
'      msg = msg & vbCrLf & "You are ONLINE"
'   End If
'   If lFlags And INTERNET_CONNECTION_CONFIGURED Then
'      msg = msg & vbCrLf & "Your connection is Configured"
'   Else
'      msg = msg & vbCrLf & "Your connection is not Configured"
'   End If

    LM_CONN = 1 '0 : Offline , 1 : Online
    
Else

    LM_CONN = 0 '0 : Offline , 1 : Online
    
End If

    'LM_CONN = 1
    
If G_SYSTEM_TYPE = "ONLINE" Then

    If LM_CONN = 1 Then '0 : Offline , 1 : Online
    
        MDI_frm1.L17_Text = "ONLINE"
        MDI_frm1.Image1.Visible = True
        MDI_frm1.Image3.Visible = False
        
    Else
    
        MDI_frm1.L17_Text = "OFFLINE"
        MDI_frm1.L18_Text = "0"
        MDI_frm1.L19_Text = "0"
        MDI_frm1.L22_Text = "0"
        MDI_frm1.Image3.Visible = True
        MDI_frm1.Image1.Visible = False
        
        Call MDI_frm1_unload_all_menu
        
    End If
    
Else
    
    If LM_CONN = 1 Then '0 : Offline , 1 : Online
        MDI_frm1.L17_Text = "ONLINE"
        MDI_frm1.Image1.Visible = True
        MDI_frm1.Image3.Visible = False
    Else
        MDI_frm1.L17_Text = "OFFLINE"
        MDI_frm1.Image3.Visible = True
        MDI_frm1.Image1.Visible = False
    End If
    
    MDI_frm1.L18_Text = "0"
    MDI_frm1.L19_Text = "0"
    MDI_frm1.L22_Text = "0"
        
End If
End Sub
Sub check_online_conn(Optional ByRef ConnectionInfo As Long, Optional ByRef sConnectionName As String)
'On Error GoTo logging:
Dim dwFlags As Long
Dim sNameBuf As String, Msg As String
Dim lPos As Long
sNameBuf = String$(513, 0)
LM_CONN = 0 '0 : Offline , 1 : Online

If InternetGetConnectedStateEx(dwFlags, sNameBuf, 512, 0&) Then
    LM_CONN = 1 '0 : Offline , 1 : Online
Else
    LM_CONN = 0 '0 : Offline , 1 : Online
End If

If MDI_frm1.L20_Text = "Semua cawangan" Then
    MsgBox "Sila buat pilihan cawangan terlebih dahulu.", vbExclamation, "Info"
    Exit Sub
End If

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
strsql = "UPDATE 91_senarai_terminal set cawangan='" & G_CAWANGAN & "' where terminal='" & G_TERMINAL & "'"

Set rs = cn.Execute(strsql)
Set rs = Nothing

If LM_CONN = 0 Then
    MsgBox "Tiada sambungan internet. Sila pastikan komputer anda ONLINE sebelum menghantar report ini.", vbExclamation, "Info"
Else
    If G_JENIS_REPORT = 0 Then '0 : Daily Report , 1 : Profit Report
        Shell "cmd.exe /c " & G_DAILY_REPORT_PATH
    ElseIf G_JENIS_REPORT = 1 Then
        Shell "cmd.exe /c " & G_PROFIT_REPORT_PATH
    End If
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_internet_conn : check_online_conn" & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If G_LM_ERR_NO = "3704" Or G_LM_ERR_NO = "-2147467259" Or G_LM_ERR_NO = "-2147217887" Then
    Call Main
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub
Sub check_connectivity_conn(Optional ByRef ConnectionInfo As Long, Optional ByRef sConnectionName As String)
'On Error Resume Next
Dim dwFlags As Long
Dim sNameBuf As String, Msg As String
Dim lPos As Long
sNameBuf = String$(513, 0)
LM_CONN = 0 '0 : Offline , 1 : Online

If InternetGetConnectedStateEx(dwFlags, sNameBuf, 512, 0&) Then
    LM_CONN = 1 '0 : Offline , 1 : Online
Else
    LM_CONN = 0 '0 : Offline , 1 : Online
End If

If LM_CONN = 0 Then
    MsgBox "Tiada sambungan internet. Sila pastikan komputer disambungkan dengan internet.", vbExclamation, "Info"
    
    End
End If
End Sub
