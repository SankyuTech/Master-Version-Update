Attribute VB_Name = "Module11"
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
Sub check_internet_connection(Optional ByRef ConnectionInfo As Long, Optional ByRef sConnectionName As String)
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

    Note = "Sistem berkemungkinan akan mengambil masa untuk menghantar email ini." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sistem akan menghantar email kepada pelanggan yang berdaftar dengan email yang sah sahaja." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan ?"
            
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

        'Frm98.Show
        'MDI_frm1.Hide
        'Frm98.Picture = MDI_frm1.Picture
        
        Call Frm97_send_email
    End If

Else
    Msg = "Komputer ini tidak disambungkan dengan internet." & vbCrLf & _
            "Oleh itu sistem tidak dapat menghantar e-mail." & vbCrLf & _
            "Sila periksa sambungan internet anda."
    
    MsgBox Msg, vbExclamation, "Internet Connection"
    
    Exit Sub
End If
   
End Sub
Public Function SendMail(sTo As String, sSubject As String, sFrom As String, _
    sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
    sSmtpUser As String, sSmtpPword As String, _
    sFilePath As String, bSmtpSSL As Boolean) As String
      
    'On Error GoTo SendMail_Error:
    Dim lobj_cdomsg      As CDO.Message
    Set lobj_cdomsg = New CDO.Message
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = sTo
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.Subject = sSubject
    lobj_cdomsg.TextBody = sBody
    If Trim$(sFilePath) <> vbNullString Then
        lobj_cdomsg.AddAttachment (sFilePath)
    End If
    lobj_cdomsg.Send
    Set lobj_cdomsg = Nothing
    SendMail = "ok"
    Exit Function
          
SendMail_Error:
    SendMail = Err.Description
End Function
Sub Frm97_Call_Promosi()
'On Error Resume Next
DATA_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 45_email_promosi where default_setting='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Frm97_LM_NAMA_KEDAI = rs!nama_kedai 'Nama Kedai
    If Not IsNull(rs!subjek) Then Frm97_LM_SUBJECT = rs!subjek 'Subject
    If Not IsNull(rs!Body) Then Frm97_LM_BODY = rs!Body 'E-mail Body
    
    DATA_FOUND = 1
End If

rs.Close
Set rs = Nothing

If DATA_FOUND = 1 Then
    Frm97.TB1 = Frm97_LM_SUBJECT
    Frm97.TB2 = Frm97_LM_BODY
End If
End Sub
Sub Frm97_Call_Promosi2()
'On Error Resume Next
DATA_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 45_email_promosi where default_setting='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Frm97_LM_NAMA_KEDAI = rs!nama_kedai 'Nama Kedai
    If Not IsNull(rs!subjek) Then Frm97_LM_SUBJECT = rs!subjek 'Subject
    If Not IsNull(rs!Body) Then Frm97_LM_BODY = rs!Body 'E-mail Body
    
    DATA_FOUND = 1
End If

rs.Close
Set rs = Nothing

If DATA_FOUND = 1 Then
    Frm97.L6_Text = "[" & Frm97_LM_NAMA_KEDAI & "] " & Frm97_LM_SUBJECT
    Frm97.L7_Text = Frm97_LM_BODY
    
    Frm97.L7_Text = "Assalamualaikum && Selamat Sejahtera." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Kepada Tuan/Puan/Encik/Cik *************" & vbCrLf & _
                vbNullString & vbCrLf & _
                Frm97_LM_BODY
End If
End Sub
Sub Frm97_send_email()
'On Error Resume Next
Dim A1 As String
Dim B1 As String
Dim C1 As String
Dim D1 As String
Dim Frm97_LM_SERVER As String
Dim Frm97_LM_USER As String
Dim Frm97_LM_PASSWORD As String
Dim Frm97_LM_PORT As Integer
Dim Frm97_LM_SSL As Boolean
Dim Frm97_LM_KEDAI As String
Dim Frm97_LM_KATEGORI As String
'Dim Frm97_LM_EMAIL As String

Dim retVal          As String
Dim objControl      As Control
x = 0

Frm98.Show
MDI_frm1.Hide
Frm98.Picture = MDI_frm1.Picture
        
'### Pilihan kategori pelanggan ### - Start
Frm97_LM_KATEGORI = InputBox("Sila pilih kategori pelanggan yang akan dihantar email." & _
        vbCrLf & "Sila masukkan nombor di bawah mengikut kategori pelanggan." & _
        vbCrLf & _
        vbCrLf & vbTab & "1 - Semua jenis pelanggan" & _
        vbCrLf & vbTab & "2 - Pelanggan biasa" & _
        vbCrLf & vbTab & "3 - Ahli Biasa" & _
        vbCrLf & vbTab & "4 - Silver" & _
        vbCrLf & vbTab & "5 - Gold" & _
        vbCrLf & vbTab & "6 - Platinum", "Pilihan kategori pelanggan")
         
Select Case Frm97_LM_KATEGORI
    Case "1"

        Frm97_LM_EMAIL = Null
        Frm97_LM_EMAIL_LOGIC = "<>"

    Case "2"
    
        Frm97_LM_EMAIL = 1
        Frm97_LM_EMAIL_LOGIC = "="
        
    Case "3"
    
        Frm97_LM_EMAIL = 2
        Frm97_LM_EMAIL_LOGIC = "="
        
    Case "4"

        Frm97_LM_EMAIL = 3
        Frm97_LM_EMAIL_LOGIC = "="
        
    Case "5"

        Frm97_LM_EMAIL = 4
        Frm97_LM_EMAIL_LOGIC = "="
        
    Case "6"

        Frm97_LM_EMAIL = 5
        Frm97_LM_EMAIL_LOGIC = "="
        
    Case "7"

        Frm97_LM_EMAIL = 6
        Frm97_LM_EMAIL_LOGIC = "="
        
    Case Else
        MsgBox "Tiada pilihan dibuat atau pilihan yang tidak sah.", vbInformation, "Info"
        
        Unload Frm98
        MDI_frm1.Show
        
        Exit Sub
        
End Select

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 45_email_promosi where default_setting='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Frm97_LM_KEDAI = rs!nama_kedai 'Nama Kedai
    If Not IsNull(rs!server) Then Frm97_LM_SERVER = rs!server 'Server
    If Not IsNull(rs!UserName) Then Frm97_LM_USER = rs!UserName 'Username
    If Not IsNull(rs!Password) Then Frm97_LM_PASSWORD = rs!Password 'Password
    If Not IsNull(rs!Port) Then Frm97_LM_PORT = rs!Port 'Port
    If Not IsNull(rs!flag_ssl) Then Frm97_LM_SSL = rs!flag_ssl 'SSL
    If Not IsNull(rs!subjek) Then Frm97_LM_SUBJECT = rs!subjek 'Subject
    If Not IsNull(rs!Body) Then Frm97_LM_BODY = rs!Body 'E-mail Body
End If

rs.Close
Set rs = Nothing

'### Carian E-mail Bagi Penerima Report ### - Start
' Add recipient email address
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from senarai_pelanggan where kategori_pelanggan " & Frm97_LM_EMAIL_LOGIC & "'" & Frm97_LM_EMAIL & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!Email) Then
        x = x + 1
        A1 = "Assalamualaikum & Selamat Sejahtera." & vbCrLf
        A1 = A1 & vbNullString & vbCrLf
        A1 = A1 & "Kepada Tuan/Puan/Encik/Cik " & rs!Nama & vbCrLf
        A1 = A1 & vbNullString & vbCrLf
        A1 = A1 & Frm97_LM_BODY & vbCrLf
        A1 = A1 & vbNullString & vbCrLf
        A1 = A1 & vbNullString & vbCrLf
        A1 = A1 & vbNullString & vbCrLf
        A1 = A1 & vbNullString & vbCrLf
        A1 = A1 & vbNullString & vbCrLf
        A1 = A1 & "Sistem ini dihasilkan oleh Sankyu System." & vbCrLf
        A1 = A1 & vbNullString & vbCrLf
        A1 = A1 & "Sebarang pertanyaan berkenaan sistem pengurusan kedai emas atau sistem perniagaan lain boleh hubungi Sankyu System (+6010 - 900 4788) , E-mail : sankyusystem@gmail.com" & vbCrLf
        A1 = A1 & "Facebook : Point Of Sales System"
        
        'cmdSend.Enabled = False
        retVal = SendMail(Trim$(rs!Email), _
            Trim$("[" & Frm97_LM_KEDAI & "] " & Frm97_LM_SUBJECT), _
            Trim$(Frm97_LM_KEDAI & " (Admin)") & "<" & Trim$(Frm97_LM_USER) & ">", _
            Trim$(A1), _
            Trim$(Frm97_LM_SERVER), _
            CInt(Trim$(Frm97_LM_PORT)), _
            Trim$(Frm97_LM_USER), _
            Trim$(Frm97_LM_PASSWORD), _
            Trim$(txtAttach), _
            CBool(Frm97_LM_SSL))
            
        'cmdSend.Enabled = True
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Carian E-mail Bagi Penerima Report ### - End

Unload Frm98
MDI_frm1.Show
    
MsgBox "Sistem telah berjaya menghantar email ini kepada " & x & " pelanggan.", vbInformation, "Info"
End Sub

