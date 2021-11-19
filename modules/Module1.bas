Attribute VB_Name = "Module1"
Public cn As ADODB.Connection
Public cn2 As ADODB.Connection
'Public rs As ADODB.Recordset
Public LogDate_Memory As String
Public LogAct_Memory As String
Public G_SERVER_IP
Public G_SERVER_USER
Public G_SERVER_PASS
Public G_SERVER_DATABASE
Public G_SERVER_PORT
Public G_TERMINAL
Sub db_connection_remote()
'On Error Resume Next
'### Set date format
Call SetDateTime

Set cn2 = New ADODB.Connection
cn2.ConnectionString = "Driver={MySQL ODBC 3.51 Driver}; Server=" & G_REMOTE_HOSTNAME & ";port=" & G_REMOTE_PORT & "; database=" & G_REMOTE_DATABASE & "; user=" & G_REMOTE_USERNAME & "; password=" & G_REMOTE_PASSWORD & "; option=3;"

cn2.Open

If cn2.State = adStateOpen Then
    'MsgBox "Connected"
Else
    MsgBox "Tiada connection antara sistem dan database , Sila pastikan XAMPP diaktifkan!", vbCritical, "Error"
    End
End If

MDI_frm1.L18_Text = "1"
End Sub
Sub Main()
'On Error Resume Next
'### Set date format
Call SetDateTime

Set cn = New ADODB.Connection
cn.ConnectionString = "Driver={MySQL ODBC 3.51 Driver}; Server=" & G_SERVER_IP & ";port=" & G_SERVER_PORT & "; database=" & G_SERVER_DATABASE & "; user=" & G_SERVER_USER & "; password=" & G_SERVER_PASS & "; option=3;"

cn.Open

If cn.State = adStateOpen Then
    'MsgBox "Connected"
Else
    MsgBox "Tiada connection antara sistem dan database , Sila pastikan XAMPP diaktifkan!", vbCritical, "Error"
    End
End If

MDI_frm1.L18_Text = "1"
End Sub
Sub Main2()
'On Error Resume Next
'### Set date format
Call SetDateTime

Set cn2 = New ADODB.Connection
cn2.ConnectionString = "Driver={MySQL ODBC 3.51 Driver}; Server=" & G_SERVER_IP & ";port=" & G_SERVER_PORT & "; database=" & G_DATABASE_SETTING & "; user=" & G_SERVER_USER & "; password=" & G_SERVER_PASS & "; option=3;"

cn2.Open

If cn2.State = adStateOpen Then
    'MsgBox "Connected"
Else
    MsgBox "Tiada connection antara sistem dan database , Sila pastikan XAMPP diaktifkan!", vbCritical, "Error"
    End
End If

MDI_frm1.L19_Text = "1"
End Sub
Sub check_db_conn_main()
'On Error GoTo logging:
LM_OPEN = 0

If G_SYSTEM_TYPE = "ONLINE" Then
    If MDI_frm1.L17_Text = "OFFLINE" Then
        MsgBox "Tiada sambungan internet. Sila pastikan komputer anda disambungkan dengan internet.", vbCritical, "Connection Failed"
        Exit Sub
    End If
End If

If MDI_frm1.L18_Text = "0" Then
    If MDI_frm1.L17_Text = "ONLINE" And G_SYSTEM_TYPE = "ONLINE" Then Call Main
End If

If G_SYSTEM_TYPE = "OFFLINE" Then Call Main

Exit Sub
logging:
G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Module1 : check_db_conn_main" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod
Resume Next
End Sub
Sub check_db_conn_main_remote()
'On Error GoTo logging:
LM_OPEN = 0

If G_SYSTEM_TYPE = "ONLINE" Then
    If MDI_frm1.L17_Text = "OFFLINE" Then
        MsgBox "Tiada sambungan internet. Sila pastikan komputer anda disambungkan dengan internet.", vbCritical, "Connection Failed"
        Exit Sub
    End If
End If

If MDI_frm1.L18_Text = "0" Then
    If MDI_frm1.L17_Text = "ONLINE" And G_SYSTEM_TYPE = "ONLINE" Then Call db_connection_remote
End If

If G_SYSTEM_TYPE = "OFFLINE" Then Call db_connection_remote

Exit Sub
logging:
G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Module1 : check_db_conn_main" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod
Resume Next
End Sub
Sub check_db_conn_main2()
'On Error GoTo logging:
If G_SYSTEM_TYPE = "ONLINE" Then

    If MDI_frm1.L17_Text = "OFFLINE" Then
        
        MsgBox "Tiada sambungan internet. Sila pastikan komputer anda disambungkan dengan internet.", vbCritical, "Connection Failed"
        
        Exit Sub
        
    End If

End If

If MDI_frm1.L22_Text = "0" Then
    If MDI_frm1.L17_Text = "ONLINE" And G_SYSTEM_TYPE = "ONLINE" Then Call Main2
End If

If G_SYSTEM_TYPE = "OFFLINE" Then Call Main2

Exit Sub
logging:
G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Module1 : check_db_conn_main2" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod
Resume Next
End Sub
Sub main_setting_system()
'On Error GoTo logging:
LM_FOUND = 0

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 1_info_sistem where setting ='" & "Sankyu" & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!tele_noti_error) Then LM_TOKEN_ERROR = rs!tele_noti_error
    If Not IsNull(rs!tele_noti_crash) Then LM_TOKEN_CRASH = rs!tele_noti_crash
    If Not IsNull(rs!tele_noti_dev) Then LM_TOKEN_DEV = rs!tele_noti_dev
    If Not IsNull(rs!tele_noti_success) Then LM_TOKEN_SUCCESS = rs!tele_noti_success
    If Not IsNull(rs!interval_reset) Then G_INTERVAL_RESET = rs!interval_reset
    If Not IsNull(rs!counter_crash) Then G_COUNTER_CRASH = rs!counter_crash
    If Not IsNull(rs!version_sistem) Then G_VER_SYSTEM = rs!version_sistem
    If Not IsNull(rs!tele_system_name) Then G_NAMA_KEDAI_TELE = rs!tele_system_name
    LM_FOUND = 1
End If

rs.Close
Set rs = Nothing

If LM_FOUND = 1 Then
    If G_VER_SYSTEM <> G_VERSION_CONTROL Then
        MsgBox "Sistem yang sedang digunakan adalah berbeza dengan sistem terbaru." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Version Ini : " & G_VERSION_CONTROL & vbCrLf & _
                "Version Terbaru : " & G_VER_SYSTEM & vbCrLf & _
                vbNullString & vbCrLf & _
                "Sila hubungi pihak Sankyu System untuk update database ini." & vbCrLf & _
                "Sistem ini akan ditutup.", vbCritical, "Database Out-Dated"
        End
    End If
End If
If InStr(1, LM_TOKEN_ERROR, "[]") <> 0 Then
    G_TOKEN_ERROR = Split(LM_TOKEN_ERROR, "[]")(0)
    G_CHAT_ERROR = Split(LM_TOKEN_ERROR, "[]")(1)
End If
If InStr(1, LM_TOKEN_CRASH, "[]") <> 0 Then
    G_TOKEN_CRASH = Split(LM_TOKEN_CRASH, "[]")(0)
    G_CHAT_CRASH = Split(LM_TOKEN_CRASH, "[]")(1)
End If
If InStr(1, LM_TOKEN_DEV, "[]") <> 0 Then
    G_TOKEN_UPDATE_DEV = Split(LM_TOKEN_DEV, "[]")(0)
    G_CHAT_UPDATE_DEV = Split(LM_TOKEN_DEV, "[]")(1)
End If
If InStr(1, LM_TOKEN_SUCCESS, "[]") <> 0 Then
    G_TOKEN_SUCCESS = Split(LM_TOKEN_SUCCESS, "[]")(0)
    G_CHAT_SUCCESS = Split(LM_TOKEN_SUCCESS, "[]")(1)
End If
MDI_frm1.Tmr4.Interval = G_INTERVAL_RESET

Exit Sub
logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Module1 : main_setting_system" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
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


