Attribute VB_Name = "mod_update_sistem"
Sub start_update_database()
'On Error GoTo logging:
Dim rs As ADODB.Recordset
Dim LM_VERSION_ASAL As Integer
Dim LM_VERSION_TERKINI As Integer
Dim LM_CURRENT_VERSION As Integer

'Verify version
If InStr(1, G_REMOTE_VERSION_ASAL, ".") <> 0 Then
    LM_VERSION_ASAL = Split(G_REMOTE_VERSION_ASAL, ".")(2)
Else
    MsgBox "Version asal yang tidak sah.", vbExclamation, "Info"
    Exit Sub
End If
If InStr(1, frm1.CBB1, ".") <> 0 Then
    LM_VERSION_TERKINI = Split(frm1.CBB1, ".")(2)
Else
    MsgBox "Version terkini yang tidak sah.", vbExclamation, "Info"
    Exit Sub
End If
If LM_VERSION_ASAL >= LM_VERSION_TERKINI Then
    MsgBox "Pilihan version yang tidak sah." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Client ini sedang menggunakan version " & LM_VERSION_ASAL & ".", vbExclamation, "Info"
    Exit Sub
End If

LM_CURRENT_VERSION = LM_VERSION_ASAL + 1

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 4_update_query where seq_version BETWEEN '" & LM_CURRENT_VERSION & "' AND '" & LM_VERSION_TERKINI & "' AND status = 1 order by seq_version ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!Version) Then G_VERSION_UPDATE = rs!Version
    G_UPDATE_QUERY = vbNullString
    If Not IsNull(rs!Query) Then G_UPDATE_QUERY = rs!Query
    Call update_version_terkini
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Call frm1_initial_frame3

GM_NEXT_PREV = 2 '0 : Next , 1 : Previous
Call frm1_senarai_client_header
Call frm1_senarai_client
MsgBox "Sistem telah berjaya dikemaskini.", vbInformation, "Info"

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_update_sistem : start_update_database" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
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
Sub update_version_terkini()
'On Error GoTo logging:
Dim rs As ADODB.Recordset

LM_NOW = Now

G_LOG(0) = G_CLIENT_INFO(0)
G_LOG(1) = G_VERSION_UPDATE
G_LOG(2) = "Start Update Version [" & G_VERSION_UPDATE & "]."
G_LOG(3) = G_TERMINAL
G_LOG(4) = MDI_frm1.L3_Text
G_LOG(5) = LM_NOW

LogDate_Memory = LM_NOW
Call UpdateLog_Database
        
If G_UPDATE_QUERY <> vbNullString Then
    i = UBound(Split(G_UPDATE_QUERY, ";"))
    
    For x = 1 To i
LM_CONN = 1
re_conn_1:
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main_remote Else Exit Sub
        strsql = Split(G_UPDATE_QUERY, ";")(x - 1)
        
        Set rs = cn2.Execute(strsql)
        Set rs = Nothing
        
        LM_FOUND = 1
    Next x
End If

'Mulakan additional update
If G_VERSION_UPDATE = "SPKE2.2.11" Then
    Call additional_update_version_2211
ElseIf G_VERSION_UPDATE = "SPKE2.2.17" Then
    MsgBox "Sila pastikan anda memasukkan table 145_disclaimer secara manual ke dalam database client ini.", vbInformation, "Info"
ElseIf G_VERSION_UPDATE = "SPKE2.2.15" Then
    'Call additional_update_version_2215
ElseIf G_VERSION_UPDATE = "SPKE2.2.18" Then
    'Call additional_update_version_2218
ElseIf G_VERSION_UPDATE = "SPKE2.2.19" Then
    'Call additional_update_version_2219
ElseIf G_VERSION_UPDATE = "SPKE2.2.23" Then
    Call additional_update_version_2223
ElseIf G_VERSION_UPDATE = "SPKE2.2.35" Then
    'Call additional_update_version_2235
ElseIf G_VERSION_UPDATE = "SPKE2.2.38" Then
    Call additional_update_version_2238
ElseIf G_VERSION_UPDATE = "SPKE2.2.44" Then
    Call additional_update_version_2244
End If
''/// Additional Update

LM_CONN = 2
re_conn_2:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main_remote Else Exit Sub
strsql = "UPDATE 56_maklumat_kedai set version_sistem='" & G_VERSION_UPDATE & "' , update_source_file='" & "http://159.65.143.104/auto_update/" & G_VERSION_UPDATE & "/Sistem Pengurusan Kedai Emas (Sankyu System).exe" & "'"

Set rs = cn2.Execute(strsql)
Set rs = Nothing

LM_NOW = Now
G_LOG(2) = "Finish Update Version [" & G_VERSION_UPDATE & "]."
G_LOG(5) = LM_NOW
Call UpdateLog_Database

LM_CONN = 3
re_conn_3:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
strsql = "UPDATE 2_client_details set version_in_use='" & G_VERSION_UPDATE & "' where token='" & G_CLIENT_INFO(4) & "'"

Set rs = cn.Execute(strsql)
Set rs = Nothing

Call send_successful_notification

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_update_sistem : update_version_terkini" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If G_LM_ERR_NO = "3704" Or G_LM_ERR_NO = "-2147467259" Or G_LM_ERR_NO = "-2147217887" Then
    Call db_connection_remote
    If LM_CONN = 1 Then
        Resume re_conn_1:
    ElseIf LM_CONN = 2 Then
        Resume re_conn_2:
    End If
Else
    Resume Next
End If
End Sub
Sub frm1_start_update_all_system()
'On Error GoTo logging:
Dim rs As ADODB.Recordset
        
LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 2_client_details where status = 1 AND allow_update = 1 order by version_in_use DESC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Erase G_CLIENT_INFO
    If Not IsNull(rs!client) Then G_CLIENT_INFO(0) = rs!client
    If Not IsNull(rs!credential_1) Then G_CLIENT_INFO(1) = rs!credential_1
    If Not IsNull(rs!credential_5) Then G_CLIENT_INFO(2) = rs!credential_5
    If Not IsNull(rs!version_in_use) Then G_CLIENT_INFO(3) = rs!version_in_use
    If Not IsNull(rs!token) Then G_CLIENT_INFO(4) = rs!token
    
    Call frm1_check_credential_remote
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

MsgBox "Selesai.", vbInformation, "Info"

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_update_sistem : start_update_database" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
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
Sub additional_update_version_2211()
'On Error GoTo logging:
Dim rs As ADODB.Recordset

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main_remote Else Exit Sub
strsql = "UPDATE default_setting SET config_komisen2='" & "0,10,1000,20" & "'"

Set rs = cn2.Execute(strsql)
Set rs = Nothing

LM_NOW = Now
G_LOG(2) = "End Additional Update Version [" & G_VERSION_UPDATE & "]."
G_LOG(5) = LM_NOW
Call UpdateLog_Database

Call send_successful_notification

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_update_sistem : additional_update_version_2211" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If G_LM_ERR_NO = "3704" Or G_LM_ERR_NO = "-2147467259" Or G_LM_ERR_NO = "-2147217887" Then
    Call db_connection_remote
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub
Sub additional_update_version_2215()
'On Error GoTo logging:
Dim rs As ADODB.Recordset

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main_remote Else Exit Sub
rs.Open "select * from 56_maklumat_kedai", cn2, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!config_sys) Then LM_CONFIG = rs!config_sys
    
    rs!config_sys = LM_CONFIG & "[]0,"
    rs.Update
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

LM_NOW = Now
G_LOG(2) = "End Additional Update Version [" & G_VERSION_UPDATE & "]."
G_LOG(5) = LM_NOW
Call UpdateLog_Database

Call send_successful_notification

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_update_sistem : additional_update_version_2215" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If G_LM_ERR_NO = "3704" Or G_LM_ERR_NO = "-2147467259" Or G_LM_ERR_NO = "-2147217887" Then
    Call db_connection_remote
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub
Sub additional_update_version_2219()
'On Error GoTo logging:
Dim rs As ADODB.Recordset

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main_remote Else Exit Sub
rs.Open "select * from 56_maklumat_kedai", cn2, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!config_sys) Then LM_CONFIG = rs!config_sys
    
    rs!config_sys = LM_CONFIG & "[]0,"
    rs.Update
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

LM_NOW = Now
G_LOG(2) = "End Additional Update Version [" & G_VERSION_UPDATE & "]."
G_LOG(5) = LM_NOW
Call UpdateLog_Database

Call send_successful_notification

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_update_sistem : additional_update_version_2219" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If G_LM_ERR_NO = "3704" Or G_LM_ERR_NO = "-2147467259" Or G_LM_ERR_NO = "-2147217887" Then
    Call db_connection_remote
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub
Sub additional_update_version_2218()
'On Error GoTo logging:
Dim rs As ADODB.Recordset
Dim LM_LIST_CAWANGAN(10)
Dim x As Integer

Erase LM_LIST_CAWANGAN
x = 0

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main_remote Else Exit Sub
rs.Open "select * from 56_maklumat_kedai", cn2, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Not IsNull(rs!cawangan) Then LM_LIST_CAWANGAN(x) = rs!cawangan
    If Not IsNull(rs!config_sys) Then LM_CONFIG = rs!config_sys
    
    rs!config_sys = LM_CONFIG & "[]0,"
    rs.Update
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

If x <> 0 Then
    For y = 0 To x
LM_CONN = 2
re_conn_2:
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main_remote Else Exit Sub
        strsql = "insert into 146_maklumat_system(cawangan,nama_sistem,caption_excel,caption_contact_us,)" & _
                    "select '" & LM_LIST_CAWANGAN(y) & "','" & "Sankyu System" & "','" & "Report Generated By Sankyu System" & "','" & "Sankyu System , +6010 - 900 4788 , info@sankyutech.com" & "'"
        
        Set rs = cn2.Execute(strsql)
        Set rs = Nothing
    Next y
End If

LM_NOW = Now
G_LOG(2) = "End Additional Update Version [" & G_VERSION_UPDATE & "]."
G_LOG(5) = LM_NOW
Call UpdateLog_Database

Call send_successful_notification

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_update_sistem : additional_update_version_2218" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If G_LM_ERR_NO = "3704" Or G_LM_ERR_NO = "-2147467259" Or G_LM_ERR_NO = "-2147217887" Then
    Call db_connection_remote
    If LM_CONN = 1 Then
        Resume re_conn_1:
    ElseIf LM_CONN = 2 Then
        Resume re_conn_2:
    End If
Else
    Resume Next
End If
End Sub

Sub additional_update_version_2223()
'On Error GoTo logging:
Dim rs As ADODB.Recordset

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main_remote Else Exit Sub
rs.Open "select * from default_setting", cn2, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!config_komisen) Then LM_CONFIG = rs!config_komisen
    
    rs!config_komisen = LM_CONFIG & ",[Sama]0,[Asing]1,[Ya]0,[Tidak]1,[Berat]0.00"
    rs.Update
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

LM_NOW = Now
G_LOG(2) = "End Additional Update Version [" & G_VERSION_UPDATE & "]."
G_LOG(5) = LM_NOW
Call UpdateLog_Database

Call send_successful_notification

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_update_sistem : additional_update_version_2223" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If G_LM_ERR_NO = "3704" Or G_LM_ERR_NO = "-2147467259" Or G_LM_ERR_NO = "-2147217887" Then
    Call db_connection_remote
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub
Sub additional_update_version_2235()
'On Error GoTo logging:
Dim rs As ADODB.Recordset

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main_remote Else Exit Sub
rs.Open "select * from 56_maklumat_kedai", cn2, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!config_sys) Then LM_CONFIG = rs!config_sys
    
    rs!config_sys = LM_CONFIG & "[]0,"
    rs.Update
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

LM_NOW = Now
G_LOG(2) = "End Additional Update Version [" & G_VERSION_UPDATE & "]."
G_LOG(5) = LM_NOW
Call UpdateLog_Database

Call send_successful_notification

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_update_sistem : additional_update_version_2235" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If G_LM_ERR_NO = "3704" Or G_LM_ERR_NO = "-2147467259" Or G_LM_ERR_NO = "-2147217887" Then
    Call db_connection_remote
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub
Sub additional_update_version_2238()
'On Error GoTo logging:
Dim rs As ADODB.Recordset

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main_remote Else Exit Sub
rs.Open "select * from 56_maklumat_kedai", cn2, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!config_sys2) Then LM_CONFIG = rs!config_sys2
    
    rs!config_sys2 = LM_CONFIG & "[]0,"
    rs.Update
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

LM_NOW = Now
G_LOG(2) = "End Additional Update Version [" & G_VERSION_UPDATE & "]."
G_LOG(5) = LM_NOW
Call UpdateLog_Database

Call send_successful_notification

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_update_sistem : additional_update_version_2238" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If G_LM_ERR_NO = "3704" Or G_LM_ERR_NO = "-2147467259" Or G_LM_ERR_NO = "-2147217887" Then
    Call db_connection_remote
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub

Sub additional_update_version_2244()
'On Error GoTo logging:
Dim rs As ADODB.Recordset
Dim config_to_add As String
Dim LAST_CHAR_CURRENT_CONFIG As Double
Dim LM_LATEST_CONFIG As String

LAST_CHAR_CURRENT_CONFIG = 0
LM_LATEST_CONFIG = 0

LM_LATEST_CONFIG = "[]0,[]0,[]0,[]0,[]1,[]1,[]0,[]0,[]10,[]0,[]1,[]0,[]2,[]0,[]10,[]https://www.sankyutechnology.com,[]0,[]0,[]0,[]1,[]1,[]0,[]0,[]0,[]0,[]0,[]0,[]0,[]1,[]0,[]0,[]0,[]1,[]0,[]1,[]0,[]0,[]0,[]0,"

LAST_CHAR_LATEST_CONFIG = InStrRev(LM_LATEST_CONFIG, ",", , vbTextCompare)

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main_remote Else Exit Sub
rs.Open "select * from 56_maklumat_kedai", cn2, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    config_to_add = ""
    LAST_CHAR_CURRENT_CONFIG = 0
    
    If Not IsNull(rs!config_sys) Then
        LAST_CHAR_CURRENT_CONFIG = InStrRev(rs!config_sys, ",", , vbTextCompare)
        LM_CONFIG = rs!config_sys
    End If
    If LAST_CHAR_LATEST_CONFIG - LAST_CHAR_CURRENT_CONFIG > 0 Then
        config_to_add = Right(LM_LATEST_CONFIG, LAST_CHAR_LATEST_CONFIG - LAST_CHAR_CURRENT_CONFIG)
        rs!config_sys = LM_CONFIG & config_to_add
        rs.Update
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

LM_NOW = Now
G_LOG(2) = "End Additional Update Version [" & G_VERSION_UPDATE & "]."
G_LOG(5) = LM_NOW
Call UpdateLog_Database

Call send_successful_notification

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_update_sistem : additional_update_version_2244" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If G_LM_ERR_NO = "3704" Or G_LM_ERR_NO = "-2147467259" Or G_LM_ERR_NO = "-2147217887" Then
    Call db_connection_remote
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub
