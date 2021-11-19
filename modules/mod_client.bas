Attribute VB_Name = "mod_client"
Sub frm1_initial_frame1()
'On Error GoTo logging:
For x = 0 To 0
    frm1.Frame1(x).Visible = False
    frm1.Frame1(x).Left = 1800
    frm1.Frame1(x).Top = 120
Next x

Exit Sub
logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_client : frm1_initial_frame1" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub frm1_initial_frame2()
'On Error GoTo logging:
For x = 0 To 1
    frm1.Frame2(x).Left = 240
    frm1.Frame2(x).Top = 1920
    frm1.Frame2(x).Visible = False
Next x

Exit Sub
logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_client : frm1_initial_frame2" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub frm1_initial_frame3()
'On Error GoTo logging:
frm1.Frame3.Left = 6600
frm1.Frame3.Top = 2280
frm1.Frame3.Visible = False
frm1.Frame3.ZOrder 1

Exit Sub
logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_client : frm1_initial_frame3" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub frm1_reset_l2()
'On Error GoTo logging:
For x = 0 To 3
    frm1.L2_Text(x) = vbNullString
Next x

Exit Sub
logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_client : frm1_reset_l2" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub frm1_senarai_version()
'On Error GoTo logging:
frm1.CBB1.Clear

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select version from 4_update_query where status = 1 order by version DESC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!Version) Then
        frm1.CBB1.AddItem rs!Version
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_client : frm1_senarai_version" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
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
Sub frm1_cek_krateria_report()
'On Error GoTo logging:
If frm1.TB1 <> vbNullString Then
    If InStr(1, frm1.TB1, "*") <> 0 Or InStr(1, frm1.TB1, "/") <> 0 Or InStr(1, frm1.TB1, "\") <> 0 Or InStr(1, frm1.TB1, "'") <> 0 Or InStr(1, frm1.TB1, "`") <> 0 Then
        MsgBox "[Keyword] mengandungi simbol yang tidak dibenarkan.", vbExclamation, "Info"
        Exit Sub
    End If
End If

Call frm1_carian_report

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_client : frm1_cek_krateria_report" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub frm1_carian_report()
'On Error GoTo logging:
If frm1.TB1 <> vbNullString Then
    LM_QUERY_1 = "(client LIKE '%" & frm1.TB1 & "%' OR token LIKE '%" & frm1.TB1 & "%' OR credential_1 LIKE '%" & frm1.TB1 & "%' OR credential_5 LIKE '%" & frm1.TB1 & "%')"
Else
    LM_QUERY_1 = "1 = 1"
End If

Call frm1_reset_paging_lv3

LM_QUERY_10 = "*"
LM_QUERY_11 = "COUNT(id)"
LM_QUERY_20 = " from 2_client_details where " & LM_QUERY_1 & " AND allow_update = 1 order by credential_1 ASC"

Erase G_QUERY_1
G_QUERY_1(0) = "select " & LM_QUERY_10 & LM_QUERY_20
G_QUERY_1(1) = "select " & LM_QUERY_11 & LM_QUERY_20

GM_NEXT_PREV = 0
Call frm1_senarai_client_header
Call frm1_senarai_client

If frm1.L104_Text(4) <> "Bil : 0" Then
    Call frm1_initial_frame2
    frm1.Frame2(1).Visible = True
Else
    MsgBox "Tiada rekod dijumpai.", vbInformation, "Info"
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_client : frm1_carian_report" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub frm1_senarai_client_header()
'On Error GoTo logging:
With frm1.LV3
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear
    frm1.LV3.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 2
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Nama Client", 2500
    .ColumnHeaders.Add 5, , "Hostname", 2000
    .ColumnHeaders.Add 6, , "Username", 1500
    .ColumnHeaders.Add 7, , "Password", 1500
    .ColumnHeaders.Add 8, , "Port", 1500
    .ColumnHeaders.Add 9, , "DB", 2300
    .ColumnHeaders.Add 10, , "Token", 2000
    .ColumnHeaders.Add 11, , "Version", 2000
End With

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_client : frm1_senarai_client_header" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub frm1_senarai_client()
'On Error GoTo logging:
Dim frm1_LM_TOTAL_PAGE As Double
frm1_LM_TOTAL_PAGE = 0
frm1_PAGE_SIZE = 22
x = 0

frm1.L104_Text(4) = "Bil : " & Format(0, "#,##0")

re_gen_report:

LM_START_ROW = frm1.L104_Text(2) 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm1_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If frm1.L104_Text(3) = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm1_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        frm1.L104_Text(0) = 1
    End If
End If

frm1_LM_PAGE_FOUND = 0

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open G_QUERY_1(0) & " LIMIT " & LM_START_ROW & "," & frm1_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm1_LM_PAGE_FOUND = 0 Then
        If frm1.L104_Text(3) = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                frm1.L104_Text(0) = frm1.L104_Text(0) + 1 'Paparan Page ke-xxx
                frm1_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(frm1.L104_Text(0)) Then
                    If frm1.L104_Text(0) <> 1 Then
                        frm1.L104_Text(0) = frm1.L104_Text(0) - 1 'Paparan Page ke-xxx
                        frm1_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    y = ((frm1.L104_Text(0) - 1) * frm1_PAGE_SIZE) + x

    With frm1.LV3.ListItems.Add(, , rs!ID)
        .ListSubItems.Add , , x
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID

        If Not IsNull(rs!client) Then
            .ListSubItems.Add , , rs!client
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!credential_1) Then
            .ListSubItems.Add , , rs!credential_1
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!credential_2) Then
            .ListSubItems.Add , , rs!credential_2
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!credential_4) Then
            .ListSubItems.Add , , rs!credential_4
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!credential_3) Then
            .ListSubItems.Add , , rs!credential_3
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!credential_5) Then
            .ListSubItems.Add , , rs!credential_5
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!token) Then
            .ListSubItems.Add , , rs!token
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!version_in_use) Then
            .ListSubItems.Add , , rs!version_in_use
        Else
            .ListSubItems.Add , , ""
        End If
    End With
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
LM_CONN = 2
re_conn_2:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open G_QUERY_1(1), cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs(0)) Then frm1_LM_TOTAL_PAGE = Format(rs(0) / frm1_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm1_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm1_LM_PAGE = Split(frm1_LM_TOTAL_PAGE, ".")(0)
        frm1_LM_PAGE_LEBIHAN = Split(frm1_LM_TOTAL_PAGE, ".")(1)
        
        If frm1_LM_PAGE_LEBIHAN <> "00" Then
            frm1.L104_Text(1) = frm1_LM_PAGE + 1
        Else
            frm1.L104_Text(1) = frm1_LM_PAGE
        End If
        
    Else
    
        frm1.L104_Text(1) = frm1_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        frm1.L104_Text(1) = 0
    End If
Else
    frm1.L104_Text(1) = 0
End If

If Not IsNull(rs(0)) Then frm1.L104_Text(4) = "Bil : " & Format(rs(0), "#,##0")

rs.Close
Set rs = Nothing

If x <> 0 Then
    frm1.L104_Text(2) = LM_START_ROW
End If

If frm1.L104_Text(0) <> vbNullString And IsNumeric(frm1.L104_Text(0)) Then
    If frm1.L104_Text(1) <> vbNullString And IsNumeric(frm1.L104_Text(1)) Then
    
        frm1_LM_CURR_PAGE = frm1.L104_Text(0)
        frm1_LM_TOTAL_PAGE = frm1.L104_Text(1)
        
        If frm1_LM_CURR_PAGE > frm1_LM_TOTAL_PAGE Then
            
            frm1.L104_Text(0) = frm1.L104_Text(0) - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
        
    End If
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_client : frm1_senarai_client" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If G_LM_ERR_NO = "3704" Or G_LM_ERR_NO = "-2147467259" Or G_LM_ERR_NO = "-2147217887" Then
    Call Main
    If LM_CONN = 1 Then
        Resume re_conn_1:
    ElseIf LM_CONN = 2 Then
        Resume re_conn_2:
    End If
Else
    Resume Next
End If
End Sub
Sub frm1_reset_paging_lv3()
'On Error GoTo logging:
frm1.L104_Text(2) = -1 'Titik Pencarian Data
frm1.L104_Text(3) = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
frm1.L104_Text(0) = 0 'Paparan Page ke-xxx
frm1.L104_Text(1) = 0

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_client : frm1_reset_paging_lv3" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub frm1_check_credential_remote()
'On Error GoTo logging:
LM_FOUND = 0

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select credential_1,credential_2,credential_3,credential_4,credential_5 from 2_client_details where credential_5='" & G_CLIENT_INFO(2) & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!credential_1) Then G_REMOTE_HOSTNAME = rs!credential_1 'Hostname
    If Not IsNull(rs!credential_2) Then G_REMOTE_USERNAME = rs!credential_2 'Username
    If Not IsNull(rs!credential_4) Then G_REMOTE_PASSWORD = rs!credential_4 'Password
    If Not IsNull(rs!credential_3) Then G_REMOTE_PORT = rs!credential_3 'Port
    If Not IsNull(rs!credential_5) Then G_REMOTE_DATABASE = rs!credential_5 'Database
    LM_FOUND = 1
Else
    MsgBox "Tiada data dijumpai.", vbExclamation, "Info"
End If

rs.Close
Set rs = Nothing

If LM_FOUND = 1 Then
    Call db_connection_remote
    Call frm1_check_version_terkini
Else
    MsgBox "Tiada data", vbInformation, "Info"
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_client : frm1_check_credential_remote" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
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
Sub frm1_check_version_sedang_digunakan()
'On Error GoTo logging:
Dim rs As ADODB.Recordset
LM_FOUND = 0

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select credential_1,credential_2,credential_3,credential_4,credential_5 from 2_client_details where credential_5='" & G_CLIENT_INFO(2) & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!credential_1) Then G_REMOTE_HOSTNAME = rs!credential_1 'Hostname
    If Not IsNull(rs!credential_2) Then G_REMOTE_USERNAME = rs!credential_2 'Username
    If Not IsNull(rs!credential_4) Then G_REMOTE_PASSWORD = rs!credential_4 'Password
    If Not IsNull(rs!credential_3) Then G_REMOTE_PORT = rs!credential_3 'Port
    If Not IsNull(rs!credential_5) Then G_REMOTE_DATABASE = rs!credential_5 'Database
    LM_FOUND = 1
Else
    MsgBox "Tiada data dijumpai.", vbExclamation, "Info"
End If

rs.Close
Set rs = Nothing

If LM_FOUND = 1 Then
    Call db_connection_remote

LM_CONN = 2
re_conn_2:
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main_remote Else Exit Sub
    rs.Open "select * from 56_maklumat_kedai order by cawangan ASC", cn2, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        If Not IsNull(rs!version_sistem) And Not IsNull(rs!cawangan) Then
            LM_SENARAI_VERSION = LM_SENARAI_VERSION & rs!cawangan & "  --->>> " & rs!version_sistem & vbCrLf
        End If
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    
    MsgBox LM_SENARAI_VERSION, vbInformation, "Info"
Else
    MsgBox "Tiada data", vbInformation, "Info"
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_client : frm1_check_version_sedang_digunakan" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If G_LM_ERR_NO = "3704" Or G_LM_ERR_NO = "-2147467259" Or G_LM_ERR_NO = "-2147217887" Then
    If LM_CONN = 1 Then
        Call Main
        Resume re_conn_1:
    ElseIf LM_CONN = 2 Then
        Call db_connection_remote
        Resume re_conn_2:
    End If
Else
    Resume Next
End If
End Sub
Sub frm1_check_version_terkini()
'On Error GoTo logging:
Dim LM_VERSION_ASAL As Integer
Dim LM_VERSION_TERKINI As Integer

LM_FOUND = 0

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main_remote Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan<>'" & "HQ" & "'", cn2, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!version_sistem) Then G_REMOTE_VERSION_ASAL = rs!version_sistem
    LM_FOUND = 1
End If

rs.Close
Set rs = Nothing

If LM_FOUND = 1 Then
    
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
        MsgBox "Pilihan version yang tidak sah.", vbExclamation, "Info"
        Exit Sub
    End If
    
    G_TOKEN_PASS = vbNullString
    Note = "Sistem client akan diupdate dari version " & G_REMOTE_VERSION_ASAL & " kepada " & frm1.CBB1 & "." & vbCrLf & _
            "Teruskan?"
            
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        Call start_update_database
    End If
Else
    MsgBox "Tiada data", vbInformation, "Info"
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm1 : frm1_check_version_terkini" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
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
