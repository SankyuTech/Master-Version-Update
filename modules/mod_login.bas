Attribute VB_Name = "mod_login"
Sub frm302_login()
'On Error GoTo logging:
USER_FOUND = 0
If frm302.TB1(0) = vbNullString Then
    MsgBox "Sila Masukkan [USERNAME].", vbInformation, "Login"
    frm302.TB1(0).SetFocus
    Exit Sub
End If
If frm302.TB1(1) = vbNullString Then
    MsgBox "Sila Masukkan [PASSWORD].", vbInformation, "Login"
    frm302.TB1(1).SetFocus
    Exit Sub
End If
If InStr(1, frm302.TB1(0), "*") <> 0 Or InStr(1, frm302.TB1(0), "/") <> 0 Or InStr(1, frm302.TB1(0), "\") <> 0 Or InStr(1, frm302.TB1(0), "'") <> 0 Or InStr(1, frm302.TB1(0), "`") <> 0 Then
    MsgBox "[USERNAME] Mengandungi Simbol Yang Tidak Dibenarkan.", vbExclamation, "Login"
    frm302.TB1(0).SetFocus
    Exit Sub
End If
If InStr(1, frm302.TB1(1), "*") <> 0 Or InStr(1, frm302.TB1(1), "/") <> 0 Or InStr(1, frm302.TB1(1), "\") <> 0 Or InStr(1, frm302.TB1(1), "'") <> 0 Or InStr(1, frm302.TB1(1), "`") <> 0 Then
    MsgBox "[PASSWORD] Mengandungi Simbol Yang Tidak Dibenarkan.", vbExclamation, "Login"
    frm302.TB1(1).SetFocus
    Exit Sub
End If

MDI_frm1.L20_Text = vbNullString

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select username,password,status from login where username='" & frm302.TB1(0) & "' and password='" & frm302.TB1(1) & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Status = 1 Then
        If Not IsNull(rs!UserName) Then LM_USERNAME = rs!UserName
        
        USER_FOUND = 1
    Else
        MsgBox "Anda tidak dibenarkan akses ke dalam sistem ini." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Sila hubungi pihak admin sistem.", vbExclamation, "Error"
    End If
Else
    MsgBox "Akses tidak sah.", vbInformation, "Login Tidak Berjaya"
    frm302.TB1(0) = vbNullString
    frm302.TB1(1) = vbNullString
    frm302.TB1(0).SetFocus
End If

rs.Close
Set rs = Nothing

If USER_FOUND = 1 Then
    Call main_setting_system
    MDI_frm1.Caption = "[Master Update : " & G_VER_SYSTEM & "][" & G_SERVER_DATABASE & "][" & G_SYSTEM_TYPE & "] , Terminal : " & G_TERMINAL & " , User : " & LM_USERNAME
    MDI_frm1.Show
    
    Unload frm302
    MDI_frm1.L3_Text = LM_USERNAME 'User
End If

Exit Sub
logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm3 : cmdlogin_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
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
Sub frm302_reset_all()
'On Error GoTo logging:
For x = 0 To 1
    frm302.TB1(x) = vbNullString
Next x

Exit Sub
logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_login : frm302_reset_all" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
