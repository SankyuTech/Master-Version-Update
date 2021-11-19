Attribute VB_Name = "Module3"
Sub UpdateLog_Database()
'On Error GoTo logging:
Dim rs As ADODB.Recordset

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
strsql = "insert into log(client,version,log_detail,terminal,username,write_timestamp)" & _
            "select '" & G_LOG(0) & "','" & G_LOG(1) & "','" & G_LOG(2) & "','" & G_LOG(3) & "','" & G_LOG(4) & "','" & G_LOG(5) & "'"

Set rs = cn.Execute(strsql)
Set rs = Nothing

LogDate_Memory = vbNullString
LogAct_Memory = vbNullString

Exit Sub
logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Module3 : UpdateLog_Database" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
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
