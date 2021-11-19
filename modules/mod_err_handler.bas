Attribute VB_Name = "mod_err_handler"
Sub carian_error_shutdown()
'On Error resume next
LM_SERVER = Split(G_CREDENTIAL_ERR, "/")(0)
LM_USER = Split(G_CREDENTIAL_ERR, "/")(1)
LM_PASSWORD = Split(G_CREDENTIAL_ERR, "/")(2)
LM_PORT = Split(G_CREDENTIAL_ERR, "/")(3)
LM_DATABASE = Split(G_CREDENTIAL_ERR, "/")(4)

Call SetDateTime

Set cn = New ADODB.Connection
cn.ConnectionString = "Driver={MySQL ODBC 3.51 Driver}; Server=" & LM_SERVER & ";port=" & LM_PORT & "; database=" & LM_DATABASE & "; user=" & LM_USER & "; password=" & LM_PASSWORD & "; option=3;"

cn.Open

If cn.State = adStateOpen Then
    'MsgBox "Connected"
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 1_error_list where err_number='" & G_LM_ERR_NO & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!Action) Then G_ERR_ACTION = "Tidakan yang mungkin boleh diambil : " & vbCrLf & _
                                                    vbNullString & vbCrLf & _
                                                    rs!Action
Else
    G_ERR_ACTION = "-- Shutdown --"
End If

rs.Close
Set rs = Nothing
End Sub
