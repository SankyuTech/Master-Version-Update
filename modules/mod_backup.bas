Attribute VB_Name = "mod_backup"
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 
Private Const SYNCHRONIZE       As Long = &H100000
Private Const INFINITE          As Long = &HFFFF
Private Sub ExecCommand(ByVal cmd As String)
    Dim result  As Long
    Dim lPid    As Long
    Dim lHnd    As Long
    Dim lRet    As Long
 
    cmd = "cmd /c " & cmd
    result = Shell(cmd, vbHide)
 
    lPid = result
    If lPid <> 0 Then
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
        If lHnd <> 0 Then
            lRet = WaitForSingleObject(lHnd, INFINITE)
            CloseHandle (lHnd)
        End If
    End If
End Sub
Private Sub cmdBackup_Click()
Dim cmd As String
    Screen.MousePointer = vbHourglass
    DoEvents
 
    cmd = Chr(34) & "C:\Program Files\MySQL\MySQL Server 5.1\bin\mysqldump" & Chr(34) & " -uroot -psecretpswd --routines --comments db_name > c:\MyBackup.sql"
    Call ExecCommand(cmd)
 
    Screen.MousePointer = vbDefault
    MsgBox "done"
End Sub
Private Sub cmdRestore_Click()
Dim cmd As String
    Screen.MousePointer = vbHourglass
    DoEvents
 
    cmd = Chr(34) & "C:\Program Files\MySQL\MySQL Server 5.1\bin\mysql" & Chr(34) & " -uroot -psecretpswd --comments db_name < c:\MyBackup.sql"
    Call ExecCommand(cmd)
 
    Screen.MousePointer = vbDefault
    MsgBox "done"
End Sub
Sub backup_database()
'On Error Resume Next
Dim cmd As String
Dim LM_DATABASE_VER As String
Dim LM_USER As String
Dim LM_PASS As String
Dim LM_SYSTEM_VER As String
Dim LM_LINK As String
Dim LM_MODE As Integer

LM_DATABASE_VER = vbNullString
LM_USER = vbNullString
LM_PASS = vbNullString
LM_SYSTEM_VER = vbNullString
LM_LINK = vbNullString
LM_MODE = 0

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    
    If Not IsNull(rs!version_database) Then LM_DATABASE_VER = rs!version_database 'Version database
    If Not IsNull(rs!u_user) Then LM_USER = rs!u_user 'User
    If Not IsNull(rs!p_pass) Then LM_PASS = rs!p_pass 'Password
    If Not IsNull(rs!version_sistem) Then LM_SYSTEM_VER = rs!version_sistem 'Version Sistem
    If Not IsNull(rs!backup_link) Then LM_LINK = rs!backup_link 'Path folder yang akan disimpan database yang di backup
    If Not IsNull(rs!Mode) Then
        If rs!Mode = 0 Then
            LM_MODE = 0
        ElseIf rs!Mode = 1 Then
            LM_MODE = 1
        End If
    Else
        LM_MODE = 0
    End If
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

If LM_DATABASE_VER <> vbNullString And LM_USER <> vbNullString And LM_PASS <> vbNullString And LM_SYSTEM_VER <> vbNullString And LM_LINK <> vbNullString Then

    If LM_MODE = 0 Then
    
        MDI_frm1.Hide
        
        Screen.MousePointer = vbHourglass
        DoEvents
        
        cmd = Chr(34) & LM_LINK & "\mysqldump.exe -hlocalhost" & " -p" & LM_PASS & " -u" & LM_USER & " " & LM_DATABASE_VER & " > " & LM_LINK & "\" & LM_DATABASE_VER & "_" & LM_SYSTEM_VER & "_" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".sql"

        Call ExecCommand(cmd)
        
        Screen.MousePointer = vbDefault
        
        Note = "Backup database telah berjaya." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda ingin membuka folder yang mengandungi database ini?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
        If Answer = vbYes Then
            Dim myPath As String
            myPath = LM_LINK
            Shell "explorer " & myPath
        End If
        
        MDI_frm1.Show
    End If
End If
End Sub
