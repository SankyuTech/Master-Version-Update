Attribute VB_Name = "telegram"
Sub log_rekod()
'On Error Resume Next
Dim sFilename As String
sFilename = App.Path & "\logging.txt"

' Archive file at certain size
If FileLen(sFilename) > 200000 Then
    FileCopy sFilename _
        , Replace(sFilename, ".txt", Format(Now, "ddmmyyyy hhmmss.txt"))
    Kill sFilename
End If

' Open the file to write
Dim filenumber As Variant
filenumber = FreeFile
Open sFilename For Append As #filenumber

Print #filenumber, G_ERROR_NAIYO

Close #filenumber

Call send_report_tele
End Sub
Sub send_report_tele()
'On Error Resume Next
Dim A1 As String
Dim File_Path As String
Dim LM_LIMIT As Integer
Dim LM_ERR_DETAIL As String

LM_ERR_DETAIL = G_ERROR_NAIYO & vbCrLf & _
                Err.Description & vbCrLf & _
                G_VER_SYSTEM & " : " & G_VER_DATABASE


LM_LIMIT = G_COUNTER_CRASH

G_X = G_X + 1

If G_X >= LM_LIMIT Then

    App.LogEvent LM_ERR_DETAIL, vbLogEventTypeError
    
    A1 = A1 & "Error detected : Crash" & "%0A" & vbCrLf
    A1 = A1 & vbNullString & vbCrLf
    A1 = A1 & "System Version : " & G_VERSION_CONTROL & "%0A" & vbCrLf
    A1 = A1 & "Station : " & G_TERMINAL & "%0A" & vbCrLf
    A1 = A1 & vbNullString & vbCrLf
    A1 = A1 & "Client : " & G_CLIENT_INFO(0) & "%0A" & vbCrLf
    A1 = A1 & "Database : " & G_CLIENT_INFO(2) & "%0A" & vbCrLf
    A1 = A1 & "Version Update : " & G_VERSION_UPDATE & "%0A" & vbCrLf
    A1 = A1 & "Error : " & G_ERROR_NAIYO & "%0A" & vbCrLf
    A1 = A1 & "Error Number : " & G_LM_ERR_NO & "%0A" & vbCrLf
    A1 = A1 & "Error Description : " & Err.Description & "%0A" & vbCrLf

    strURL = G_TOKEN_CRASH & "/" & "sendmessage?chat_id=" & G_CHAT_CRASH & "&text=" & A1 & ""

    Set XMLHttpRequest = New MSXML2.XMLHTTP
    XMLHttpRequest.Open "GET", strURL, False
    XMLHttpRequest.Send
    
    strResponse = XMLHttpRequest.responseText
    Set XMLHttpRequest = Nothing
    
    LM_ERR_NAIYO = G_ERROR_NAIYO & vbCrLf & _
                    vbNullString & vbCrLf & _
                    G_ERR_ACTION
    
    MsgBox "Telah berlaku error di dalam sistem." & vbCrLf & _
            vbNullString & vbCrLf & _
            ">>> " & LM_ERR_NAIYO & " <<<" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sila hubungi pihak Sankyu System jika error ini berulang." & vbCrLf & _
            "Sistem akan ditutup.", vbCritical, "Error"
    End

End If

App.LogEvent LM_ERR_DETAIL, vbLogEventTypeWarning

A1 = A1 & "Error detected : Temporary" & "%0A" & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & "System Version : " & G_VERSION_CONTROL & "%0A" & vbCrLf
A1 = A1 & "Station : " & G_TERMINAL & "%0A" & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & "Client : " & G_CLIENT_INFO(0) & "%0A" & vbCrLf
A1 = A1 & "Database : " & G_CLIENT_INFO(2) & "%0A" & vbCrLf
A1 = A1 & "Version Update : " & G_VERSION_UPDATE & "%0A" & vbCrLf
A1 = A1 & "Error : " & G_ERROR_NAIYO & "%0A" & vbCrLf
A1 = A1 & "Error Number : " & G_LM_ERR_NO & "%0A" & vbCrLf
A1 = A1 & "Error Description : " & Err.Description & "%0A" & vbCrLf

strURL = G_TOKEN_ERROR & "/" & "sendmessage?chat_id=" & G_CHAT_ERROR & "&text=" & A1 & ""

Set XMLHttpRequest = New MSXML2.XMLHTTP
XMLHttpRequest.Open "GET", strURL, False
XMLHttpRequest.Send

strResponse = XMLHttpRequest.responseText
Set XMLHttpRequest = Nothing
End Sub
Sub send_developer_token()
'On Error GoTo logging:
Dim A1 As String
Dim File_Path As String
Dim LM_LIMIT As Integer
Dim LM_ERR_DETAIL As String

App.LogEvent "Request Developer Pass", vbLogEventTypeInformation

A1 = A1 & "**Request Developer Pass**" & "%0A" & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & "Sistem : " & G_VERSION_CONTROL & "%0A" & vbCrLf
A1 = A1 & "Station : " & G_TERMINAL & "%0A" & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & "Client : " & G_CLIENT_INFO(0) & "%0A" & vbCrLf
A1 = A1 & "Database : " & G_CLIENT_INFO(2) & "%0A" & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & "Pass : **" & G_TOKEN_PASS & "**" & "%0A" & vbCrLf

strURL = G_TOKEN_UPDATE_DEV & "/" & "sendmessage?chat_id=" & G_CHAT_UPDATE_DEV & "&text=" & A1 & ""
    
Set XMLHttpRequest = New MSXML2.XMLHTTP
XMLHttpRequest.Open "GET", strURL, False
'XMLHttpRequest.setRequestHeader "Content-Type", "text/xml"
XMLHttpRequest.Send

strResponse = XMLHttpRequest.responseText
Set XMLHttpRequest = Nothing

MsgBox "Token telah dihantar kepada telegram anda.", vbInformation, "Info"

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " telegram : send_developer_token" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub send_access_telegram()
'On Error GoTo logging:
Dim A1 As String
Dim File_Path As String
Dim LM_LIMIT As Integer
Dim LM_ERR_DETAIL As String

App.LogEvent "System Access", vbLogEventTypeInformation

A1 = A1 & "**----------------------**" & "%0A" & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & "Token : " & G_TOKEN & "%0A" & vbCrLf
A1 = A1 & "Terminal : " & G_TERMINAL & "%0A" & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & "Detail : **" & G_ERROR_NAIYO & "**" & "%0A" & vbCrLf

strURL = G_TOKEN_LICENSE & "/" & "sendmessage?chat_id=" & G_CHAT_DEV & "&text=" & A1 & ""

Set XMLHttpRequest = New MSXML2.XMLHTTP
XMLHttpRequest.Open "GET", strURL, False
'XMLHttpRequest.setRequestHeader "Content-Type", "text/xml"
XMLHttpRequest.Send

strResponse = XMLHttpRequest.responseText
Set XMLHttpRequest = Nothing

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " telegram : send_access_telegram" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub generate_token_pass()
'On Error GoTo logging:
all_chars = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")

Randomize
For i = 1 To 10
    random_index = Int(Rnd() * 61)
    clave = clave & all_chars(random_index)
Next

G_TOKEN_PASS = clave

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " telegram : generate_token_pass" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub send_successful_notification()
'On Error GoTo logging:
Dim A1 As String
Dim File_Path As String
Dim LM_LIMIT As Integer
Dim LM_ERR_DETAIL As String

App.LogEvent "Send Log Report Update Version", vbLogEventTypeInformation

A1 = A1 & "** Version Update **" & "%0A" & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & "Sistem : " & G_VERSION_CONTROL & "%0A" & vbCrLf
A1 = A1 & "Station : " & G_TERMINAL & "%0A" & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & "Client : " & G_CLIENT_INFO(0) & "%0A" & vbCrLf
A1 = A1 & "Database : " & G_CLIENT_INFO(2) & "%0A" & vbCrLf
A1 = A1 & "Version : " & G_VERSION_UPDATE & "%0A" & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & "!! Sistem Telah Berjaya Dikemaskini !!" & "%0A" & vbCrLf

strURL = G_TOKEN_SUCCESS & "/" & "sendmessage?chat_id=" & G_CHAT_SUCCESS & "&text=" & A1 & ""
    
Set XMLHttpRequest = New MSXML2.XMLHTTP
XMLHttpRequest.Open "GET", strURL, False
'XMLHttpRequest.setRequestHeader "Content-Type", "text/xml"
XMLHttpRequest.Send

strResponse = XMLHttpRequest.responseText
Set XMLHttpRequest = Nothing

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " telegram : send_successful_notification" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub


