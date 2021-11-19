Attribute VB_Name = "mod_api"
Public G_API_URL(5)
Public G_API_URL2(5)
Public G_API_URL_2
Public G_API_URL_3
Public G_DATA
Public G_API_RESULT
Public G_TOKEN
Public G_PING_API As Integer
Sub ping_api()
On Error GoTo logging:
Dim data As String
LM_STATUS = 0
G_DATA = "{""samaran"":""" & G_TOKEN & """}"

re_try:
'url / api / function
sRequest = G_API_URL(G_PING_API) '& G_API_URL_2 & G_API_URL_3
data = G_DATA
    
'Assumed MSXML2 core service well installed
Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
Set oxmlhttp2 = New MSXML2.ServerXMLHTTP
oxmlhttp2.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
oxmlhttp2.Open "post", sRequest, False
oxmlhttp2.setRequestHeader "Content-Type", "application/json"
oxmlhttp2.Send data
Set oXMLDoc2 = New MSXML2.DOMDocument
oXMLDoc2.async = False
If oxmlhttp2.readyState = 4 Or oxmlhttp2.Status = 200 Then
  G_API_RESULT = oXMLDoc2.LoadXML(oxmlhttp2.responseText)
Else
  G_API_RESULT = False
End If

Dim p As Object

Dim sInputJson As String
G_API_RESULT = oxmlhttp2.responseText

sInputJson = G_API_RESULT

If sInputJson = "null" Then
    MsgBox "Invalid Access", vbCritical, "Invalid Access"
    End
End If

'MsgBox "Input JSON string: " & sInputJson

'MsgBox "Input JSON string: " & sInputJson

Set p = JSON.parse(sInputJson)

'MsgBox JSON.toString(p)

LM_CREDENTIAL = JSON.toString(p)

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(2)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

LM_STATUS = LM_CREDENTIAL_3

'Telegram
LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(3)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_CHAT_DEV = LM_CREDENTIAL_3

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(4)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_TOKEN_LICENSE = "https://api.telegram.org/" & LM_CREDENTIAL_3 & ":"

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(5)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_TOKEN_LICENSE = G_TOKEN_LICENSE & LM_CREDENTIAL_3

If LM_STATUS = 0 Then 'Access Not Allowed
    G_ERROR_NAIYO = Now & " Access Not Granted"
    Call send_access_telegram
    
    MsgBox "Invalid Access.", vbCritical, "Info"
    End
ElseIf LM_STATUS = 2 Then 'Access Under Monitoring
    G_ERROR_NAIYO = Now & " Access Under Monitoring"
    Call send_access_telegram
End If

Call ping_api_2nd_layer

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_api : ping_api" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
'Call log_rekod

G_PING_API = G_PING_API + 1
frm303.L1_Text.Caption = "Authentication 2nd layer...."
Call ping_api2
End Sub
Sub ping_api2()
On Error GoTo logging:
Dim data As String
LM_STATUS = 0
G_DATA = "{""samaran"":""" & G_TOKEN & """}"

re_try:
'url / api / function
sRequest = G_API_URL(G_PING_API) '& G_API_URL_2 & G_API_URL_3
data = G_DATA
    
'Assumed MSXML2 core service well installed
Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
Set oxmlhttp2 = New MSXML2.ServerXMLHTTP
oxmlhttp2.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
oxmlhttp2.Open "post", sRequest, False
oxmlhttp2.setRequestHeader "Content-Type", "application/json"
oxmlhttp2.Send data
Set oXMLDoc2 = New MSXML2.DOMDocument
oXMLDoc2.async = False
If oxmlhttp2.readyState = 4 Or oxmlhttp2.Status = 200 Then
  G_API_RESULT = oXMLDoc2.LoadXML(oxmlhttp2.responseText)
Else
  G_API_RESULT = False
End If


Dim p As Object

Dim sInputJson As String
G_API_RESULT = oxmlhttp2.responseText

sInputJson = G_API_RESULT

If sInputJson = "null" Then
    MsgBox "Invalid Access", vbCritical, "Invalid Access"
    End
End If

'MsgBox "Input JSON string: " & sInputJson

'MsgBox "Input JSON string: " & sInputJson

Set p = JSON.parse(sInputJson)

'MsgBox JSON.toString(p)

LM_CREDENTIAL = JSON.toString(p)

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(2)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

LM_STATUS = LM_CREDENTIAL_3

'Telegram
LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(3)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_CHAT_DEV = LM_CREDENTIAL_3

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(4)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_TOKEN_LICENSE = "https://api.telegram.org/" & LM_CREDENTIAL_3 & ":"

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(5)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_TOKEN_LICENSE = G_TOKEN_LICENSE & LM_CREDENTIAL_3

If LM_STATUS = 0 Then 'Access Not Allowed
    G_ERROR_NAIYO = Now & " Access Not Granted"
    Call send_access_telegram
    
    MsgBox "Invalid Access.", vbCritical, "Info"
    End
ElseIf LM_STATUS = 2 Then 'Access Under Monitoring
    G_ERROR_NAIYO = Now & " Access Under Monitoring"
    Call send_access_telegram
End If

Call ping_api_2nd_layer

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_api : ping_api2" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
'Call log_rekod

G_PING_API = G_PING_API + 1
frm303.L1_Text.Caption = "Authentication 3rd layer...."
Call ping_api3
End Sub
Sub ping_api3()
On Error GoTo logging:
Dim data As String
LM_STATUS = 0
G_DATA = "{""samaran"":""" & G_TOKEN & """}"

re_try:
'url / api / function
sRequest = G_API_URL(G_PING_API) '& G_API_URL_2 & G_API_URL_3
data = G_DATA
    
'Assumed MSXML2 core service well installed
Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
Set oxmlhttp2 = New MSXML2.ServerXMLHTTP
oxmlhttp2.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
oxmlhttp2.Open "post", sRequest, False
oxmlhttp2.setRequestHeader "Content-Type", "application/json"
oxmlhttp2.Send data
Set oXMLDoc2 = New MSXML2.DOMDocument
oXMLDoc2.async = False
If oxmlhttp2.readyState = 4 Or oxmlhttp2.Status = 200 Then
  G_API_RESULT = oXMLDoc2.LoadXML(oxmlhttp2.responseText)
Else
  G_API_RESULT = False
End If


Dim p As Object

Dim sInputJson As String
G_API_RESULT = oxmlhttp2.responseText

sInputJson = G_API_RESULT

If sInputJson = "null" Then
    MsgBox "Invalid Access", vbCritical, "Invalid Access"
    End
End If

'MsgBox "Input JSON string: " & sInputJson

'MsgBox "Input JSON string: " & sInputJson

Set p = JSON.parse(sInputJson)

'MsgBox JSON.toString(p)

LM_CREDENTIAL = JSON.toString(p)

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(2)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

LM_STATUS = LM_CREDENTIAL_3

'Telegram
LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(3)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_CHAT_DEV = LM_CREDENTIAL_3

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(4)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_TOKEN_LICENSE = "https://api.telegram.org/" & LM_CREDENTIAL_3 & ":"

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(5)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_TOKEN_LICENSE = G_TOKEN_LICENSE & LM_CREDENTIAL_3

If LM_STATUS = 0 Then 'Access Not Allowed
    G_ERROR_NAIYO = Now & " Access Not Granted"
    Call send_access_telegram
    
    MsgBox "Invalid Access.", vbCritical, "Info"
    End
ElseIf LM_STATUS = 2 Then 'Access Under Monitoring
    G_ERROR_NAIYO = Now & " Access Under Monitoring"
    Call send_access_telegram
End If

Call ping_api_2nd_layer

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_api : ping_api3" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
'Call log_rekod

G_PING_API = G_PING_API + 1
frm303.L1_Text.Caption = "Authentication 4th layer...."
Call ping_api4
End Sub
Sub ping_api4()
On Error GoTo logging:
Dim data As String
LM_STATUS = 0
G_DATA = "{""samaran"":""" & G_TOKEN & """}"

re_try:
'url / api / function
sRequest = G_API_URL(G_PING_API) '& G_API_URL_2 & G_API_URL_3
data = G_DATA
    
'Assumed MSXML2 core service well installed
Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
Set oxmlhttp2 = New MSXML2.ServerXMLHTTP
oxmlhttp2.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
oxmlhttp2.Open "post", sRequest, False
oxmlhttp2.setRequestHeader "Content-Type", "application/json"
oxmlhttp2.Send data
Set oXMLDoc2 = New MSXML2.DOMDocument
oXMLDoc2.async = False
If oxmlhttp2.readyState = 4 Or oxmlhttp2.Status = 200 Then
  G_API_RESULT = oXMLDoc2.LoadXML(oxmlhttp2.responseText)
Else
  G_API_RESULT = False
End If


Dim p As Object

Dim sInputJson As String
G_API_RESULT = oxmlhttp2.responseText

sInputJson = G_API_RESULT

If sInputJson = "null" Then
    MsgBox "Invalid Access", vbCritical, "Invalid Access"
    End
End If

'MsgBox "Input JSON string: " & sInputJson

'MsgBox "Input JSON string: " & sInputJson

Set p = JSON.parse(sInputJson)

'MsgBox JSON.toString(p)

LM_CREDENTIAL = JSON.toString(p)

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(2)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

LM_STATUS = LM_CREDENTIAL_3

'Telegram
LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(3)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_CHAT_DEV = LM_CREDENTIAL_3

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(4)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_TOKEN_LICENSE = "https://api.telegram.org/" & LM_CREDENTIAL_3 & ":"

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(5)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_TOKEN_LICENSE = G_TOKEN_LICENSE & LM_CREDENTIAL_3

If LM_STATUS = 0 Then 'Access Not Allowed
    G_ERROR_NAIYO = Now & " Access Not Granted"
    Call send_access_telegram
    
    MsgBox "Invalid Access.", vbCritical, "Info"
    End
ElseIf LM_STATUS = 2 Then 'Access Under Monitoring
    G_ERROR_NAIYO = Now & " Access Under Monitoring"
    Call send_access_telegram
End If

Call ping_api_2nd_layer

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_api : ping_api4" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
'Call log_rekod

Resume Next
End Sub
Sub ping_api_2nd_layer()
'On Error GoTo logging:
Dim data As String
LM_STATUS = 0
G_DATA = "{""samaran"":""" & G_TOKEN & """}"

'url / api / function
sRequest = G_API_URL2(G_PING_API)
data = G_DATA
    
'Assumed MSXML2 core service well installed
Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
Set oxmlhttp2 = New MSXML2.ServerXMLHTTP
oxmlhttp2.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
oxmlhttp2.Open "post", sRequest, False
oxmlhttp2.setRequestHeader "Content-Type", "application/json"
oxmlhttp2.Send data
Set oXMLDoc2 = New MSXML2.DOMDocument
oXMLDoc2.async = False
If oxmlhttp2.readyState = 4 Or oxmlhttp2.Status = 200 Then
  G_API_RESULT = oXMLDoc2.LoadXML(oxmlhttp2.responseText)
Else
  G_API_RESULT = False
End If
    

Dim p As Object

Dim sInputJson As String
G_API_RESULT = oxmlhttp2.responseText

sInputJson = G_API_RESULT

If sInputJson = "null" Then
    MsgBox "Invalid Access", vbCritical, "Invalid Access"
    End
End If

'MsgBox "Input JSON string: " & sInputJson

'MsgBox "Input JSON string: " & sInputJson

Set p = JSON.parse(sInputJson)

'MsgBox JSON.toString(p)

LM_CREDENTIAL = JSON.toString(p)

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(2)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_SERVER_IP = LM_CREDENTIAL_3

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(3)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_SERVER_USER = LM_CREDENTIAL_3

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(5)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_SERVER_PORT = LM_CREDENTIAL_3

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(4)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

LM_STATUS = LM_CREDENTIAL_3

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(6)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_SERVER_PASS = LM_CREDENTIAL_3

LM_CREDENTIAL_1 = Split(LM_CREDENTIAL, ",")(7)
LM_CREDENTIAL_2 = Split(LM_CREDENTIAL_1, ":")(1)
LM_CREDENTIAL_3 = Split(LM_CREDENTIAL_2, """")(1)

G_SERVER_DATABASE = LM_CREDENTIAL_3

If LM_STATUS = 0 Then 'Access Not Allowed
    MsgBox "Invalid Access.", vbCritical, "Info"
    End
End If

Call Main

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_api : ping_api_2nd_layer" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub grab_api_data()
'On Error GoTo logging:
Dim data As String

Dim File_Path As String
File_Path = App.Path & "\mu_configuration.prn.txt"
Open File_Path For Input As #1

Line Input #1, LM_DUMMY
Line Input #1, LM_DUMMY
Line Input #1, LM_DUMMY
Line Input #1, LM_DUMMY
Line Input #1, LM_DUMMY
Line Input #1, LM_DUMMY
Line Input #1, LM_DUMMY
Line Input #1, LM_DUMMY
Line Input #1, LM_DUMMY
Line Input #1, LM_DUMMY
Line Input #1, LM_API_1 'tk
Line Input #1, LM_API_2
Line Input #1, LM_API_3
Line Input #1, LM_API_4
Line Input #1, LM_API_5
Line Input #1, LM_API_6 'Terminal
Line Input #1, LM_API_7 'sankyuauth
Line Input #1, LM_API_8
Line Input #1, LM_API_9
Line Input #1, LM_API_10 'Cawangan
Line Input #1, x
Line Input #1, x
Line Input #1, x
Line Input #1, x
Line Input #1, x
Line Input #1, x
Line Input #1, x
Line Input #1, x
Line Input #1, x
Line Input #1, x
Line Input #1, LM_API_11 '---.---.---.yyy
Line Input #1, LM_API_12 'yyy.yyy.yyy.---
Line Input #1, LM_API_13 'tk
Line Input #1, LM_API_14 'sankyuauth
Line Input #1, LM_API_15 '---.---.---.yyy
Line Input #1, LM_API_16 'yyy.yyy.yyy.---

Close #1

G_PING_API = 0

LM_URL = Split(LM_API_2, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0)
    
G_SYSTEM_TYPE = LM_URL_2

LM_URL = Split(LM_API_6, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0)

G_TERMINAL = LM_URL_2

LM_URL = Split(LM_API_3, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0)

G_TOKEN = LM_URL_2

Erase G_API_URL
Erase G_API_URL2

Call check_connectivity_conn

LM_URL = Split(LM_API_4, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0)

G_API_URL(0) = LM_URL_2
G_API_URL(1) = LM_URL_2
G_API_URL(2) = LM_URL_2
G_API_URL(3) = LM_URL_2

'ip & domain depan - start
LM_URL = Split(LM_API_7, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0) 'domain depan 1st layer

G_API_URL(0) = G_API_URL(0) + LM_URL_2

LM_URL = Split(LM_API_12, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0) 'ip depan 2nd layer

G_API_URL(1) = G_API_URL(1) + LM_URL_2

LM_URL = Split(LM_API_14, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0) 'domain depan 3rd layer

G_API_URL(2) = G_API_URL(2) + LM_URL_2

LM_URL = Split(LM_API_16, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0) 'ip depan 4th layer

G_API_URL(3) = G_API_URL(3) + LM_URL_2
'ip & domain depan - end

LM_URL = Split(LM_API_1, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0) 'domain belakang 1st layer

G_API_URL(0) = G_API_URL(0) + "." + LM_URL_2 + "/"

LM_URL = Split(LM_API_11, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0) 'ip belakang 2nd layer

G_API_URL(1) = G_API_URL(1) + "." + LM_URL_2 + "/"

LM_URL = Split(LM_API_13, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0) 'domain belakang 3rd layer

G_API_URL(2) = G_API_URL(2) + "." + LM_URL_2 + "/"

LM_URL = Split(LM_API_15, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0) 'ip belakang 4th layer

G_API_URL(3) = G_API_URL(3) + "." + LM_URL_2 + "/"

LM_URL = Split(LM_API_9, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0)

G_API_URL(0) = G_API_URL(0) + LM_URL_2 + "/"
G_API_URL(1) = G_API_URL(1) + LM_URL_2 + "/"
G_API_URL(2) = G_API_URL(2) + LM_URL_2 + "/"
G_API_URL(3) = G_API_URL(3) + LM_URL_2 + "/"

LM_URL = Split(LM_API_8, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0)

G_API_URL(0) = G_API_URL(0) + LM_URL_2 + "/"
G_API_URL(1) = G_API_URL(1) + LM_URL_2 + "/"
G_API_URL(2) = G_API_URL(2) + LM_URL_2 + "/"
G_API_URL(3) = G_API_URL(3) + LM_URL_2 + "/"

LM_URL = Split(LM_API_5, "''")(1)
LM_URL_1 = Split(LM_URL, "[")(1)
LM_URL_2 = Split(LM_URL_1, "]")(0)

G_API_URL2(0) = G_API_URL(0) + "access_detail"
G_API_URL2(1) = G_API_URL(1) + "access_detail"
G_API_URL2(2) = G_API_URL(2) + "access_detail"
G_API_URL2(3) = G_API_URL(3) + "access_detail"

G_API_URL(0) = G_API_URL(0) + LM_URL_2
G_API_URL(1) = G_API_URL(1) + LM_URL_2
G_API_URL(2) = G_API_URL(2) + LM_URL_2
G_API_URL(3) = G_API_URL(3) + LM_URL_2

'MsgBox G_API_URL(0) & vbCrLf & _
        G_API_URL(1) & vbCrLf & _
        G_API_URL(2) & vbCrLf & _
        G_API_URL(3)

'MsgBox G_API_URL2(0) & vbCrLf & _
        G_API_URL2(1) & vbCrLf & _
        G_API_URL2(2) & vbCrLf & _
        G_API_URL2(3)
frm303.L1_Text.Caption = "Authentication 1st layer...."
Call ping_api

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_api : grab_api_data" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description

Call log_rekod
Resume Next
End Sub
