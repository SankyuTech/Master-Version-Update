VERSION 5.00
Begin VB.Form frm302 
   BackColor       =   &H80000004&
   Caption         =   "Login"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10800
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frm302.frx":0000
   ScaleHeight     =   8040
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TB1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "TB1"
      Top             =   6960
      Width           =   4620
   End
   Begin VB.TextBox TB1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Text            =   "TB1"
      Top             =   6360
      Width           =   4620
   End
   Begin VB.Label L1_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila klik di sini jika anda lupa username atau password."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   480
      MouseIcon       =   "frm302.frx":25CA
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   7560
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BATAL"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   1095
      Left            =   9000
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   1095
      Left            =   7320
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   1080
      Left            =   9000
      Picture         =   "frm302.frx":28D4
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   1
      Left            =   -480
      TabIndex        =   3
      Top             =   6975
      Width           =   2745
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   0
      Left            =   -480
      TabIndex        =   2
      Top             =   6375
      Width           =   2745
   End
   Begin VB.Image Image1 
      Height          =   6150
      Left            =   0
      Picture         =   "frm302.frx":C222
      Top             =   0
      Width           =   10905
   End
   Begin VB.Image Image2 
      Height          =   1080
      Left            =   7320
      Picture         =   "frm302.frx":1FF77
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1305
   End
End
Attribute VB_Name = "frm302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()
'On Error GoTo logging:
Call frm302_login

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm302 : Image2_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub Image3_Click()
'On Error GoTo logging:
Note = "Batal Login Ke Dalam Sistem?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Sistem Akan Ditutup Jika Anda Teruskan."

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Login")

If Answer = vbYes Then
    End
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm302 : Image3_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub

Private Sub L1_Text_Click()
'On Error GoTo logging:
Note = "Sila masukkan e-mail anda yang didaftarkan ke dalam sistem ini." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Sistem akan menghantar USERNAME dan PASSWORD anda ke email ini."

G_MAIL = InputBox(Note, "Username dan password", "Masukkan e-mail anda")

If StrPtr(G_MAIL) = 0 Then
    Exit Sub
End If

If StrPtr(G_MAIL) <> 0 Then
    myAt = InStr(1, G_MAIL, "@", vbTextCompare)
    myDot = InStr(myAt + 2, G_MAIL, ".", vbTextCompare)
    myDotDot = InStr(myAt + 2, G_MAIL, "..", vbTextCompare)
    
    If myAt = 0 Or myDot = 0 Or Not myDotDot = 0 Or Right(G_MAIL, 1) = "." Then
        MsgBox "E-mail yang tidak sah.", vbExclamation, "Info"
        
        Exit Sub
    End If

    Call Frm3_check_internet
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm302 : L1_Text_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub

Private Sub TB1_KeyPress(Index As Integer, KeyAscii As Integer)
'On Error GoTo logging:
If Index = 0 Then
    If KeyAscii = 13 Then
        frm302.TB1(1).SetFocus
    End If
ElseIf Index = 1 Then
    If KeyAscii = 13 Then
        Call frm302_login
    End If
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm302 : TB1_KeyPress" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
