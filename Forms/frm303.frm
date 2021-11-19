VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frm303 
   Caption         =   "Loading......."
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7440
   ControlBox      =   0   'False
   Icon            =   "frm303.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frm303.frx":25CA
   ScaleHeight     =   9510
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr2 
      Interval        =   200
      Left            =   360
      Top             =   240
   End
   Begin VB.Timer tmr1 
      Interval        =   200
      Left            =   960
      Top             =   240
   End
   Begin XPControls.ProgBarXP ProgBarXP1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   873
      Style           =   3
      BarColor        =   16711680
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COPYRIGHT @ Sankyu System 2013"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   9120
      Width           =   4935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Website : www.sankyutechnology.com"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   8760
      Width           =   6975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail : info@sankyutech.com"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   8400
      Width           =   6975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No. : +6010 - 900 4788"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   8040
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SANKYU SYSTEM"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by Sankyu Tech Sdn. Bhd."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7680
      Width           =   6975
   End
   Begin VB.Label L1_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait........"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   6240
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   6165
      Left            =   0
      Picture         =   "frm303.frx":4B94
      Top             =   0
      Width           =   7440
   End
End
Attribute VB_Name = "frm303"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error GoTo logging:
Dim rc As Long

If App.PrevInstance Then
    rc = MsgBox("Sistem Pengurusan Kedai Emas Telah Dibuka Sebelum Ini.", vbCritical, App.Title)
    
    End
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm303 : Form_Load" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub

Private Sub Tmr1_Timer()
'On Error Resume Next
Static counting As Integer
counting = counting + 1

If counting = 3 Then
    frm303.L1_Text.Caption = "Preparing ..."
ElseIf counting = 10 Then
    frm303.L1_Text.Caption = "Loading forms ..."
ElseIf counting = 20 Then
    frm303.L1_Text.Caption = "Checking Connection System And Databases ..."
    
    G_VERSION_CONTROL = "MU1.0.0"
    'Call check_internet_connection_main
    Call grab_api_data
    Call check_internet_connection_main
    If G_SYSTEM_TYPE = "ONLINE" Then
        If MDI_frm1.L17_Text = "ONLINE" Then
            Call Main
        Else
            MsgBox "Tiada sambungan internet. Sila pastikan komputer anda disambungkan dengan internet bagi membolehkan sistem beroperasi.", vbCritical, App.Title
            
            End
        End If
    End If
ElseIf counting = 25 Then
    frm303.L1_Text.Caption = "Loading Databases ..."
ElseIf counting = 35 Then
    'frm303.tmr1.Enabled = False
    frm303.L1_Text.Caption = "Done ..."
ElseIf counting = 45 Then
    frm303.L1_Text.Caption = "Starting Sistem Pengurusan Kedai Emas ..."
ElseIf counting = 60 Then

End If
End Sub
Private Sub Tmr2_Timer()
'On Error Resume Next
frm303.ProgBarXP1.Max = 600
frm303.ProgBarXP1.Value = Int(frm303.ProgBarXP1.Value) + 10
If frm303.ProgBarXP1.Value >= frm303.ProgBarXP1.Max Then
    Unload Me
    
    Call frm302_reset_all
    frm302.Show
    frm302.TB1(0).SetFocus
End If
End Sub
