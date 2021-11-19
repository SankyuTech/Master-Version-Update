VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm1 
   Caption         =   "Client"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17670
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   17670
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   6600
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton CMD2 
         Caption         =   "Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3720
         MouseIcon       =   "frm1.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frm1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Batal"
         Top             =   4800
         Width           =   1425
      End
      Begin VB.CommandButton CMD4 
         Caption         =   "Simpan Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   2160
         MouseIcon       =   "frm1.frx":13D4
         MousePointer    =   99  'Custom
         Picture         =   "frm1.frx":16DE
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Simpan Data"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.CommandButton CMD32 
         Caption         =   "Token"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5160
         MouseIcon       =   "frm1.frx":2088
         MousePointer    =   99  'Custom
         Picture         =   "frm1.frx":2392
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Refresh"
         Top             =   3520
         Width           =   1425
      End
      Begin VB.TextBox TB2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2040
         MaxLength       =   255
         TabIndex        =   30
         Text            =   "TB2"
         Top             =   3720
         Width           =   3015
      End
      Begin VB.ComboBox CBB1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2880
         Width           =   4575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Token * :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   5
         Left            =   -120
         TabIndex        =   31
         Top             =   3735
         Width           =   2055
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version * :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   9
         Left            =   -120
         TabIndex        =   29
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label L2_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L2_Text"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   3
         Left            =   2040
         TabIndex        =   27
         Top             =   1920
         Width           =   4215
      End
      Begin VB.Label L1_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   3
         Left            =   -120
         TabIndex        =   26
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label L2_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L2_Text"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   2
         Left            =   2040
         TabIndex        =   25
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label L1_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Database :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   2
         Left            =   -120
         TabIndex        =   24
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label L2_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L2_Text"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   1
         Left            =   2040
         TabIndex        =   23
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label L1_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hostname :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   1
         Left            =   -120
         TabIndex        =   22
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label L2_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L2_Text"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   0
         Left            =   2040
         TabIndex        =   21
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label L1_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Client :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   0
         Left            =   -120
         TabIndex        =   20
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9560
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   15735
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Senarai Client"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Index           =   1
         Left            =   960
         TabIndex        =   9
         Top             =   2160
         Width           =   15375
         Begin VB.CommandButton CMD21 
            Caption         =   "Next"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   1
            Left            =   14040
            MouseIcon       =   "frm1.frx":495C
            MousePointer    =   99  'Custom
            Picture         =   "frm1.frx":4C66
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Paparan Seterusnya"
            Top             =   6480
            Width           =   1095
         End
         Begin VB.CommandButton CMD21 
            Caption         =   "Back"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   0
            Left            =   12840
            MouseIcon       =   "frm1.frx":5D30
            MousePointer    =   99  'Custom
            Picture         =   "frm1.frx":603A
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Paparan Sebelum"
            Top             =   6480
            Width           =   1095
         End
         Begin MSComctlLib.ListView LV3 
            Height          =   6100
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   15075
            _ExtentX        =   26591
            _ExtentY        =   10769
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :          / "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   10440
            TabIndex        =   18
            Top             =   6480
            Width           =   2295
         End
         Begin VB.Label L104_Text 
            BackColor       =   &H8000000C&
            Caption         =   "L104_Text"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   10560
            TabIndex        =   17
            Top             =   6840
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label L104_Text 
            BackColor       =   &H8000000C&
            Caption         =   "L104_Text"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   11520
            TabIndex        =   16
            Top             =   6840
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label L104_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L104_Text"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   0
            Left            =   11760
            TabIndex        =   15
            Top             =   6480
            Width           =   375
         End
         Begin VB.Label L104_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L104_Text"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   12360
            TabIndex        =   14
            Top             =   6480
            Width           =   615
         End
         Begin VB.Label L104_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L104_Text"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   4
            Left            =   240
            TabIndex        =   13
            Top             =   6480
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Carian"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   15375
         Begin VB.TextBox TB1 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   1500
            MaxLength       =   255
            TabIndex        =   5
            Text            =   "TB1"
            Top             =   360
            Width           =   4140
         End
         Begin VB.CommandButton CMD7 
            BackColor       =   &H80000004&
            Caption         =   "Carian"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   2760
            MaskColor       =   &H00400000&
            Picture         =   "frm1.frx":7104
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Carian Report"
            Top             =   2280
            Width           =   1425
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "> Keyword boleh mengandungi HOSTNAME , Token atau Nama Kedai."
            Height          =   615
            Index           =   1
            Left            =   1560
            TabIndex        =   8
            Top             =   1560
            Width           =   4335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "> Sila tinggalkan ruangan ini kosong bagi mencari semua data."
            Height          =   615
            Index           =   0
            Left            =   1560
            TabIndex        =   7
            Top             =   840
            Width           =   4335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Keyword :"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   -480
            TabIndex        =   6
            Top             =   375
            Width           =   1905
         End
      End
      Begin MSComctlLib.ListView LV2 
         Height          =   1575
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   15350
         _ExtentX        =   27067
         _ExtentY        =   2778
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Scan Item"
            Object.Width           =   2540
            ImageIndex      =   1
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Maklumat Pembeli"
            Object.Width           =   2540
            ImageIndex      =   2
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Bayaran"
            Object.Width           =   2540
            ImageIndex      =   3
         EndProperty
      End
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   16748
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Scan Item"
         Object.Width           =   2540
         ImageIndex      =   1
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Maklumat Pembeli"
         Object.Width           =   2540
         ImageIndex      =   2
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bayaran"
         Object.Width           =   2540
         ImageIndex      =   3
      EndProperty
   End
   Begin VB.Menu frm1_pm_menu 
      Caption         =   "Main Menu"
      Visible         =   0   'False
      Begin VB.Menu frm1_sm_version_sedang_digunakan 
         Caption         =   "Version Sedang Digunakan"
      End
      Begin VB.Menu frm1_sm_update_system 
         Caption         =   "Update Sistem Client Ini"
      End
      Begin VB.Menu frm1_sm_spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu frm1_sm_update_semua_client 
         Caption         =   "Update Sistem SEMUA Client"
      End
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CMD2_Click()
'On Error GoTo logging:
Note = "Batal ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    G_TOKEN_PASS = vbNullString
    
    frm1.Frame3.Visible = False
    frm1.Frame3.ZOrder 1
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm1 : CMD2_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub

Private Sub CMD21_Click(Index As Integer)
'On Error GoTo logging:
Dim frm1_LM_CURR_PAGE As Double
Dim frm1_LM_TOTAL_PAGE As Double

frm1_LM_CURR_PAGE = 0
frm1_LM_TOTAL_PAGE = 0

If frm1.L104_Text(0) <> vbNullString And IsNumeric(frm1.L104_Text(0)) Then
    If frm1.L104_Text(1) <> vbNullString And IsNumeric(frm1.L104_Text(1)) Then
        frm1_LM_CURR_PAGE = frm1.L104_Text(0)
        frm1_LM_TOTAL_PAGE = frm1.L104_Text(1)
        
        If Index = 0 Then
            If frm1_LM_CURR_PAGE <> 1 And frm1_LM_CURR_PAGE <> 0 Then
            
                GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                Call frm1_senarai_client_header
                Call frm1_senarai_client
            End If
        ElseIf Index = 1 Then
            If frm1_LM_CURR_PAGE < frm1_LM_TOTAL_PAGE Then
            
                GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
                Call frm1_senarai_client_header
                Call frm1_senarai_client
            End If
        End If
    End If
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm1 : CMD21_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub CMD32_Click()
'On Error GoTo logging:
frm1.TB2 = vbNullString
G_TOKEN_PASS = vbNullString

Call generate_token_pass
Call send_developer_token

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm1 : CMD32_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub

Private Sub CMD4_Click()
'On Error GoTo logging:
If G_TOKEN_PASS = vbNullString Then
    MsgBox "Sila request token untuk meneruskan urusan ini.", vbExclamation, "Info"
    Exit Sub
End If
If frm1.CBB1 = vbNullString Then
    MsgBox "Sila buat pilihan version.", vbExclamation, "Info"
    Exit Sub
End If
If G_CLIENT_INFO(2) = vbNullString Then
    MsgBox "Tiada maklumat database client.", vbExclamation, "Info"
    Exit Sub
End If
If G_TOKEN_PASS <> frm1.TB2 Then
    MsgBox "Token yang tidak sah.", vbExclamation, "Info"
    Exit Sub
End If

Call frm1_check_credential_remote

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm1 : CMD4_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub

Private Sub CMD7_Click()
'On Error GoTo logging:
Call frm1_cek_krateria_report

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm1 : CMD7_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub

Private Sub Form_Load()
'On Error GoTo logging:
frm1.Picture = MDI_frm1.Picture
frm1.LV1.ListItems.Clear

With frm1.LV1
    Set .SmallIcons = MDI_frm1.ImageList4
    Set .Icons = MDI_frm1.ImageList4

    .ListItems.Add , "Client", "Client", 24
    .ListItems.Add , "Log", "Log", 42
End With

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm1 : Form_Load" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub

Private Sub frm1_sm_update_semua_client_Click()
'On Error GoTo logging:
LM_FOUND = 0

frm1_LM_No_ID = vbNullString

If IsNumeric(frm1.LV3.SelectedItem.Index) Then
    frm1_LM_No_ID = frm1.LV3.ListItems(frm1.LV3.SelectedItem.Index)
    
    If frm1_LM_No_ID <> vbNullString Then

        G_TOKEN_PASS = vbNullString
        
        Call generate_token_pass
        Call send_developer_token

        Note = "Update sistem bagi SEMUA client?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Jika anda yakin , sila masukkan TOKEN bagi mulakan proses update." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
        
        LM_TOKEN = InputBox(Note, "Token", "")
        
        If StrPtr(LM_TOKEN) = 0 Then
            MsgBox "Tiada TOKEN dimasukkan.", vbInformation, "Info"
            Exit Sub
        End If
        
        If StrPtr(LM_TOKEN) <> 0 Then
            If G_TOKEN_PASS = LM_TOKEN Then
                Call frm1_start_update_all_system
            Else
                MsgBox "TOKEN yang dimasukkan tidak sah.", vbExclamation, "Info"
                Exit Sub
            End If
        End If
    Else
        MsgBox "Tiada data", vbInformation, "Info"
    End If
Else
    MsgBox "Tiada data", vbInformation, "Info"
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm1 : frm1_sm_update_semua_client_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub

Private Sub frm1_sm_update_system_Click()
'On Error GoTo logging:
Dim rs1 As ADODB.Recordset
LM_FOUND = 0

frm1_LM_No_ID = vbNullString

If IsNumeric(frm1.LV3.SelectedItem.Index) Then
    frm1_LM_No_ID = frm1.LV3.ListItems(frm1.LV3.SelectedItem.Index)
    
    If frm1_LM_No_ID <> vbNullString Then
        Note = "Update sistem bagi client ini?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
LM_CONN = 1
re_conn_1:
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 2_client_details where ID='" & frm1_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic

            If Not rs.EOF Then
                Call frm1_reset_l2
                Erase G_CLIENT_INFO
                If Not IsNull(rs!client) Then G_CLIENT_INFO(0) = rs!client
                If Not IsNull(rs!credential_1) Then G_CLIENT_INFO(1) = rs!credential_1
                If Not IsNull(rs!credential_5) Then G_CLIENT_INFO(2) = rs!credential_5
                If Not IsNull(rs!version_in_use) Then G_CLIENT_INFO(3) = rs!version_in_use
                If Not IsNull(rs!token) Then G_CLIENT_INFO(4) = rs!token
                LM_FOUND = 1
            Else
                MsgBox "Status data bagi data ini telah berubah. Sila periksa status terbaru data ini.", vbExclamation, "Info"
            End If
            
            rs.Close
            Set rs = Nothing
            
            If LM_FOUND = 1 Then
                frm1.L2_Text(0) = G_CLIENT_INFO(0)
                frm1.L2_Text(1) = G_CLIENT_INFO(1)
                frm1.L2_Text(2) = G_CLIENT_INFO(2)
                frm1.L2_Text(3) = G_CLIENT_INFO(3)
            
                Call frm1_initial_frame3
                Call frm1_senarai_version
                G_TOKEN_PASS = vbNullString
                frm1.TB2 = vbNullString
                
                frm1.Frame3.Visible = True
                frm1.Frame3.ZOrder 0
            Else
                MsgBox "Tiada data", vbInformation, "Info"
            End If
        End If
    Else
        MsgBox "Tiada data", vbInformation, "Info"
    End If
Else
    MsgBox "Tiada data", vbInformation, "Info"
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm1 : frm1_sm_update_system_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
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

Private Sub frm1_sm_version_sedang_digunakan_Click()
'On Error GoTo logging:
Dim rs1 As ADODB.Recordset
LM_FOUND = 0

frm1_LM_No_ID = vbNullString

If IsNumeric(frm1.LV3.SelectedItem.Index) Then
    frm1_LM_No_ID = frm1.LV3.ListItems(frm1.LV3.SelectedItem.Index)
    
    If frm1_LM_No_ID <> vbNullString Then
        Note = "Carian version yang sedanng digunakan?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
LM_CONN = 1
re_conn_1:
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 2_client_details where ID='" & frm1_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic

            If Not rs.EOF Then
                Erase G_CLIENT_INFO
                If Not IsNull(rs!client) Then G_CLIENT_INFO(0) = rs!client
                If Not IsNull(rs!credential_1) Then G_CLIENT_INFO(1) = rs!credential_1
                If Not IsNull(rs!credential_5) Then G_CLIENT_INFO(2) = rs!credential_5
                If Not IsNull(rs!version_in_use) Then G_CLIENT_INFO(3) = rs!version_in_use
                If Not IsNull(rs!token) Then G_CLIENT_INFO(4) = rs!token
                LM_FOUND = 1
            Else
                MsgBox "Status data bagi data ini telah berubah. Sila periksa status terbaru data ini.", vbExclamation, "Info"
            End If
            
            rs.Close
            Set rs = Nothing
            
            If LM_FOUND = 1 Then
                Call frm1_check_version_sedang_digunakan
            Else
                MsgBox "Tiada data", vbInformation, "Info"
            End If
        End If
    Else
        MsgBox "Tiada data", vbInformation, "Info"
    End If
Else
    MsgBox "Tiada data", vbInformation, "Info"
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " frm1 : frm1_sm_version_sedang_digunakan_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
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

Private Sub LV1_Click()
'On Error Resume Next
LM_KEY = frm1.LV1.SelectedItem.key

If LM_KEY = "Client" Then
    Call frm1_initial_frame1
    Call frm1_initial_frame2
    Call frm1_initial_frame3
    
    frm1.LV2.ListItems.Clear
    
    With frm1.LV2
        Set .SmallIcons = MDI_frm1.ImageList4
        Set .Icons = MDI_frm1.ImageList4
    
        .ListItems.Add , "Client", "Client", 67
        .ListItems.Add , "Senarai Client", "Senarai Client", 24
    End With

    frm1.Frame1(0).Visible = True
End If
End Sub
Private Sub LV2_Click()
'On Error Resume Next
LM_KEY = frm1.LV2.SelectedItem.key

If LM_KEY = "Client" Then
    Call frm1_initial_frame2
    Call frm1_initial_frame3
    frm1.TB1 = vbNullString
    frm1.Frame2(0).Visible = True
    frm1.TB1.SetFocus
ElseIf LM_KEY = "Senarai Client" Then
    Call frm1_initial_frame2
    Call frm1_initial_frame3
    frm1.Frame2(1).Visible = True
End If
End Sub
Private Sub LV3_DblClick()
'On Error Resume Next
frm1_LM_No_ID = vbNullString

If IsNumeric(frm1.LV3.SelectedItem.Index) Then
    frm1_LM_No_ID = frm1.LV3.ListItems(frm1.LV3.SelectedItem.Index)
    
    If frm1_LM_No_ID <> vbNullString Then
        PopupMenu frm1_pm_menu
    End If
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
