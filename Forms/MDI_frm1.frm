VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.MDIForm MDI_frm1 
   BackColor       =   &H80000016&
   ClientHeight    =   10755
   ClientLeft      =   225
   ClientTop       =   270
   ClientWidth     =   20010
   Icon            =   "MDI_frm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   480
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Tmr4 
      Interval        =   60000
      Left            =   5280
      Top             =   1320
   End
   Begin VB.Timer Tmr3 
      Interval        =   50
      Left            =   4680
      Top             =   1320
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   20010
      _ExtentX        =   35295
      _ExtentY        =   1905
      ButtonWidth     =   1931
      ButtonHeight    =   1852
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SETTING"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "LOGOUT"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SHUTDOWN"
            ImageIndex      =   24
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   5400
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   9240
         ScaleHeight     =   1095
         ScaleWidth      =   16215
         TabIndex        =   13
         Top             =   0
         Width           =   16215
         Begin VB.Label L17_Text 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ONLINE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   9000
            TabIndex        =   14
            Top             =   720
            Width           =   1815
         End
         Begin VB.Image Image3 
            Height          =   720
            Left            =   9600
            Picture         =   "MDI_frm1.frx":25CA
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   9720
            Picture         =   "MDI_frm1.frx":4B94
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label L9_Text 
            Alignment       =   2  'Center
            Caption         =   "Master System"
            BeginProperty Font 
               Name            =   "Script MT Bold"
               Size            =   39.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1215
            Left            =   -3600
            TabIndex        =   15
            Top             =   0
            Width           =   12225
         End
      End
      Begin VB.Image Image4 
         Height          =   720
         Left            =   10120
         Picture         =   "MDI_frm1.frx":153DE
         Top             =   240
         Width           =   720
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   1320
         Picture         =   "MDI_frm1.frx":179A8
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   840
      End
   End
   Begin VB.Timer Tmr2 
      Interval        =   50
      Left            =   4200
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   3720
      Top             =   1320
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   3240
      Top             =   1320
   End
   Begin VB.PictureBox Pic1 
      Align           =   4  'Align Right
      Height          =   9675
      Left            =   17595
      ScaleHeight     =   9615
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
      Begin VB.Label L23_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L23_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   5400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label L22_Text 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label L20_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L20_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   16
         Top             =   990
         Width           =   1365
      End
      Begin VB.Label L19_Text 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label L18_Text 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label L16_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L16_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   8040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label L12_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Tukar user"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   0
         MouseIcon       =   "MDI_frm1.frx":19F72
         MousePointer    =   99  'Custom
         TabIndex        =   8
         ToolTipText     =   "Tukar password"
         Top             =   7680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label L5_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L5_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   8040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         Height          =   1035
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   540
         Width           =   2300
      End
      Begin VB.Label L4_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L4_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   810
         Width           =   2025
      End
      Begin VB.Label L2_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L2_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label L1_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L1_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   2775
      End
      Begin VB.Label L3_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L3_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   610
         Width           =   2025
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   ":    :     :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User        Level  Branch"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2160
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":1A27C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":1C856
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":1EE30
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":2140A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":239E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":25FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":28598
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":3D932
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":52CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":68066
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":7D400
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":9279A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":A7B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":C8B66
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":DDF00
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":F329A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":11D9AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":11FF86
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":122560
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":124B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":127114
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":1296EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":12BCC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":12E2A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   96
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":13087C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":132E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":135430
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":137A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":139FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":13C5BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":13EB98
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":141172
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":1453CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":155C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":159E80
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":16A6DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":16CCB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":17D50E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":18DD68
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":19E5C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":1AEE1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":1B3076
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":1C38D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":1D412A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":1D6704
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":1D8CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":1E9538
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":1F9D92
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":20A5EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":20CBC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":221F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":2372FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":242FB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":251226
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":2665C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":274832
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":289BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":29EF66
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":2B4300
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":2C969A
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":2DC47C
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":2E5B7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":2F63A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":30B742
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":320ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":335E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":34D7C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":36E7FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":383B94
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":398F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":3AE2C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":3C3662
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":3D89FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":40EDA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":445150
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":45597A
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":48008C
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":495426
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":4A8208
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":4BD5A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":4D293C
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":4E7CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":4EA2B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":4EC88A
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":4EEE64
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":4F143E
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5067D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":508DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":50B38C
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":50D966
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":50FF40
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":51251A
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":514AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5170CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5196A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":51BC82
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":51E25C
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":520836
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":522E10
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5381AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":54641C
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5489F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":54AFD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":54D5AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":54FB84
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":55215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":554738
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":556D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5592EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":55B8C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":55DEA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":56047A
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":562A54
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":56502E
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":567608
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":569BE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":56C1BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":56E796
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":570D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":57334A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   12720
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   70
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5736D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":575CAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":578289
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":57A863
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":57CE3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":57F417
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5819F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":583FCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5865A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":588B7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":58B159
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":58D733
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":58FD0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5922E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5948C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":596E9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":599475
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":59BA4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":59E029
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5A0603
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5A2BDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5A51B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5A7791
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5A9D6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5AC345
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5AE91F
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5B0EF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5B34D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5B5AAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5B8087
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5BA661
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5BCC3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5BF215
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5C17EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5C3DC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5C63A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5C897D
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5CAF57
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5CD531
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5CFB0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5D20E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5D46BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5D6C99
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5D9273
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5DB84D
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5DDE27
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5E0401
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5E29DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5E4FB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5E758F
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5E9B69
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5EC143
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5EE71D
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5F0CF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5F32D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5F58AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5F7E85
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5FA45F
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5FCA39
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":5FF013
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6015ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":603BC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6061A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":60877B
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":60AD55
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":60D32F
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":60F909
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":611EE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6144BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":616A97
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   17160
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   73
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":619071
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":61B64B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":61DC25
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6201FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6227D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":624DB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":62738D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":629967
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":62BF41
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":62E51B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":630AF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6330CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6356A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":637C83
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":63A25D
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":63C837
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":63EE11
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6413EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6439C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":645F9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":648579
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":64AB53
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":64D12D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":64F707
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":651CE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6542BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":656895
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":658E6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":65B449
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":65DA23
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":65FFFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6625D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":664BB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":66718B
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":669765
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":66BD3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":66E319
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6708F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":672ECD
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6754A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":677A81
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":67A05B
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":67C635
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":67EC0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6811E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6837C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":685D9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":688377
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":68A951
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":68CF2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":68F505
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":691ADF
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6B9EF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6BC4CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6BEAA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6C1081
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6C365B
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6C5C35
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6C820F
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6CA7E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6CCDC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6CF39D
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6D1977
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6D3F51
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6D652B
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6D8B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6DB0DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6DD6B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6DFC93
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6E226D
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6E4847
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6E6E21
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":6E93FB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   4320
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MDI_PM_menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MDI_SM_tukar_background 
         Caption         =   "&Tukar warna background sistem"
         Begin VB.Menu MD1_SM_SM_1 
            Caption         =   "1"
         End
         Begin VB.Menu MD1_SM_SM_2 
            Caption         =   "2"
         End
         Begin VB.Menu MD1_SM_SM_3 
            Caption         =   "3"
         End
         Begin VB.Menu MD1_SM_SM_4 
            Caption         =   "4"
         End
         Begin VB.Menu MD1_SM_SM_5 
            Caption         =   "5"
         End
         Begin VB.Menu MD1_SM_SM_6 
            Caption         =   "6"
         End
         Begin VB.Menu MD1_SM_SM_7 
            Caption         =   "7"
         End
      End
   End
End
Attribute VB_Name = "MDI_frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xLogOff As Boolean



Private Sub Command1_Click()
text1 = "[]0,[]0,[]0,[]0,[]1,[]1,[]0,[]0,[]10,[]0,[]1,[]0,[]3,[]0,[]10,[]https://www.sankyutechnology.com,[]0,[]0,[]0,[]1,[]1,[]0,[]0,[]0,[]0,[]0,[]0,[]0,[]1,[]0,"
text2 = "[]0,[]0,[]0,[]0,[]1,[]1,[]0,[]0,[]10,[]0,[]1,[]0,[]2,[]0,[]10,[]https://www.sankyutechnology.com,[]0,[]0,[]0,[]1,[]1,[]0,[]0,[]0,[]0,[]0,[]0,[]0,[]1,[]0,[]0,"

lastcharacter_text1 = InStrRev(text1, ",", , vbTextCompare)
lastcharacter_text2 = InStrRev(text2, ",", , vbTextCompare)

If lastcharacter_text2 - lastcharacter_text1 > 0 Then
    config_to_add = Right(text2, lastcharacter_text2 - lastcharacter_text1)
End If

MsgBox config_to_add
End Sub



Private Sub L12_Text_Click()
'On Error Resume Next
Frm3.cmdlogin.Visible = False
Frm3.CMD1.Visible = False

Frm3.CMD2.Visible = True
Frm3.CMD3.Visible = True

Frm3.Show 1
End Sub
Private Sub L17_Text_Change()
'On Error Resume Next
If G_SYSTEM_TYPE = "ONLINE" And MDI_frm1.L17_Text = "OFFLINE" Then
    Call MDI_frm1_unload_all_menu
    
    'MsgBox "Tiada sambungan sistem dan internet. OFFLINE." & vbCrLf & _
            vbNullString & vbrlf & _
            "Semua menu telah ditutup.", vbCritical, "Connecti'On Error"
    
End If
End Sub

Private Sub L17_Text_Click()
'On Error Resume Next
If MDI_frm1.L17_Text = "ONLINE" Then
    G_MAIN_CONN = 1
ElseIf MDI_frm1.L17_Text = "OFFLINE" Then
    G_MAIN_CONN = 0
End If
End Sub

Private Sub MDIForm_Load()
'On Error Resume Next
xLogOff = False

MDI_frm1.L3_Text = "-----------------" 'User
MDI_frm1.L4_Text = "-----------------" 'Level

'MDI_frm1.Picture = LoadPicture(App.Path & "\Backgrounds\1.jpg")

MDI_frm1.L5_Text = 0

MDI_frm1.L23_Text = "0"
MDI_frm1.L18_Text = "0"
MDI_frm1.L19_Text = "0"
MDI_frm1.L22_Text = "0"
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
'On Error Resume Next

'### Tutup form dengan klik [X] ### - Start
If xLogOff = False Then

    MDI_frm1.L5_Text = 0

    Note = "Adakah anda ingin menutup sistem ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    If Answer = vbYes Then
        End
    End If
    
End If

'### Tutup form dengan klik [X] ### - End

End Sub
Private Sub Tmr1_Timer()
'On Error Resume Next
MDI_frm1.L1_Text = DateTime.Date
MDI_frm1.L2_Text = DateTime.Time
End Sub

Private Sub Tmr2_Timer()
'On Error Resume Next
If MDI_frm1.L17_Text = "OFFLINE" Then
    
    If MDI_frm1.Image3.Visible = True Then
        MDI_frm1.Image3.Visible = False
    Else
        MDI_frm1.Image3.Visible = True
    End If
    
End If
End Sub

Private Sub Tmr3_Timer()
'On Error Resume Next
Call check_internet_interval
End Sub

Private Sub Tmr4_Timer()
'On Error GoTo logging:
Dim LM_INTERVAL As Double

LM_INTERVAL = G_INTERVAL_RESET

If MDI_frm1.Tmr4.Interval = LM_INTERVAL Then
    G_X = 0
End If

Exit Sub

logging:

G_LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " MDI_frm1 : Tmr4_Timer" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error Resume Next
MDI_LM_INDEX = Button.Index

If MDI_LM_INDEX = 1 Then
    Call MDI_frm1_unload_all_menu

    frm131.LV1.ListItems.Clear
    
    With frm131.LV1
        Set .SmallIcons = MDI_frm1.ImageList3
        Set .Icons = MDI_frm1.ImageList3
'Setting
        .ListItems.Add , "Senarai Client", "Senarai Client", 48
        .ListItems.Add , "Log", "Log", 37
    End With
    
    frm131.Show
ElseIf MDI_LM_INDEX = 2 Then
    Note = "Adakah anda ingin KELUAR dari sistem ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
        xLogOff = True
        Call frm302_reset_all
        frm302.Show
        frm302.TB1(0).SetFocus
        Unload MDI_frm1
    End If
ElseIf MDI_LM_INDEX = 3 Then
    Note = "Adakah anda ingin TUTUP dari sistem ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
        End
    End If
End If

Exit Sub
End Sub
Private Function HexString(EvalString As String) As String
Dim intStrLen As Integer
Dim intLoop As Integer
Dim strHex As String

HexString ("U-2403QXMG")

EvalString = Trim(EvalString)
intStrLen = Len(EvalString)
For intLoop = 1 To intStrLen
strHex = strHex & " " & Right$("0" & Hex(Asc(Mid(EvalString, intLoop, 1))), 2)
Next
HexString = strHex

MsgBox HexString
End Function
