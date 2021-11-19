VERSION 5.00
Begin VB.UserControl JoeWinTab 
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   ScaleHeight     =   76
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   Begin VB.Timer timerMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -15
      Top             =   270
   End
   Begin VB.Image imgClose 
      Height          =   240
      Left            =   1890
      Picture         =   "JoeWinTab.ctx":0000
      ToolTipText     =   "Close Window"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgLeft 
      Height          =   360
      Index           =   3
      Left            =   4320
      Picture         =   "JoeWinTab.ctx":058A
      Top             =   120
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgRight 
      Height          =   360
      Index           =   3
      Left            =   5310
      Picture         =   "JoeWinTab.ctx":07AC
      Top             =   90
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgCenter 
      Height          =   360
      Index           =   3
      Left            =   4860
      Picture         =   "JoeWinTab.ctx":09CE
      Top             =   90
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image imgHand 
      Height          =   480
      Left            =   3570
      Picture         =   "JoeWinTab.ctx":0A70
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCenter 
      Height          =   360
      Index           =   2
      Left            =   2745
      Picture         =   "JoeWinTab.ctx":133A
      Stretch         =   -1  'True
      Top             =   570
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Image imgRight 
      Height          =   360
      Index           =   2
      Left            =   3780
      Picture         =   "JoeWinTab.ctx":13DC
      Top             =   690
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgLeft 
      Height          =   360
      Index           =   2
      Left            =   2610
      Picture         =   "JoeWinTab.ctx":15FE
      Top             =   750
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgCenter 
      Height          =   360
      Index           =   1
      Left            =   5430
      Picture         =   "JoeWinTab.ctx":1820
      Top             =   540
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image imgRight 
      Height          =   360
      Index           =   1
      Left            =   5850
      Picture         =   "JoeWinTab.ctx":18C2
      Top             =   570
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgRight 
      Height          =   360
      Index           =   0
      Left            =   2190
      Picture         =   "JoeWinTab.ctx":1AE4
      Top             =   120
      Width           =   90
   End
   Begin VB.Image imgLeft 
      Height          =   360
      Index           =   1
      Left            =   4890
      Picture         =   "JoeWinTab.ctx":1D06
      Top             =   570
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgLeft 
      Height          =   360
      Index           =   0
      Left            =   180
      Picture         =   "JoeWinTab.ctx":1F28
      Top             =   60
      Width           =   90
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JOEWinTab"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C25418&
      Height          =   375
      Left            =   750
      TabIndex        =   0
      Top             =   540
      Width           =   855
   End
   Begin VB.Image imgCenter 
      Height          =   360
      Index           =   0
      Left            =   840
      Picture         =   "JoeWinTab.ctx":214A
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1095
   End
End
Attribute VB_Name = "JoeWinTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'JOEWinTab
'by: Joehel V. Del Rosario
'
'Created: 9:04 pm  June 11, 2006
'Modified: 11:20 pm June 11,2006

Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Event Click()
Event CloseClick()

'Default Property Values:
Const m_def_CloseButton = True
Const m_def_Value = False
'Const m_def_Value = False
'Property Variables:
Dim m_CloseButton As Boolean
Dim m_Value As Boolean
'Dim m_Value As Boolean

Dim OnDown As Boolean


Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function




Private Sub imgCenter_Click(Index As Integer)
    RaiseEvent Click
End Sub

Private Sub imgCenter_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgLeft(0).Picture = imgLeft(2).Picture
    Set imgCenter(0).Picture = imgCenter(2).Picture
    Set imgRight(0).Picture = imgRight(2).Picture
    OnDown = True
    DoEvents
End Sub

Private Sub imgCenter_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtlMouseOver
End Sub

Private Sub imgCenter_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgLeft(0).Picture = imgLeft(3).Picture
    Set imgCenter(0).Picture = imgCenter(3).Picture
    Set imgRight(0).Picture = imgRight(3).Picture
    OnDown = False
End Sub

Private Sub imgClose_Click()
    RaiseEvent CloseClick
End Sub

Private Sub imgLeft_Click(Index As Integer)
    RaiseEvent Click
End Sub

Private Sub imgLeft_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgLeft(0).Picture = imgLeft(2).Picture
    Set imgCenter(0).Picture = imgCenter(2).Picture
    Set imgRight(0).Picture = imgRight(2).Picture
    OnDown = True
    DoEvents
End Sub

Private Sub imgLeft_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtlMouseOver
End Sub

Private Sub imgLeft_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgLeft(0).Picture = imgLeft(3).Picture
    Set imgCenter(0).Picture = imgCenter(3).Picture
    Set imgRight(0).Picture = imgRight(3).Picture
    OnDown = False
End Sub

Private Sub imgRight_Click(Index As Integer)
    RaiseEvent Click
End Sub

Private Sub imgRight_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgLeft(0).Picture = imgLeft(2).Picture
    Set imgCenter(0).Picture = imgCenter(2).Picture
    Set imgRight(0).Picture = imgRight(2).Picture
    OnDown = True
    DoEvents
End Sub

Private Sub imgRight_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtlMouseOver
End Sub

Private Sub imgRight_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgLeft(0).Picture = imgLeft(3).Picture
    Set imgCenter(0).Picture = imgCenter(3).Picture
    Set imgRight(0).Picture = imgRight(3).Picture
    OnDown = False
End Sub

Private Sub lblCaption_Click()
    RaiseEvent Click
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgLeft(0).Picture = imgLeft(2).Picture
    Set imgCenter(0).Picture = imgCenter(2).Picture
    Set imgRight(0).Picture = imgRight(2).Picture
    OnDown = True
    DoEvents
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtlMouseOver
End Sub


Private Sub CtlMouseOver()

    If OnDown = True Then
        Exit Sub
    End If
    
    'move close button
    imgClose.Move GetWidth - imgClose.Width - 4, 2
    
    Set imgLeft(0).Picture = imgLeft(3).Picture
    Set imgCenter(0).Picture = imgCenter(3).Picture
    Set imgRight(0).Picture = imgRight(3).Picture
    
    On Error Resume Next
    UserControl.Parent.MouseIcon = imgHand.Picture
    UserControl.Parent.MousePointer = vbCustom
    UserControl.MouseIcon = imgHand.Picture
    UserControl.MousePointer = vbCustom

    

    timerMouse.Enabled = True
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgLeft(0).Picture = imgLeft(3).Picture
    Set imgCenter(0).Picture = imgCenter(3).Picture
    Set imgRight(0).Picture = imgRight(3).Picture
    OnDown = False

End Sub

Private Sub timerMouse_Timer()
    Dim p As POINTAPI
    Dim r As RECT

    If OnDown = True Then
        Exit Sub
    End If
    
    GetWindowRect UserControl.hwnd, r
    GetCursorPos p
    
    If p.X < r.Left Or p.X > r.Right Or p.Y < r.Top Or p.Y > r.Bottom Then
        timerMouse.Enabled = False
        
        'restore close button position
        imgClose.Move GetWidth - imgClose.Width - 2, 0
        
        If Value = True Then
            Set imgLeft(0).Picture = imgLeft(1).Picture
            Set imgCenter(0).Picture = imgCenter(1).Picture
            Set imgRight(0).Picture = imgRight(1).Picture
        Else
            Set imgLeft(0).Picture = imgLeft(2).Picture
            Set imgCenter(0).Picture = imgCenter(2).Picture
            Set imgRight(0).Picture = imgRight(2).Picture
        End If
                
        On Error Resume Next
        UserControl.MousePointer = vbDefault
        UserControl.Parent.MousePointer = vbDefault
    End If

End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl.Height = imgLeft(0).Height * Screen.TwipsPerPixelY
    imgLeft(0).Move 0, 0
    imgRight(0).Move GetWidth - imgRight(0).Width, 0
    imgCenter(0).Move imgLeft(0).Width, 0, GetWidth - imgLeft(0).Width - imgRight(0).Width
    lblCaption.Move (GetWidth / 2) - (lblCaption.Width / 2), (GetHeight / 2) - (lblCaption.Height / 2)
    
    imgClose.Move GetWidth - imgClose.Width - 2, 0
    
    If lblCaption.Left < imgLeft(0).Width Then
        lblCaption.Left = imgLeft(0).Width
    End If

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_Value = m_def_Value
    m_Value = m_def_Value
    m_CloseButton = m_def_CloseButton
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "JOEWinTab")
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_CloseButton = PropBag.ReadProperty("CloseButton", m_def_CloseButton)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "JOEWinTab")
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("CloseButton", m_CloseButton, m_def_CloseButton)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    
    'tool tips
    imgClose.ToolTipText = "Close " & New_Caption & " window"
    lblCaption.ToolTipText = New_Caption
    
    UserControl_Resize
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    UserControl_Resize
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    m_Value = New_Value
    If Value = True Then
        Set imgLeft(0).Picture = imgLeft(1).Picture
        Set imgCenter(0).Picture = imgCenter(1).Picture
        Set imgRight(0).Picture = imgRight(1).Picture
    Else
        Set imgLeft(0).Picture = imgLeft(2).Picture
        Set imgCenter(0).Picture = imgCenter(2).Picture
        Set imgRight(0).Picture = imgRight(2).Picture
    End If

    PropertyChanged "Value"
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get CloseButton() As Boolean
    CloseButton = m_CloseButton
End Property

Public Property Let CloseButton(ByVal New_CloseButton As Boolean)
    m_CloseButton = New_CloseButton
    imgClose.Visible = New_CloseButton
    PropertyChanged "CloseButton"
End Property

