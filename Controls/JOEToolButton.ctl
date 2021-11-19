VERSION 5.00
Begin VB.UserControl JOEToolButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5F5F5&
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   Begin VB.Timer timerMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4410
      Top             =   1080
   End
   Begin VB.Line ll 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1024
   End
   Begin VB.Line lr 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   1024
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00808080&
      Height          =   1095
      Left            =   3060
      Top             =   150
      Width           =   1605
   End
   Begin VB.Image imgHand 
      Height          =   480
      Left            =   210
      Picture         =   "JOEToolButton.ctx":0000
      Top             =   1350
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon2 
      Height          =   585
      Left            =   1470
      Top             =   1680
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JoeToolButton"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   930
      TabIndex        =   0
      Top             =   285
      Width           =   1035
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Left            =   75
      Top             =   75
      Width           =   825
   End
End
Attribute VB_Name = "JOEToolButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Event Click()

Dim m_ForeColor As OLE_COLOR
'Default Property Values:
Const m_def_BgColorHover = &HFFFFFF
Const m_def_BgColorNormal = &HF5F5F5
Const m_def_BgColorDisabled = &HF5F5F5
Const m_def_BgColorDown = &HE0E0E0
'Property Variables:
Dim m_BgColorHover As OLE_COLOR
Dim m_BgColorNormal As OLE_COLOR
Dim m_BgColorDisabled As OLE_COLOR
Dim m_BgColorDown As OLE_COLOR

Dim dOnDown As Boolean





Private Function Ctl_OnDown()
    
    
    dOnDown = True
    
    UserControl.BackColor = m_BgColorDown
    UserControl.Refresh
    DoEvents
      
    Dim s As Single
    s = GetTickCount + 40
    While GetTickCount < s
    Wend
    

End Function

Private Sub CtlMouseOver()

    UserControl.Parent.MousePointer = vbCustom
    UserControl.Parent.MouseIcon = imgHand.Picture
   
    If dOnDown = False Then
        UserControl.BackColor = m_BgColorHover
    End If
    
    timerMouse.Enabled = True

End Sub


Private Sub Ctl_OnUp()

    Dim p As POINTAPI
    Dim R As RECT

    dOnDown = False
    
    GetWindowRect UserControl.hwnd, R
    GetCursorPos p
    
    UserControl.BackColor = m_BgColorNormal
    
    If Not (p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom) Then
        RaiseEvent Click
    End If
End Sub

Private Sub timerMouse_Timer()
    Dim p As POINTAPI
    Dim R As RECT

    GetWindowRect UserControl.hwnd, R
    GetCursorPos p
    
    If p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom Then
        timerMouse.Enabled = False
            'out
            UserControl.BackColor = m_BgColorNormal
            
            UserControl.Parent.MousePointer = vbDefault
    End If
End Sub







Private Sub imgIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ctl_OnDown
End Sub

Private Sub imgIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CtlMouseOver
End Sub

Private Sub imgIcon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ctl_OnUp
End Sub



Private Sub imgIcon2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ctl_OnUp
End Sub


Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ctl_OnDown
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CtlMouseOver
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ctl_OnUp
End Sub




Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ctl_OnDown
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CtlMouseOver
End Sub



Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ctl_OnUp
End Sub

Private Sub UserControl_Resize()

    shpBorder.Move 0, 0, GetWidth, GetHeight
    
    If imgIcon.Left + imgIcon.Width + 3 < GetWidth Then
        lblCaption.Move imgIcon.Left + imgIcon.Width + 3, (GetHeight / 2) - (lblCaption.Height / 2)
    Else
        lblCaption.Move (GetWidth - lblCaption.Height) / 2, (GetHeight / 2) - (lblCaption.Height / 2)
    End If
    
End Sub



Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgIcon,imgIcon,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = imgIcon.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set imgIcon.Picture = New_Picture
    UserControl_Resize
    PropertyChanged "Picture"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set imgIcon.Picture = PropBag.ReadProperty("Picture", Nothing)
    lblCaption.Alignment = PropBag.ReadProperty("Alignment", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "JOEToolButton")
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.FontBold = PropBag.ReadProperty("FontBold", 0)
    lblCaption.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    lblCaption.FontName = PropBag.ReadProperty("FontName", lblCaption.FontName)
    lblCaption.FontSize = PropBag.ReadProperty("FontSize", lblCaption.FontSize)
    lblCaption.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_ForeColor = lblCaption.ForeColor
    lblCaption.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set imgIcon2.Picture = PropBag.ReadProperty("DisabledPicture", Nothing)
    m_BgColorHover = PropBag.ReadProperty("BgColorHover", m_def_BgColorHover)
    m_BgColorNormal = PropBag.ReadProperty("BgColorNormal", m_def_BgColorNormal)
    m_BgColorDisabled = PropBag.ReadProperty("BgColorDisabled", m_def_BgColorDisabled)
    m_BgColorDown = PropBag.ReadProperty("BgColorDown", m_def_BgColorDown)
    shpBorder.BorderColor = PropBag.ReadProperty("BorderColor", 8421504)
    shpBorder.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    shpBorder.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Picture", imgIcon.Picture, Nothing)
    Call PropBag.WriteProperty("Alignment", lblCaption.Alignment, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "JOEToolButton")
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", lblCaption.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", lblCaption.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", lblCaption.FontName, "")
    Call PropBag.WriteProperty("FontSize", lblCaption.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", lblCaption.FontStrikethru, 0)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("FontUnderline", lblCaption.FontUnderline, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("DisabledPicture", imgIcon2.Picture, Nothing)
    Call PropBag.WriteProperty("BgColorHover", m_BgColorHover, m_def_BgColorHover)
    Call PropBag.WriteProperty("BgColorNormal", m_BgColorNormal, m_def_BgColorNormal)
    Call PropBag.WriteProperty("BgColorDisabled", m_BgColorDisabled, m_def_BgColorDisabled)
    Call PropBag.WriteProperty("BgColorDown", m_BgColorDown, m_def_BgColorDown)
    Call PropBag.WriteProperty("BorderColor", shpBorder.BorderColor, 8421504)
    Call PropBag.WriteProperty("BorderStyle", shpBorder.BorderStyle, 1)
    Call PropBag.WriteProperty("BorderWidth", shpBorder.BorderWidth, 1)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Alignment
Public Property Get Alignment() As Integer
    Alignment = lblCaption.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    lblCaption.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    UserControl_Resize
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontBold
Public Property Get FontBold() As Boolean
    FontBold = lblCaption.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    lblCaption.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontItalic
Public Property Get FontItalic() As Boolean
    FontItalic = lblCaption.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    lblCaption.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontName
Public Property Get FontName() As String
    FontName = lblCaption.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    lblCaption.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontSize
Public Property Get FontSize() As Single
    FontSize = lblCaption.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    lblCaption.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = lblCaption.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    lblCaption.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
    FontUnderline = lblCaption.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    lblCaption.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    
    If New_Enabled = True Then
        lblCaption.ForeColor = m_ForeColor
        imgIcon.Visible = True
        imgIcon2.Visible = False
        UserControl.BackColor = m_BgColorNormal
    Else
        lblCaption.ForeColor = &HBBCACD
        imgIcon.Visible = False
        imgIcon2.Move imgIcon.Top, imgIcon.Left
        imgIcon2.Visible = True
        UserControl.BackColor = m_BgColorDisabled
    End If
    
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgIcon2,imgIcon2,-1,Picture
Public Property Get DisabledPicture() As Picture
Attribute DisabledPicture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set DisabledPicture = imgIcon2.Picture
End Property

Public Property Set DisabledPicture(ByVal New_DisabledPicture As Picture)
    Set imgIcon2.Picture = New_DisabledPicture
    PropertyChanged "DisabledPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BgColorHover() As OLE_COLOR
    BgColorHover = m_BgColorHover
End Property

Public Property Let BgColorHover(ByVal New_BgColorHover As OLE_COLOR)
    m_BgColorHover = New_BgColorHover
    PropertyChanged "BgColorHover"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BgColorNormal() As OLE_COLOR
    BgColorNormal = m_BgColorNormal
End Property

Public Property Let BgColorNormal(ByVal New_BgColorNormal As OLE_COLOR)
    m_BgColorNormal = New_BgColorNormal
    PropertyChanged "BgColorNormal"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BgColorDisabled() As OLE_COLOR
    BgColorDisabled = m_BgColorDisabled
End Property

Public Property Let BgColorDisabled(ByVal New_BgColorDisabled As OLE_COLOR)
    m_BgColorDisabled = New_BgColorDisabled
    PropertyChanged "BgColorDisabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BgColorDown() As OLE_COLOR
    BgColorDown = m_BgColorDown
End Property

Public Property Let BgColorDown(ByVal New_BgColorDown As OLE_COLOR)
    m_BgColorDown = New_BgColorDown
    PropertyChanged "BgColorDown"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BgColorHover = m_def_BgColorHover
    m_BgColorNormal = m_def_BgColorNormal
    m_BgColorDisabled = m_def_BgColorDisabled
    m_BgColorDown = m_def_BgColorDown
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpBorder,shpBorder,-1,BorderColor
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = shpBorder.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    shpBorder.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpBorder,shpBorder,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = shpBorder.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    shpBorder.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpBorder,shpBorder,-1,BorderWidth
Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "Returns or sets the width of a control's border."
    BorderWidth = shpBorder.BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    shpBorder.BorderWidth() = New_BorderWidth
    PropertyChanged "BorderWidth"
End Property

