VERSION 5.00
Begin VB.UserControl JOETitleBar 
   Alignable       =   -1  'True
   BackColor       =   &H00808080&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8745
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   Begin VB.Timer timerCloseHover 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3780
      Top             =   0
   End
   Begin VB.Image imgClose 
      Height          =   360
      Index           =   1
      Left            =   4920
      Picture         =   "JoeTitleBar.ctx":0000
      Top             =   30
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JOe Title Bar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   165
      TabIndex        =   1
      Top             =   30
      Width           =   1200
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00C0C0C0&
      Height          =   405
      Left            =   1560
      Top             =   30
      Width           =   375
   End
   Begin VB.Image imgClose 
      Height          =   360
      Index           =   0
      Left            =   4470
      Picture         =   "JoeTitleBar.ctx":06EA
      Top             =   60
      Width           =   360
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JOe Title Bar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   150
      TabIndex        =   0
      Top             =   30
      Width           =   1200
   End
End
Attribute VB_Name = "JOETitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

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


'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."


Dim bOnDown As Boolean
'Property Variables:
Dim m_Image As Picture




Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelY
End Function
Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelX
End Function


Private Sub imgClose_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    If Button = vbLeftButton Then
        bOnDown = True
        imgClose(1).Visible = False
    End If
End Sub

Private Sub imgClose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If bOnDown = False Then
        'switch to hover view
        imgClose(1).Visible = True
        timerCloseHover.Enabled = True
    End If
    
End Sub

Private Sub imgClose_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim p As POINTAPI
    Dim R As RECT
    
    
    If bOnDown = True And Button = vbLeftButton Then
    
        GetWindowRect UserControl.hWnd, R
        GetCursorPos p
        
        'resize rect to state button size
        R.Left = R.Left + imgClose(0).Left
        R.Top = R.Top + imgClose(0).Top
        R.Right = R.Left + imgClose(0).Width
        R.Bottom = R.Top + imgClose(0).Height
        
        If Not (p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom) Then
                    
            'close button clicked
            On Error Resume Next
            Unload UserControl.Parent
            
        End If
        
    End If
    
    bOnDown = False
    timerCloseHover.Enabled = True

    
End Sub

Private Sub timerCloseHover_Timer()

    Dim p As POINTAPI
    Dim R As RECT
    
    GetWindowRect UserControl.hWnd, R
    GetCursorPos p
    
    'resize rect to state button size
    R.Left = R.Left + imgClose(0).Left
    R.Top = R.Top + imgClose(0).Top
    R.Right = R.Left + imgClose(0).Width
    R.Bottom = R.Top + imgClose(0).Height
    
    If p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom Then
                
        'restore normal view
        timerCloseHover.Enabled = False
        imgClose(1).Visible = False
        
    End If
    
End Sub

Private Sub UserControl_Resize()

    UserControl.AutoRedraw = False
    
    UserControl.Height = 375
    
    If GetWidth < 100 Then
        UserControl.Width = 100 * Screen.TwipsPerPixelX
    End If
    

    imgClose(0).Move GetWidth - imgClose(0).Width - 2, (GetHeight / 2) - (imgClose(0).Height / 2)
    imgClose(1).Move imgClose(0).Left, imgClose(0).Top
    
    lblCaption.Move 8, (GetHeight - lblCaption.Height) / 2, imgClose(0).Left - lblCaption.Left
    lblShadow.Move lblCaption.Left - 1, lblCaption.Top + 1, imgClose(0).Left - lblShadow.Left
    
    shpBorder.Move 0, 0, GetWidth, GetHeight
    lblCaption.ZOrder 0
    UserControl.AutoRedraw = True
    Refresh
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = lblCaption.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    lblCaption.Alignment() = New_Alignment
    lblShadow.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    lblShadow.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    Set lblShadow.Font = New_Font
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
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=imgCaptionBg,imgCaptionBg,-1,Image
'Public Property Get Image() As Picture
'    Set Image = imgCaptionBg.Picture
'End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblCaption.Alignment = PropBag.ReadProperty("Alignment", 0)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "JOE Title Bar")
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H808080)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    lblShadow.Alignment = PropBag.ReadProperty("Alignment", 0)
    lblShadow.Caption = PropBag.ReadProperty("Caption", "JOE Title Bar")
    lblShadow.ForeColor = PropBag.ReadProperty("ForeColor", &H808080)
    Set lblShadow.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set m_Image = PropBag.ReadProperty("Image", Nothing)
    lblShadow.ForeColor = PropBag.ReadProperty("ShadowColor", &HFFFFFF)
    shpBorder.Bordercolor = PropBag.ReadProperty("BorderColor", 12632256)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
End Sub


'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Alignment", lblCaption.Alignment, 0)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "JOE Title Bar")
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H808080)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Image", m_Image, Nothing)
    Call PropBag.WriteProperty("ShadowColor", lblShadow.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("BorderColor", shpBorder.Bordercolor, 12632256)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = m_Image
End Property

Public Property Set Image(ByVal New_Image As Picture)
    Set m_Image = New_Image
    PropertyChanged "Image"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblShadow,lblShadow,-1,ForeColor
Public Property Get ShadowColor() As OLE_COLOR
Attribute ShadowColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ShadowColor = lblShadow.ForeColor
End Property

Public Property Let ShadowColor(ByVal New_ShadowColor As OLE_COLOR)
    lblShadow.ForeColor() = New_ShadowColor
    PropertyChanged "ShadowColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set m_Image = LoadPicture("")
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpBorder,shpBorder,-1,BorderColor
Public Property Get Bordercolor() As OLE_COLOR
Attribute Bordercolor.VB_Description = "Returns/sets the color of an object's border."
    Bordercolor = shpBorder.Bordercolor
End Property

Public Property Let Bordercolor(ByVal New_BorderColor As OLE_COLOR)
    shpBorder.Bordercolor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

