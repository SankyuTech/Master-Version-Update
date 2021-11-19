VERSION 5.00
Begin VB.UserControl JOeMenu 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5025
   KeyPreview      =   -1  'True
   ScaleHeight     =   95
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   335
   Begin VB.Timer timerMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   0
   End
   Begin VB.Line Line4 
      X1              =   36
      X2              =   76
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Line Line3 
      X1              =   76
      X2              =   76
      Y1              =   64
      Y2              =   32
   End
   Begin VB.Line Line2 
      X1              =   36
      X2              =   76
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Line Line1 
      X1              =   42
      X2              =   42
      Y1              =   34
      Y2              =   56
   End
   Begin VB.Image imgHand 
      Height          =   480
      Left            =   2040
      Picture         =   "JoeMenu.ctx":0000
      Top             =   690
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "JOeMenu"
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

'Property Variables:
Dim m_Font As Font
Dim m_Caption As String
Dim m_BorderColorHover As OLE_COLOR
Dim m_BorderColorNormal As OLE_COLOR
Dim m_BorderColorDown As OLE_COLOR
Dim m_BackColorHover As OLE_COLOR
Dim m_BackColorNormal As OLE_COLOR
Dim m_BackColorDown As OLE_COLOR
Dim m_Menu As Menu


Dim dOnDown As Boolean
Dim bOnPopUp As Boolean

Private WithEvents lblCaption As Label
Attribute lblCaption.VB_VarHelpID = -1
'Default Property Values:
Const m_def_ForeColor = &HE7D3C7
Const m_def_Caption = "JOE Menu"
Const m_def_BorderColorHover = &HA85E33
Const m_def_BorderColorNormal = &H8000000F
Const m_def_BorderColorDown = &H8000000F

Const m_def_BackColorHover = &HE7D3C7
Const m_def_BackColorNormal = &HEDEBE9
Const m_def_BackColorDown = &H8000000F


Public Event Click()

Private Sub DrawNormal()

    UserControl.BackColor = m_BackColorNormal
        
    Line1.BorderColor = m_BorderColorNormal
    Line2.BorderColor = m_BorderColorNormal
    Line3.BorderColor = m_BorderColorNormal
    Line4.BorderColor = m_BorderColorNormal
    
    On Error Resume Next
    UserControl.Parent.MousePointer = vbDefault
    Err.Clear

End Sub

Private Sub DrawDown()

    UserControl.BackColor = m_BackColorDown
        
    Line1.BorderColor = m_BorderColorDown
    Line2.BorderColor = m_BorderColorDown
    Line3.BorderColor = m_BorderColorDown
    Line4.BorderColor = m_BorderColorDown
    
End Sub

Private Sub DrawHover()

    UserControl.BackColor = m_BackColorHover
        
    Line1.BorderColor = m_BorderColorHover
    Line2.BorderColor = m_BorderColorHover
    Line3.BorderColor = m_BorderColorHover
    Line4.BorderColor = m_BorderColorHover
End Sub

Private Function Ctl_OnDown()
    
    
    dOnDown = True
    
    Call DrawDown
        
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
        
        Call DrawHover

    End If
    
    timerMouse.Enabled = True

End Sub


Private Sub Ctl_OnUp()

    Dim p As POINTAPI
    Dim R As RECT

    dOnDown = False
    
    GetWindowRect UserControl.hwnd, R
    GetCursorPos p

    If Not (p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom) Then
        
        RaiseEvent Click
        
        bOnPopUp = True
        
        ShowPopUp
        
        bOnPopUp = False
        
    End If
    
    DrawNormal
    
Errh:
    Err.Clear
    
End Sub


Public Sub ShowPopUp()

    Dim R As RECT
    
    Dim iX As Integer
    Dim iY As Integer
    
    
    If IsObject(m_Menu) Then
        On Error Resume Next
        
        GetWindowRect UserControl.Parent.hwnd, R

        iX = 0 'R.Left * 15
        iY = GetHeight
        
        Call DrawDown
           
        UserControl.PopupMenu m_Menu, , iX, iY
    
        Call DrawNormal
    End If
    
End Sub



Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ctl_OnDown
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtlMouseOver
End Sub

Private Sub timerMouse_Timer()
    
    Dim p As POINTAPI
    Dim R As RECT

    If bOnPopUp = True Then
        Exit Sub
    End If
    
    GetWindowRect UserControl.hwnd, R
    GetCursorPos p
    
    If p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom Then
            timerMouse.Enabled = False
            'out
            Call DrawNormal
    End If

End Sub




Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Ctl_OnUp

End Sub


Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function

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

Private Sub UserControl_GotFocus()
    Call DrawHover
End Sub

Private Sub UserControl_Initialize()

    Set lblCaption = UserControl.Controls.Add("VB.Label", "lblCaption")
    With lblCaption
        .AutoSize = True
        .Alignment = vbCenter
        .BackStyle = vbTransparent
        .Caption = m_Caption
        .Move 0, 0
        .Visible = True
    End With
    
End Sub


Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ShowPopUp
    End If
End Sub

Private Sub UserControl_LostFocus()
    Call DrawNormal
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ctl_OnDown
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CtlMouseOver
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ctl_OnUp
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "Label1")
  
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_BorderColorHover = PropBag.ReadProperty("BorderColorHover", m_def_BorderColorHover)
    m_BorderColorNormal = PropBag.ReadProperty("BorderColorNormal", m_def_BorderColorNormal)
    m_BorderColorDown = PropBag.ReadProperty("BorderColorDown", m_def_BorderColorDown)
    m_BackColorHover = PropBag.ReadProperty("BackColorHover", m_def_BackColorHover)
    m_BackColorNormal = PropBag.ReadProperty("BackColorNormal", m_def_BackColorNormal)
    m_BackColorDown = PropBag.ReadProperty("BackColorDown", m_def_BackColorDown)

End Sub


Private Sub UserControl_Resize()
        
    Dim iH As Integer
    Dim iW As Integer
    
    iW = GetWidth
    iH = GetHeight
    
    On Error Resume Next
    
    With Line1
        .X1 = 0
        .X2 = 0
        .Y1 = 0
        .Y2 = iH
    End With
    
    With Line2
        .X1 = 0
        .X2 = iW
        .Y1 = 0
        .Y2 = 0
    End With
    
    With Line3
        .X1 = 0
        .X2 = iW
        .Y1 = iH - 1
        .Y2 = iH - 1
    End With
    
    With Line4
        .X1 = iW - 1
        .X2 = iW - 1
        .Y1 = 0
        .Y2 = iH
    End With
    
    lblCaption.Move 0, (iH - lblCaption.Height) / 2, GetWidth
    
    Err.Clear
End Sub

Private Sub UserControl_Show()

    Call DrawNormal
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Label1")
   
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("BorderColorHover", m_BorderColorHover, m_def_BorderColorHover)
    Call PropBag.WriteProperty("BorderColorNormal", m_BorderColorNormal, m_def_BorderColorNormal)
    Call PropBag.WriteProperty("BorderColorDown", m_BorderColorDown, m_def_BorderColorDown)
    Call PropBag.WriteProperty("BackColorHover", m_BackColorHover, m_def_BackColorHover)
    Call PropBag.WriteProperty("BackColorNormal", m_BackColorNormal, m_def_BackColorNormal)
    Call PropBag.WriteProperty("BackColorDown", m_BackColorDown, m_def_BackColorDown)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,0,0,0
Public Property Get Menu() As Object
    Set Menu = m_Menu
End Property

Public Property Set Menu(ByVal New_Menu As Object)
    Set m_Menu = New_Menu
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    lblCaption.Caption = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BorderColorHover() As OLE_COLOR
    BorderColorHover = m_BorderColorHover
End Property

Public Property Let BorderColorHover(ByVal New_BorderColorHover As OLE_COLOR)
    m_BorderColorHover = New_BorderColorHover
    PropertyChanged "BorderColorHover"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BorderColorNormal() As OLE_COLOR
    BorderColorNormal = m_BorderColorNormal
End Property

Public Property Let BorderColorNormal(ByVal New_BorderColorNormal As OLE_COLOR)
    m_BorderColorNormal = New_BorderColorNormal
    PropertyChanged "BorderColorNormal"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BorderColorDown() As OLE_COLOR
    BorderColorDown = m_BorderColorDown
End Property

Public Property Let BorderColorDown(ByVal New_BorderColorDown As OLE_COLOR)
    m_BorderColorDown = New_BorderColorDown
    PropertyChanged "BorderColorDown"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColorHover() As OLE_COLOR
    BackColorHover = m_BackColorHover
End Property

Public Property Let BackColorHover(ByVal New_BackColorHover As OLE_COLOR)
    m_BackColorHover = New_BackColorHover
    PropertyChanged "BackColorHover"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColorNormal() As OLE_COLOR
    BackColorNormal = m_BackColorNormal
End Property

Public Property Let BackColorNormal(ByVal New_BackColorNormal As OLE_COLOR)
    m_BackColorNormal = New_BackColorNormal
    PropertyChanged "BackColorNormal"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColorDown() As OLE_COLOR
    BackColorDown = m_BackColorDown
End Property

Public Property Let BackColorDown(ByVal New_BackColorDown As OLE_COLOR)
    m_BackColorDown = New_BackColorDown
    PropertyChanged "BackColorDown"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    lblCaption.ForeColor = m_def_ForeColor
    Set m_Font = Ambient.Font
    m_Caption = m_def_Caption
    m_BorderColorHover = m_def_BorderColorHover
    m_BorderColorNormal = m_def_BorderColorNormal
    m_BorderColorDown = m_def_BorderColorDown
    m_BackColorHover = m_def_BackColorHover
    m_BackColorNormal = m_def_BackColorNormal
    m_BackColorDown = m_def_BackColorDown
End Sub

