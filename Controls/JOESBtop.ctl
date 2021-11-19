VERSION 5.00
Begin VB.UserControl JOESBtop 
   AutoRedraw      =   -1  'True
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   ClipBehavior    =   0  'None
   ControlContainer=   -1  'True
   ScaleHeight     =   63
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   263
   Begin VB.Timer timerStateButton 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1380
      Top             =   240
   End
   Begin VB.Image imgSB 
      Height          =   345
      Left            =   3480
      Picture         =   "JOESBtop.ctx":0000
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgStateButton 
      Height          =   345
      Index           =   3
      Left            =   1950
      Picture         =   "JOESBtop.ctx":0886
      Top             =   60
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgStateButton 
      Height          =   345
      Index           =   2
      Left            =   1950
      Picture         =   "JOESBtop.ctx":110C
      Top             =   390
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgStateButton 
      Height          =   345
      Index           =   1
      Left            =   3000
      Picture         =   "JOESBtop.ctx":1992
      Top             =   300
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgStateButton 
      Height          =   345
      Index           =   0
      Left            =   3480
      Picture         =   "JOESBtop.ctx":2218
      Top             =   330
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgRight 
      Height          =   945
      Left            =   3810
      MousePointer    =   9  'Size W E
      Picture         =   "JOESBtop.ctx":2A9E
      Top             =   0
      Width           =   120
   End
   Begin VB.Image imgLeft 
      Height          =   945
      Left            =   0
      Picture         =   "JOESBtop.ctx":30C8
      Top             =   0
      Width           =   120
   End
   Begin VB.Image imgCenter 
      Height          =   945
      Left            =   120
      Picture         =   "JOESBtop.ctx":36F2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3705
   End
End
Attribute VB_Name = "JOESBtop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
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

Public Enum eSizeState
    ssContracted = 0
    ssExpanded = 1
End Enum

'Default Property Values:
Const m_def_MinWidth = 160
Const m_def_SizeState = 0
'Property Variables:
Dim m_MinWidth As Integer
Dim m_SizeState As eSizeState

'events
Public Event SizeChange(ByVal newSizeState As eSizeState)
Public Event Resize()

Dim bMouseDown As Boolean
Dim iMX As Integer


Private Sub imgRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMouseDown = True
    iMX = X
End Sub

Private Sub imgRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim iNewWidth As Integer
    
    If bMouseDown = True Then
        
        On Error Resume Next
            
        iNewWidth = UserControl.Width + (X - iMX)
        
        Debug.Print iNewWidth
        
        If iNewWidth >= m_MinWidth * Screen.TwipsPerPixelX Then
            UserControl.Width = iNewWidth
        Else
            'bMouseDown = False
        End If
        
        'RaiseEvent Resize
        
    End If
    
    
End Sub

Private Sub imgRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMouseDown = False
End Sub

Private Sub imgSB_Click()
    'change state
    If m_SizeState = ssContracted Then
        m_SizeState = ssExpanded
    Else
        m_SizeState = ssContracted
    End If
    
    RaiseEvent SizeChange(m_SizeState)
End Sub

Private Sub imgSB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call StateButMHover
End Sub

Private Sub timerStateButton_Timer()
    
    Dim p As POINTAPI
    Dim R As RECT
    
    GetWindowRect UserControl.hwnd, R
    GetCursorPos p
    
    'resize rect to state button size
    R.Left = R.Left + imgSB.Left
    R.Bottom = R.Top + imgSB.Height
    
    If p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom Then
                
        StateButMLeave
                
    End If
End Sub

Private Sub StateButMHover()

    Call RefreshSBButton
    
    timerStateButton.Enabled = True
    
End Sub

Private Sub StateButMLeave()

    timerStateButton.Enabled = False
    
    Call RefreshSBButton
    
End Sub

Private Sub RefreshSBButton()
    If timerStateButton.Enabled = True Then
        Call ShowHover
    Else
        Call ShowNormal
    End If
End Sub
Private Sub ShowNormal()
    If m_SizeState = ssContracted Then
        'contracted
        Set imgSB.Picture = imgStateButton(0).Picture
    Else
        'expanded
        Set imgSB.Picture = imgStateButton(2).Picture
    End If
End Sub
Private Sub ShowHover()
    If m_SizeState = ssContracted Then
        'contracted
        Set imgSB.Picture = imgStateButton(1).Picture
    Else
        'expanded
        Set imgSB.Picture = imgStateButton(3).Picture
    End If
End Sub










Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function


Private Sub UserControl_Resize()
    

    If GetWidth < m_MinWidth Then
        UserControl.Width = m_MinWidth * Screen.TwipsPerPixelX
    End If
    UserControl.Height = 63 * Screen.TwipsPerPixelY

    imgLeft.Move 0, 0
    imgRight.Move GetWidth - imgRight.Width
    imgCenter.Move imgLeft.Left + imgLeft.Width, 0, imgRight.Left - (imgLeft.Left + imgLeft.Width)
    
    
    imgSB.Move GetWidth - imgSB.Width, 0

    RaiseEvent Resize
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get SizeState() As eSizeState
    SizeState = m_SizeState
End Property

Public Property Let SizeState(ByVal New_SizeState As eSizeState)
    
    m_SizeState = New_SizeState

    Call RefreshSBButton
    
    PropertyChanged "SizeState"
    
    RaiseEvent SizeChange(m_SizeState)
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_SizeState = m_def_SizeState
    m_MinWidth = m_def_MinWidth
End Sub

Private Sub UserControl_Show()
    ShowNormal
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_SizeState = PropBag.ReadProperty("SizeState", m_def_SizeState)
    m_MinWidth = PropBag.ReadProperty("MinWidth", m_def_MinWidth)
End Sub



'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("SizeState", m_SizeState, m_def_SizeState)
    Call PropBag.WriteProperty("MinWidth", m_MinWidth, m_def_MinWidth)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,160
Public Property Get MinWidth() As Integer
    MinWidth = m_MinWidth
End Property

Public Property Let MinWidth(ByVal New_MinWidth As Integer)
    m_MinWidth = New_MinWidth
    PropertyChanged "MinWidth"
End Property

