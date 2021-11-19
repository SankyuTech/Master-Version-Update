VERSION 5.00
Begin VB.UserControl JOEClientWin 
   Alignable       =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7950
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   530
   Begin MOVERS.JOEWinTabs JOEWTabs 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgBR 
      Height          =   1605
      Left            =   0
      MousePointer    =   9  'Size W E
      Picture         =   "JOEClientWin.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   90
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   75
      Picture         =   "JOEClientWin.ctx":9A52
      Stretch         =   -1  'True
      Top             =   15
      Width           =   19995
   End
End
Attribute VB_Name = "JOEClientWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'declare

Private Const SW_HIDE As Long = 0           ' Hide window constant
Private Const SW_SHOW As Long = 5           ' Show window constant

Private Const WS_EX_MDICHILD As Long = &H40&
Private Const WS_EX_WINDOWEDGE As Long = &H100&
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_BORDER As Long = &H800000
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CHILDWINDOW As Long = (WS_CHILD)
Private Const WM_MOUSEACTIVATE As Long = &H21
'private Const WM_CLOSE As Long = &H10
Private Const WM_COMMAND As Long = &H111

Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const WS_EX_NOPARENTNOTIFY As Long = &H4&

Private Const GWL_STYLE As Long = (-16)
Private Const GWL_HWNDPARENT As Long = -8

Private Const SWP_FRAMECHANGED As Long = &H20 ' Set window position constant - sends message frame changed to the window

' Rectangle
Private Type RECT
   Left As Long     ' Left of the rectangle
   Top As Long      ' Top of the rectangle
   Right As Long    ' Right of the rectangle
   Bottom As Long   ' Bottom of the rectangle
End Type

' Point
Private Type POINTAPI
   X As Long        ' X position of the point.
   Y As Long        ' Y position of the point.
End Type

' Window border style constants.
Private Enum VbWindowStyle
   VbNone = 0           ' No border
   VbToolWin = 1        ' Tool window
End Enum

Private Declare Function GetWindowRect Lib "user32.dll" ( _
   ByVal hwnd As Long, _
   ByRef lpRect As RECT _
) As Long

Private Declare Function DeleteDC Lib "gdi32" ( _
   ByVal hDc As Long _
) As Long

Private Declare Function DrawFocusRect Lib "user32" ( _
   ByVal hDc As Long, _
   lpRect As RECT _
) As Long

Private Declare Function ShowWindow Lib "user32.dll" ( _
   ByVal hwnd As Long, _
   ByVal nCmdShow As Long _
) As Long


Private Declare Function SetWindowPos Lib "user32.dll" ( _
   ByVal hwnd As Long, _
   ByVal hWndInsertAfter As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal wFlags As Long _
) As Long

Private Declare Function SetParent Lib "user32.dll" ( _
   ByVal hWndChild As Long, _
   ByVal hWndNewParent As Long _
) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
 ) As Long
 
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
   ByVal hwnd As Long, _
   ByVal nIndex As Long _
) As Long


Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long



'events
Public Event FormTabClick(ByVal sFormName As String, ByVal Index As Integer)
Public Event CloseClick(ByVal sFormName As String, ByVal Index As Integer)
Private m_FormStyle As Long


'Default Property Values:
Const m_def_SBWidth = 0
'Property Variables:
Dim m_SBWidth As Integer











Private Sub SetWindowStyle(ByVal hwnd As Long, ByVal BorderStyle As VbWindowStyle)
     
   Dim l_Style As Long
  
   If m_FormStyle = 0 Then
      m_FormStyle = GetWindowLong(hwnd, GWL_STYLE)
   End If
   
   ' set new window style
   If BorderStyle = VbNone Then
      
      l_Style = m_FormStyle And Not WS_DLGFRAME And Not WS_EX_APPWINDOW _
                            And Not WS_BORDER And Not WS_EX_WINDOWEDGE Or _
                            WS_EX_MDICHILD Or WS_CHILDWINDOW And Not WS_EX_NOPARENTNOTIFY
   
      SetWindowLong hwnd, GWL_STYLE, l_Style
      
   Else
   
      l_Style = GetWindowLong(hwnd, GWL_STYLE) ' Get current style
      l_Style = l_Style And Not WS_CAPTION Or WS_EX_NOPARENTNOTIFY
   
      SetWindowLong hwnd, GWL_STYLE, l_Style
                  
   End If

End Sub


Public Sub LoadChildWindow(ByRef ParentHwnd As Long, _
                            ByRef WinhWnd As Long, _
                            ByVal sFormName As String, _
                            ByVal sFormCaption As String, _
                            ByVal iTop As Integer, _
                            ByVal iLeft As Integer, _
                            ByVal iRight As Integer, _
                            ByVal iBottom As Integer, Optional ShowCloseButton As Boolean = True)


    'add window to tab
    JOEWTabs.AddForm sFormName, sFormCaption, ShowCloseButton
    JOEWTabs.SetForm sFormName
    
    'remove form's border
    SetWindowStyle WinhWnd, VbNone
    
    'resize form
    SetWindowPos WinhWnd, 0&, iLeft, iTop, iRight, iBottom, 0&  ' SWP_FRAMECHANGED 'SWP_NOSIZE Or SWP_NOMOVE

    
End Sub

Public Sub SetActiveWindow(ByVal sFormName As String)
    JOEWTabs.SetForm sFormName
End Sub

Public Sub RemoveChildWindow(ByVal sFormName As String)
    JOEWTabs.RemoveForm sFormName
End Sub

Public Sub ResizeClientWin(ByRef WinhWnd As Long, _
                            ByVal iTop As Integer, _
                            ByVal iLeft As Integer, _
                            ByVal iRight As Integer, _
                            ByVal iBottom As Integer)

    SetWindowPos WinhWnd, 0&, iLeft, iTop, iRight, iBottom, 0&  ' SWP_FRAMECHANGED 'SWP_NOSIZE Or SWP_NOMOVE

End Sub












Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get SBWidth() As Integer
    SBWidth = m_SBWidth
End Property

Public Property Let SBWidth(ByVal New_SBWidth As Integer)
    
    If New_SBWidth < 0 Then
        New_SBWidth = 0
    End If
    
    If New_SBWidth > GetWidth Then
        New_SBWidth = GetWidth
    End If
    
    m_SBWidth = New_SBWidth
    
    PropertyChanged "SBWidth"
    
    Call UserControl_Resize
End Property

Private Sub JOEWTabs_Click(sFormName As String, Index As Integer)
    RaiseEvent FormTabClick(sFormName, Index)
End Sub

Private Sub JOEWTabs_CloseClick(sFormName As String, Index As Integer)
    RaiseEvent CloseClick(sFormName, Index)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_SBWidth = m_def_SBWidth
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_SBWidth = PropBag.ReadProperty("SBWidth", m_def_SBWidth)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    imgBR.Move (m_SBWidth - imgBR.Width) + 2, 0, imgBR.Width, GetHeight
    JOEWTabs.Move (m_SBWidth + 2), 0, GetWidth - (m_SBWidth + 2)
    Err.Clear
End Sub

Private Sub UserControl_Show()
    Call UserControl_Resize
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("SBWidth", m_SBWidth, m_def_SBWidth)
End Sub

