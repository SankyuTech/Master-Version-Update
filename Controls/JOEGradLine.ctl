VERSION 5.00
Begin VB.UserControl JOEGradLine 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ClipBehavior    =   0  'None
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   72
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   210
      ScaleHeight     =   175
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   283
      TabIndex        =   0
      Top             =   30
      Width           =   4245
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1260
         TabIndex        =   1
         Top             =   90
         Width           =   45
      End
   End
End
Attribute VB_Name = "JOEGradLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_Color1 = &H0&
Const m_def_Color2 = &HFFFFFF
Const m_def_Angle = 1
'Property Variables:
Dim m_Color1 As OLE_COLOR
Dim m_Color2 As OLE_COLOR
Dim m_Angle As Integer

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Color1() As OLE_COLOR
    Color1 = m_Color1
End Property

Public Property Let Color1(ByVal New_Color1 As OLE_COLOR)
    m_Color1 = New_Color1
    PropertyChanged "Color1"
    RedrawGrad
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Color2() As OLE_COLOR
    Color2 = m_Color2
End Property

Public Property Let Color2(ByVal New_Color2 As OLE_COLOR)
    m_Color2 = New_Color2
    PropertyChanged "Color2"
    RedrawGrad
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Angle() As Integer
    Angle = m_Angle
End Property

Public Property Let Angle(ByVal New_Angle As Integer)
    m_Angle = New_Angle
    PropertyChanged "Angle"
    RedrawGrad
End Property



Private Sub RedrawGrad()
    Dim cGrad As New clsGrad
    'On Error Resume Next
    
    cGrad.Color1 = CLng(m_Color1)
    cGrad.Color2 = CLng(m_Color2)
    cGrad.Angle = m_Angle
    cGrad.Draw bgMain
    bgMain.Refresh
    
    Set cGrad = Nothing
    
    Err.Clear
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Color1 = m_def_Color1
    m_Color2 = m_def_Color2
    m_Angle = m_def_Angle
End Sub



'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Color1 = PropBag.ReadProperty("Color1", m_def_Color1)
    m_Color2 = PropBag.ReadProperty("Color2", m_def_Color2)
    m_Angle = PropBag.ReadProperty("Angle", m_def_Angle)
    lblCaption.Alignment = PropBag.ReadProperty("Alignment", 0)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "")
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
End Sub

Private Sub UserControl_Resize()
    If GetHeight < 2 Then
        UserControl.Height = 30
    End If
    bgMain.Move 0, -1, GetWidth, GetHeight
    lblCaption.Move 3, (bgMain.Height - lblCaption.Height) / 2
    RedrawGrad
End Sub



Private Sub UserControl_Show()
    RedrawGrad
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Color1", m_Color1, m_def_Color1)
    Call PropBag.WriteProperty("Color2", m_Color2, m_def_Color2)
    Call PropBag.WriteProperty("Angle", m_Angle, m_def_Angle)
    Call PropBag.WriteProperty("Alignment", lblCaption.Alignment, 0)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "")
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H80000012)
End Sub

Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function
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
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
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
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property


