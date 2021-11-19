VERSION 5.00
Begin VB.UserControl JOEWinTabs 
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   ScaleHeight     =   52
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   750
   Begin MOVERS.JoeWinTab JoeWinTab 
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1860
      _ExtentX        =   3493
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12735512
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   1024
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   -165
      Picture         =   "JOEWinTabs.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18135
   End
End
Attribute VB_Name = "JOEWinTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



'JOEWinTabs
'by: Vincent J. Jamero
'
'Created: 11:45 pm  June 11, 2006
'Modified: 2:10 am June 11,2006


Option Explicit

Private Type tForm
    Name As String
    Caption As String
    CloseButton As Boolean
End Type

Dim MyForms(25) As tForm
Dim FUBound As Integer

Public Event Click(sFormName As String, Index As Integer)
Public Event Change(sFormName As String, Index As Integer)
Public Event CloseClick(sFormName As String, Index As Integer)
'Default Property Values:
Const m_def_ForeColor = vbBlack
Const m_def_MaxButtonWidth = 157
Const m_def_CurIndex = -1
'Property Variables:
Dim m_ForeColor As OLE_COLOR
Dim m_Font As Font
Dim m_MaxButtonWidth As Integer
Dim m_CurIndex As Integer



Public Function UpdateButtons()
    Dim i As Integer
    Dim iCount As Integer
    Dim lWidth As Long
    Dim iCurIndex As Integer
    Dim NewWidth As Integer

    'Remove Blank
    RemoveBlank
    
    iCount = GetUBound
    
    On Error Resume Next
    If iCount < 0 Then
        For i = 0 To JoeWinTab.UBound
            JoeWinTab(i).Visible = False
            Unload JoeWinTab(i)
        Next
        Exit Function
    End If
    
    lWidth = GetWidth / (iCount + 1)
    NewWidth = IIf(lWidth > m_MaxButtonWidth, m_MaxButtonWidth, lWidth)
    
    iCurIndex = CurIndex
    
    For i = 0 To iCount
        Load JoeWinTab(i)
        
        JoeWinTab(i).Caption = MyForms(i).Caption
        JoeWinTab(i).Move i * NewWidth, 0, NewWidth - 2
        JoeWinTab(i).CloseButton = MyForms(i).CloseButton
        
        If i = iCurIndex Then
            If Len(Trim(JoeWinTab(i).Caption)) > 0 Then
                JoeWinTab(i).Value = True
                JoeWinTab(i).Font.Bold = True
            End If
        Else
            JoeWinTab(i).Value = False
            JoeWinTab(i).Font.Bold = False
        End If
        
        JoeWinTab(i).Visible = True
    Next
    
    
    If JoeWinTab.UBound > iCount Then
    
        For i = iCount + 1 To JoeWinTab.UBound
            JoeWinTab(i).Visible = False
        Next
    End If
    
    
    
    
End Function

Public Function SetForm(sName As String)
    
    Dim i As Integer
    Dim oldIndex As Integer
    
    oldIndex = m_CurIndex
    
    If GetUBound >= 0 Then
        For i = 0 To GetUBound
            If MyForms(i).Name = sName Then
                CurIndex = i
                Exit For
            End If
        Next
    End If
    
    UpdateButtons
    
    If oldIndex <> m_CurIndex Then
        RaiseEvent Change(MyForms(m_CurIndex).Name, m_CurIndex)
    End If
    
End Function

Public Function ClsButtons()
    Dim i As Integer
    
    For i = 0 To 25
        MyForms(i).Caption = ""
        MyForms(i).Name = ""
    Next
    
    On Error Resume Next
    For i = 0 To JoeWinTab.UBound
        JoeWinTab(i).Visible = False
        If i <> 0 Then
            Unload JoeWinTab(i)
        End If
    Next
End Function

Public Function AddForm(sName As String, sCaption As String, Optional ShowCloseButton As Boolean = True)
    Dim i As Integer
    Dim newIndex As Integer
    Dim UB As Integer
    
    UB = GetUBound
    
    If UB >= 25 Then
        Exit Function
    End If
    
    If UB >= 0 Then
        For i = 0 To 25
            If LCase(Trim(MyForms(i).Name)) = LCase(Trim(sName)) Then
                Exit Function
            End If
        Next
    End If

    
    newIndex = UB + 1
    MyForms(newIndex).Name = sName
    MyForms(newIndex).Caption = sCaption
    MyForms(newIndex).CloseButton = ShowCloseButton

    UpdateButtons
End Function

Private Sub RemoveBlank()
    
    Dim i As Integer
    Dim ti As Integer
    Dim tmpForm(25) As tForm
    
    ti = 0
    For i = 0 To 25
        If Len(Trim(MyForms(i).Name)) > 0 Then
            tmpForm(ti).Name = MyForms(i).Name
            tmpForm(ti).Caption = MyForms(i).Caption
            ti = ti + 1
        End If
        MyForms(i).Caption = ""
        MyForms(i).Name = ""
    Next
    
    For i = 0 To ti
        MyForms(i).Caption = tmpForm(i).Caption
        MyForms(i).Name = tmpForm(i).Name
    Next
    
End Sub

Public Function RemoveForm(sName As String)
    Dim i As Integer
    Dim X As Integer
    Dim si As Integer
   
    X = GetUBound
    si = -1
    
    If X >= 0 Then
        For i = 0 To X
            If LCase(MyForms(i).Name) = LCase(sName) Then
                MyForms(i).Name = ""
                MyForms(i).Caption = ""
                si = i
                Exit For
            End If
        Next
    End If

    UpdateButtons
End Function
Public Function GetUBound() As Integer
    Dim i As Integer
    Dim iGetUBound As Integer

    iGetUBound = -1
    For i = 0 To 25
        If Len(Trim(MyForms(i).Name)) > 0 Then
            iGetUBound = iGetUBound + 1
        End If
    Next

    GetUBound = iGetUBound
End Function

Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function

Private Sub JoeWinTab_Click(Index As Integer)
    RaiseEvent Click(MyForms(Index).Name, Index)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,-1
Public Property Get CurIndex() As Integer
    CurIndex = m_CurIndex
End Property

Public Property Let CurIndex(ByVal New_CurIndex As Integer)
    
    Dim iCount As Integer
    Dim i As Integer
    iCount = GetUBound
    If iCount < 0 Then
        Exit Property
    End If
    
    If New_CurIndex < 0 Or New_CurIndex > iCount Then
        Exit Property
    End If
    
    m_CurIndex = New_CurIndex
    PropertyChanged "CurIndex"
    
    For i = 0 To iCount
        If i = New_CurIndex Then
            JoeWinTab(i).Value = True
            JoeWinTab(i).Font.Bold = True
        Else
            JoeWinTab(i).Value = False
            JoeWinTab(i).Font.Bold = False
        End If
    Next
    
End Property



Private Sub JoeWinTab_CloseClick(Index As Integer)
    RaiseEvent CloseClick(MyForms(Index).Name, Index)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_CurIndex = m_def_CurIndex
    m_ForeColor = m_def_ForeColor
    Set m_Font = Ambient.Font
    m_MaxButtonWidth = m_def_MaxButtonWidth
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_CurIndex = PropBag.ReadProperty("CurIndex", m_def_CurIndex)
    Set JoeWinTab(0).Font = PropBag.ReadProperty("Font", Ambient.Font)
    JoeWinTab(0).ForeColor = PropBag.ReadProperty("ForeColor", &H808080)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_MaxButtonWidth = PropBag.ReadProperty("MaxButtonWidth", m_def_MaxButtonWidth)
End Sub

Private Sub UserControl_Resize()
    
    Dim lWidth As Integer
    Dim iCount As Integer
    Dim i As Integer
    iCount = GetUBound
    
    On Error Resume Next
    
    If iCount >= 0 Then
        lWidth = GetWidth / (iCount + 1)
        lWidth = IIf(lWidth > 157, 157, lWidth)
        For i = 0 To iCount
            JoeWinTab(i).Move i * lWidth, 0, lWidth - 2
        Next
    End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("CurIndex", m_CurIndex, m_def_CurIndex)
    Call PropBag.WriteProperty("Font", JoeWinTab(0).Font, "Tahoma")
    Call PropBag.WriteProperty("ForeColor", JoeWinTab(0).ForeColor, &H808080)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("MaxButtonWidth", m_MaxButtonWidth, m_def_MaxButtonWidth)
End Sub
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=JoeWinTab(0),JoeWinTab,0,Font
'Public Property Get Font() As Font
'    Set Font = JoeWinTab(0).Font
'End Property
'
'Public Property Set Font(ByVal New_Font As Font)
'    Dim i As Integer
'
'    For i = 0 To JoeWinTab.UBound
'        Set JoeWinTab(i).Font = New_Font
'    Next
'
'    PropertyChanged "Font"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=JoeWinTab(0),JoeWinTab,0,ForeColor
'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = JoeWinTab(0).ForeColor
'End Property
''
'Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
'    Dim i As Integer
'
'    For i = 0 To JoeWinTab.UBound
'        JoeWinTab(i).ForeColor() = New_ForeColor
'    Next
'    PropertyChanged "ForeColor"
'End Property
'
'
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,
Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,157
Public Property Get MaxButtonWidth() As Integer
    MaxButtonWidth = m_MaxButtonWidth
End Property

Public Property Let MaxButtonWidth(ByVal New_MaxButtonWidth As Integer)
    m_MaxButtonWidth = New_MaxButtonWidth
    PropertyChanged "MaxButtonWidth"
End Property


Public Property Get FormCaption(Index As Integer) As String
    FormCaption = MyForms(Index).Caption
End Property
