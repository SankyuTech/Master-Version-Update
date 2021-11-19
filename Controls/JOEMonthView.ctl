VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl JOEMonthView 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   146
   Begin VB.TextBox txtYear 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   570
   End
   Begin MSComCtl2.UpDown updYear 
      Height          =   315
      Left            =   570
      TabIndex        =   1
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1900
      BuddyControl    =   "txtYear"
      BuddyDispid     =   196609
      OrigLeft        =   2190
      OrigTop         =   1410
      OrigRight       =   2445
      OrigBottom      =   1695
      Max             =   2200
      Min             =   1900
      Enabled         =   -1  'True
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "JOEMonthView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long



'Default Property Values:
Const m_def_MonthToString = ""
'Property Variables:
'Event Declarations:
Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."



Private Sub cmbMonth_Click()
    If Me.Enabled = True Then
        RaiseEvent Change
    End If
End Sub

Private Sub cmbMonth_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        MonthToString = cmbMonth.Text
    End If
End Sub

Private Sub cmbMonth_LostFocus()
    MonthToString = cmbMonth.Text
End Sub



Private Sub cmbMonth_Validate(Cancel As Boolean)
    If cmbMonth.ListIndex < 0 Then
        Cancel = True
    End If
End Sub

Private Sub txtYear_Change()
    

    If IsNumeric(txtYear.Text) = False Then
        Exit Sub
    End If
    If Val(txtYear.Text) < updYear.Min Or _
        Val(txtYear.Text) > updYear.Max Then
        Exit Sub
    End If

    updYear.Value = Val(txtYear.Text)
    If Me.Enabled = True Then
        RaiseEvent Change
    End If
End Sub

Private Sub txtYear_Validate(Cancel As Boolean)
    Dim sYear As String
    
    sYear = Trim(txtYear.Text)
    
    If IsNumeric(txtYear.Text) = False Then
        txtYear.Text = Year(Now)
        Exit Sub
    End If
    If Val(txtYear.Text) < updYear.Min Or _
        Val(txtYear.Text) > updYear.Max Then
        txtYear.Text = Year(Now)
    End If
    
    updYear.Value = Val(txtYear.Text)
    
    If sYear <> Trim(txtYear.Text) Then
        If Me.Enabled = True Then
            RaiseEvent Change
        End If
    End If
End Sub

Private Sub updYear_Change()
    txtYear.Text = updYear.Value
End Sub

Private Sub UserControl_Initialize()
    InitCommonControls

    cmbMonth.Clear
    cmbMonth.AddItem "January"
    cmbMonth.AddItem "February"
    cmbMonth.AddItem "March"
    cmbMonth.AddItem "April"
    cmbMonth.AddItem "May"
    cmbMonth.AddItem "June"
    cmbMonth.AddItem "July"
    cmbMonth.AddItem "August"
    cmbMonth.AddItem "September"
    cmbMonth.AddItem "October"
    cmbMonth.AddItem "November"
    cmbMonth.AddItem "December"
    
    cmbMonth.ListIndex = Month(Now) - 1
    updYear.Value = Year(Now)
    txtYear.Text = Year(Now)
End Sub
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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMonth,cmbMonth,-1,ListIndex
Public Property Get MonthVal() As Integer
Attribute MonthVal.VB_Description = "Returns/sets the index of the currently selected item in the control."
    MonthVal = cmbMonth.ListIndex + 1
End Property

Public Property Let MonthVal(ByVal New_MonthVal As Integer)
    cmbMonth.ListIndex() = New_MonthVal
    PropertyChanged "MonthVal"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=updYear,updYear,-1,Value
Public Property Get YearVal() As Long
Attribute YearVal.VB_Description = "Get/Set the current position in the scroll range"
    YearVal = updYear.Value
End Property

Public Property Let YearVal(ByVal New_YearVal As Long)
    If New_YearVal < updYear.Min Or _
        New_YearVal > updYear.Max Then
        New_YearVal = Year(Now)
    End If
    

    updYear.Value() = New_YearVal
    PropertyChanged "YearVal"
    If Me.Enabled = True Then
        RaiseEvent Change
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get MonthToString() As String
    MonthToString = cmbMonth.Text
End Property

Public Property Let MonthToString(ByVal New_MonthToString As String)
    Dim oM As String
    Dim i As Integer
    
    If LCase(Trim(New_MonthToString)) = LCase(Trim(cmbMonth.Text)) Then
        Exit Property
    End If
    
    oM = cmbMonth.Text
    
    For i = 0 To 11
        If LCase(Trim(cmbMonth.List(i))) = LCase(Trim(New_MonthToString)) Then
            cmbMonth.ListIndex = i
            Exit For
        End If
    Next
    
    PropertyChanged "MonthToString"
    If Me.Enabled = True Then
        RaiseEvent Change
    End If
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 0)
    cmbMonth.ListIndex = PropBag.ReadProperty("MonthVal", 0)
    updYear.Value = PropBag.ReadProperty("YearVal", 1900)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    End Sub


Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function
Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function


Private Sub UserControl_Resize()
    On Error Resume Next
    
    cmbMonth.Width = GetWidth - cmbMonth.Left
    cmbMonth.SelStart = 0
    cmbMonth.SelLength = 0
End Sub

Private Sub UserControl_Show()
    cmbMonth.ListIndex = Month(Now) - 1
    updYear.Value = Year(Now)
    cmbMonth.SelStart = 0
    cmbMonth.SelLength = 0
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 0)
    Call PropBag.WriteProperty("MonthVal", cmbMonth.ListIndex, 0)
    Call PropBag.WriteProperty("YearVal", updYear.Value, 1900)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    
    txtYear.Enabled = New_Enabled
    updYear.Enabled = New_Enabled
    cmbMonth.Enabled = New_Enabled
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

