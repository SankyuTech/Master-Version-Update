VERSION 5.00
Begin VB.UserControl JOELine 
   BackColor       =   &H00E0E0E0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7140
   ScaleHeight     =   90
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   Begin VB.Line Line2 
      BorderColor     =   &H00F6F8F8&
      X1              =   2
      X2              =   248
      Y1              =   2
      Y2              =   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00ACD0D7&
      X1              =   0
      X2              =   246
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "JOELine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Sub UserControl_InitProperties()
     MsgBox "JoeLine" & vbNewLine & "Code By: Joehel V. Del Rosario"
End Sub

Private Sub UserControl_Resize()

    UserControl.Height = Screen.TwipsPerPixelY * 2
    
    Line1.X1 = 0
    Line1.X2 = UserControl.Width / Screen.TwipsPerPixelX
    Line1.Y1 = 0
    Line1.Y2 = 0
    
    Line2.X1 = 0
    Line2.X2 = UserControl.Width / Screen.TwipsPerPixelX
    Line2.Y1 = 1
    Line2.Y2 = 1
    
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line1,Line1,-1,BorderColor
Public Property Get BorderColor1() As OLE_COLOR
Attribute BorderColor1.VB_Description = "Returns/sets the color of an object's border."
    BorderColor1 = Line1.BorderColor
End Property

Public Property Let BorderColor1(ByVal New_BorderColor1 As OLE_COLOR)
    Line1.BorderColor() = New_BorderColor1
    PropertyChanged "BorderColor1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line2,Line2,-1,BorderColor
Public Property Get BorderColor2() As OLE_COLOR
Attribute BorderColor2.VB_Description = "Returns/sets the color of an object's border."
    BorderColor2 = Line2.BorderColor
End Property

Public Property Let BorderColor2(ByVal New_BorderColor2 As OLE_COLOR)
    Line2.BorderColor() = New_BorderColor2
    PropertyChanged "BorderColor2"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Line1.BorderColor = PropBag.ReadProperty("BorderColor1", &HF6F8F8)
    Line2.BorderColor = PropBag.ReadProperty("BorderColor2", &H80000010)
    Line1.BorderStyle = PropBag.ReadProperty("BorderStyle1", 1)
    Line2.BorderStyle = PropBag.ReadProperty("BorderStyle2", 1)
End Sub


'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderColor1", Line1.BorderColor, &HF6F8F8)
    Call PropBag.WriteProperty("BorderColor2", Line2.BorderColor, &H80000010)
    Call PropBag.WriteProperty("BorderStyle1", Line1.BorderStyle, 1)
    Call PropBag.WriteProperty("BorderStyle2", Line2.BorderStyle, 1)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line1,Line1,-1,BorderStyle
Public Property Get BorderStyle1() As BorderStyleConstants
Attribute BorderStyle1.VB_Description = "Returns/sets the border style for an object."
    BorderStyle1 = Line1.BorderStyle
End Property

Public Property Let BorderStyle1(ByVal New_BorderStyle1 As BorderStyleConstants)
    Line1.BorderStyle() = New_BorderStyle1
    PropertyChanged "BorderStyle1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line2,Line2,-1,BorderStyle
Public Property Get BorderStyle2() As BorderStyleConstants
Attribute BorderStyle2.VB_Description = "Returns/sets the border style for an object."
    BorderStyle2 = Line2.BorderStyle
    
End Property

Public Property Let BorderStyle2(ByVal New_BorderStyle2 As BorderStyleConstants)
    Line2.BorderStyle() = New_BorderStyle2
    PropertyChanged "BorderStyle2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

