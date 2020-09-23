VERSION 5.00
Begin VB.UserControl SkinnableMeter 
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4395
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   2190
   ScaleWidth      =   4395
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Left            =   240
      Max             =   27
      TabIndex        =   2
      Top             =   0
      Value           =   27
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picDEST 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   210
      TabIndex        =   1
      Top             =   0
      Width           =   210
   End
   Begin VB.PictureBox picSRC0 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   480
      Picture         =   "SkinnableMeter.ctx":0000
      ScaleHeight     =   1935
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
   End
End
Attribute VB_Name = "SkinnableMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Coding credit for bitblt function code goes to "DosAscii"

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020     'Copies the source bitmap to destination bitmap.
Private Const SRCAND = &H8800C6      'Combines pixels of the destination with source bitmap
                                    'using the Boolean AND operator.
Private Const SRCINVERT = &H660046   'Combines pixels of the destination with source bitmap
                                    'using the Boolean XOR operator.
Private Const SRCPAINT = &HEE0086    'Combines pixels of the destination with source bitmap
                                    'using the Boolean OR operator.
Private Const SRCERASE = &H4400328   'Inverts the destination bitmap and then combines the
                                    'results with the source bitmap using the Boolean AND
                                    'operator.
Private Const WHITENESS = &HFF0062   'Turns all output white.
Private Const BLACKNESS = &H42       'Turn output black.
 'This foreces all varibles to be declared now.
Dim i As Integer
Public XPos 'X-axis
Public YPos 'Y-axis
Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
    Private Const SW_NORMAL = 1


Public Sub OpenWebsite(strWebsite As String)


    If ShellExecute(&O0, "Open", strWebsite, vbNullString, vbNullString, SW_NORMAL) < 33 Then
        ' Insert Error handling code here
    End If
End Sub
Private Sub UserControl_Initialize()
    
    Call BitBlt(picDEST.hDC, 0, 0, 13, 63, picSRC0.hDC, 0, 1, SRCAND)
    picDEST.BorderStyle = 1

End Sub

Private Sub VScroll1_Change()
    
    SwapImages

End Sub
Function SwapImages()
    picDEST.Picture = Nothing
    i = VScroll1.Value
    Dim iY As Integer
        
    If i < 14 Then
        iY = 1
        
    Else
        iY = 66
        i = i - 14
    End If
    
'    Label1.Caption = i
    Call BitBlt(picDEST.hDC, 0, 0, 13, 63, picSRC0.hDC, i * 15, iY, SRCAND)
    picDEST.BorderStyle = 1
End Function
Private Sub VScroll1_Scroll()
    
    SwapImages
    
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picSRC0,picSRC0,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = picSRC0.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set picSRC0.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=VScroll1,VScroll1,-1,Value
Public Property Get Value() As Integer
Attribute Value.VB_Description = "Returns/sets the value of an object."
    Value = VScroll1.Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    VScroll1.Value() = New_Value
    PropertyChanged "Value"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    VScroll1.Value = PropBag.ReadProperty("Value", 27)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Value", VScroll1.Value, 27)
End Sub

