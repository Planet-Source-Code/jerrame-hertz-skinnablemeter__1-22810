VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1275
   ClientLeft      =   10485
   ClientTop       =   4185
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Left            =   1560
      Max             =   27
      TabIndex        =   4
      Top             =   120
      Value           =   27
      Width           =   135
   End
   Begin Project1.SkinnableMeter SkinnableMeter1 
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1720
      Picture         =   "Form1.frx":0000
      Value           =   0
   End
   Begin Project1.SkinnableMeter SkinnableMeter1 
      Height          =   975
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1720
      Picture         =   "Form1.frx":13CC6
      Value           =   0
   End
   Begin Project1.SkinnableMeter SkinnableMeter1 
      Height          =   975
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1720
      Picture         =   "Form1.frx":2798C
      Value           =   0
   End
   Begin Project1.SkinnableMeter SkinnableMeter1 
      Height          =   975
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1720
      Picture         =   "Form1.frx":3B652
      Value           =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "This Code Is 'OK'"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "I Like This Code?"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()
    SkinnableMeter1(0).OpenWebsite ("http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=22810&optCodeRatingValue=5")
End Sub

Private Sub Label2_Click()
    SkinnableMeter1(0).OpenWebsite ("http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=22810&optCodeRatingValue=4")
End Sub

Private Sub VScroll1_Change()

    ScrollMeter

End Sub

Private Sub VScroll1_Scroll()
    
    ScrollMeter

End Sub

Function ScrollMeter()
    
    Dim i As Integer
    For i = 0 To SkinnableMeter1.Count - 1
    
        SkinnableMeter1(i).Value = VScroll1.Max - VScroll1.Value
    
    Next
    
End Function
