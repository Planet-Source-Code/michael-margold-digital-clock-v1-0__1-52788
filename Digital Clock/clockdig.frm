VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   555
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4665
   Icon            =   "clockdig.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "clockdig.frx":0CCA
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   2190
      Top             =   1830
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6300
      Top             =   840
   End
   Begin VB.Image Image4 
      Height          =   465
      Left            =   1860
      ToolTipText     =   "Hide for 10 seconds"
      Top             =   45
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   13
      Left            =   4350
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   12
      Left            =   4080
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   11
      Left            =   3810
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   10
      Left            =   3540
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   9
      Left            =   3135
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   8
      Left            =   2865
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   7
      Left            =   2460
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   6
      Left            =   2190
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   5
      Left            =   1620
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   4
      Left            =   1350
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   3
      Left            =   990
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   2
      Left            =   720
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   1
      Left            =   360
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   0
      Left            =   90
      Top             =   90
      Width           =   225
   End
   Begin VB.Image Image6 
      Height          =   90
      Left            =   4050
      ToolTipText     =   "Exit"
      Top             =   -15
      Width           =   630
   End
   Begin VB.Image Image5 
      Height          =   90
      Left            =   30
      ToolTipText     =   "Help"
      Top             =   -15
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   0
      Left            =   4245
      Picture         =   "clockdig.frx":9454
      Top             =   4170
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   4575
      Picture         =   "clockdig.frx":9946
      Top             =   4170
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   4935
      Picture         =   "clockdig.frx":9E38
      Top             =   4170
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   3
      Left            =   5295
      Picture         =   "clockdig.frx":A32A
      Top             =   4170
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   4
      Left            =   5655
      Picture         =   "clockdig.frx":A81C
      Top             =   4170
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   5
      Left            =   6015
      Picture         =   "clockdig.frx":AD0E
      Top             =   4170
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   6
      Left            =   6375
      Picture         =   "clockdig.frx":B200
      Top             =   4170
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   7
      Left            =   6735
      Picture         =   "clockdig.frx":B6F2
      Top             =   4170
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   8
      Left            =   7095
      Picture         =   "clockdig.frx":BBE4
      Top             =   4170
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   9
      Left            =   7455
      Picture         =   "clockdig.frx":C0D6
      Top             =   4170
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image3 
      Height          =   90
      Left            =   690
      ToolTipText     =   "Move"
      Top             =   -15
      Width           =   3330
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lastx As Integer
Dim Lasty As Integer
Dim Saveleft As Variant
Dim Savetop As Variant
Dim DelayCounter As Long
Private Sub Form_Load()
  Saveleft = Val(GetSetting("Digital Clock", "Settings", "Left", 0))
  Savetop = Val(GetSetting("Digital Clock", "Settings", "Top", 0))
  MainFrm.Left = Saveleft
  MainFrm.Top = Savetop
End Sub
Private Sub image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Lastx = X
  Lasty = Y
End Sub
Private Sub image3_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    MainFrm.Left = MainFrm.Left + (X - Lastx)
    MainFrm.Top = MainFrm.Top + (Y - Lasty)
  End If
End Sub
Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call SaveSetting("Digital Clock", "Settings", "Left", Str(MainFrm.Left))
  Call SaveSetting("Digital Clock", "Settings", "Top", Str(MainFrm.Top))
End Sub
Private Sub Image4_Click()
  MainFrm.Hide
  DelayCounter = 0
  Timer2.Interval = 100
End Sub
Private Sub Image5_Click()
  HelpFrm.Show
End Sub
Private Sub Image6_Click()
  End
End Sub
Private Sub Timer1_Timer()
  Image2(0) = Image1(Mid$(Time$, 1, 1))
  Image2(1) = Image1(Mid$(Time$, 2, 1))
  Image2(2) = Image1(Mid$(Time$, 4, 1))
  Image2(3) = Image1(Mid$(Time$, 5, 1))
  Image2(4) = Image1(Mid$(Time$, 7, 1))
  Image2(5) = Image1(Mid$(Time$, 8, 1))
  Image2(6) = Image1(Mid$(Date$, 1, 1))
  Image2(7) = Image1(Mid$(Date$, 2, 1))
  Image2(8) = Image1(Mid$(Date$, 4, 1))
  Image2(9) = Image1(Mid$(Date$, 5, 1))
  Image2(10) = Image1(Mid$(Date$, 7, 1))
  Image2(11) = Image1(Mid$(Date$, 8, 1))
  Image2(12) = Image1(Mid$(Date$, 9, 1))
  Image2(13) = Image1(Mid$(Date$, 10, 1))
End Sub
Private Sub Timer2_Timer()
  DelayCounter = DelayCounter + 1
  If DelayCounter >= 100 Then
     MainFrm.Show
     Timer2.Interval = 0
  End If
End Sub
