VERSION 5.00
Begin VB.Form HelpFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   Icon            =   "Clockhelp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   330
      Left            =   2280
      TabIndex        =   0
      Top             =   2985
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Visit Our Web Site"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   60
      MouseIcon       =   "Clockhelp.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2580
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created by Michael Margold"
      Height          =   240
      Left            =   60
      TabIndex        =   1
      Top             =   2250
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   2100
      Left            =   0
      Picture         =   "Clockhelp.frx":0E1C
      Top             =   0
      Width           =   5865
   End
End
Attribute VB_Name = "HelpFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Function Link(ByVal URL As String) As Long
  Link = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function
Private Sub Label1_Click()
  Call Link("http://www.soft-collection.com")
End Sub
Private Sub Command1_Click()
  Me.Hide
End Sub
