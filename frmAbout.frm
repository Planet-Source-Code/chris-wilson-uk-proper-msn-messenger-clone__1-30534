VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   4260
      Begin VB.Image picLogo2 
         Height          =   495
         Left            =   1005
         Picture         =   "frmAbout.frx":08CA
         Top             =   165
         Width           =   3105
      End
      Begin VB.Image picLogo 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   120
         Picture         =   "frmAbout.frx":597C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   780
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "For comments, help or suggestions email: chris@wilsonr1.karoo.co.uk"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Version 2.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Graphics and Coding by Chris Wilson 2002"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "MSN Messenger Clone - Visual Basic"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Galaxy Messenger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form3

End Sub

Private Sub Form_Load()
Label4 = Label4 & " - Build: " & App.Revision
End Sub
