VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Galaxy Messenger Login"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "&Remember username and pswrd"
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   2880
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Sign in"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
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
      Height          =   1035
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   4380
      Begin VB.Image picLogo 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   120
         Picture         =   "frmLogin.frx":08CA
         Stretch         =   -1  'True
         Top             =   240
         Width           =   660
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   2040
         Picture         =   "frmLogin.frx":2340
         Stretch         =   -1  'True
         Top             =   50
         Width           =   2175
      End
      Begin VB.Image picLogo2 
         Height          =   495
         Left            =   840
         Picture         =   "frmLogin.frx":8102
         Top             =   240
         Width           =   3105
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Please sign in with your .NET passport to see your online contacts,  have online converstations and receive alerts."
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Username:"
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
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SaveUN As Boolean


Private Sub Check1_Click()
If Check1.Value = 0 Then SaveSetting "Galaxy Messenger", "Login", "Remember", "No": SaveUN = False
If Check1.Value = 1 Then SaveSetting "Galaxy Messenger", "Login", "Remember", "Yes": SaveUN = True

End Sub

Private Sub Command1_Click()
If SaveUN = False Then
SaveSetting "Galaxy Messenger", "Login", "Password", ""
SaveSetting "Galaxy Messenger", "Login", "Username", ""
End If


If Text1 = "" Then Text1.SetFocus: Exit Sub
If Text2 = "" Then Text2.SetFocus: Exit Sub
Form1.StartConnect Text1, Text2
Unload Form2


End Sub

Private Sub Command2_Click()
Unload Form2
End Sub

Private Sub Form_Load()
Text1 = GetSetting("Galaxy Messenger", "Login", "Username")
Text2 = GetSetting("Galaxy Messenger", "Login", "Password")
Text1.SelStart = Len(Text1)
Text2.SelStart = Len(Text2)

If GetSetting("Galaxy Messenger", "Login", "Remember", "Yes") = "Yes" Then
SaveUN = True
Check1.Value = 1
Else
SaveUN = False
Check1.Value = 0
End If

End Sub

Private Sub Text1_Change()
If SaveUN = False Then Exit Sub
SaveSetting "Galaxy Messenger", "Login", "Username", Text1
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0: Text1.SelLength = Len(Text1)

End Sub

Private Sub Text2_Change()
If SaveUN = False Then Exit Sub
SaveSetting "Galaxy Messenger", "Login", "Password", Text2
End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0: Text2.SelLength = Len(Text2)
End Sub

