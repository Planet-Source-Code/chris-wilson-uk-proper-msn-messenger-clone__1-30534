VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "username@username.com - Instant Message"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3600
      Width           =   4095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4920
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4455
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Last Message: None"
            TextSave        =   "Last Message: None"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6218
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5530
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMessage.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ThisUser As String
Dim IncConnect As Integer
Dim ThisCKIHash As String
Dim ThisUserName As String
Dim ThisSessionID As String
Dim MessageNumber As Long
Dim WaitingToSend As String
Dim Connected As Boolean



Private Sub Command1_Click()
Dim mess As String
Dim mess2 As String

Text1.Text = Text1.Text & Form1.UserName & ":" & vbCrLf & Text2 & vbCrLf & vbCrLf

mess = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=Arial; EF=B; CO=0; CS=0; PF=22" & vbCrLf & vbCrLf & Text2
'THE GREEN = fa200
mess2 = "MSG " & MessageNumber & " N " & Len(mess) & vbCrLf & mess
Text2 = ""
SendData mess2, True

End Sub

Private Sub Form_Load()
Text1 = "Never give out your password or credit card number in an instant message conversation." & vbCrLf & " ---- " & vbCrLf & vbCrLf
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
End Sub



Public Sub OutConnection(RemoteUser As String, SwitchboardIP As String, CKIHash As String)
ThisUser = RemoteUser
ThisCKIHash = CKIHash
Winsock1.Close: Winsock1.Connect SwitchboardIP, 1863
Me.Caption = RemoteUser & " - Instant Message"



End Sub


Public Sub IncConnection(SessionID As String, SwitchboardIP As String, CKIHash As String, UserEmail As String, UserName As String, LocalEmail As String)
ThisSessionID = SessionID
ThisCKIHash = CKIHash
ThisUser = UserEmail
ThisUserName = UserName
ThisUserName = RemoveString(ThisUserName, "%20", " ")
ThisUserName = RemoveString(ThisUserName, "%25", "%")

IncConnect = 1
Me.Caption = UserEmail & " - Instant Message"
Winsock1.Close: Winsock1.Connect SwitchboardIP, 1863
End Sub

Private Sub Winsock1_Connect()
If IncConnect = 1 Then
'ANS 1 venky_dude@hotmail.com 989495494.750408580 11742066
SendData "ANS 1 " & Form1.UserName & " " & ThisCKIHash & " " & ThisSessionID, False
Connected = True
End If

If IncConnect = 0 Then
SendData "USR 1 " & Form1.UserName & " " & ThisCKIHash, False
End If

End Sub

Private Sub SendData(Data As String, NoVbCrLf As Boolean)
If NoVbCrLf = True Then Winsock1.SendData Data Else Winsock1.SendData Data & vbCrLf
Form1.StartAni
Debug.Print "SWITCHBOARD DATA - OUTGOING: " & Data

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim IncData As String
Winsock1.GetData IncData
Form1.StartAni
Debug.Print "SWITCHBOARD DATA - INCOMMING: " & IncData


'DISCONNECTED
If Mid(IncData, 1, 3) = "NAK" Then
Command1.Enabled = False
StatusBar1.Panels(2).Text = "You are not connected to the switchboard"
End If


'INCOMMING MESSAGE
If InStr(1, IncData, "text/plain;") Then
Me.Show
Me.Text2.SetFocus
Command1.Enabled = True


Dim TheX As Integer
Dim CrS As Integer
Do
TheX = TheX + 1
If Mid(IncData, TheX, 2) = vbCrLf Then CrS = CrS + 1
If CrS = 4 Then GoTo 10
Loop


10
If ThisUserName = "" Then
ThisUserName = GetItem(IncData, 2)
ThisUserName = RemoveString(ThisUserName, "%20", " ")
ThisUserName = RemoveString(ThisUserName, "%25", "%")
End If

Text1.Text = Text1.Text & ThisUserName & ":" & vbCrLf & Mid(IncData, TheX + 4) & vbCrLf & vbCrLf


Text1.SelStart = Len(Text1.Text)
StatusBar1.Panels.Item(2).Text = ""
StatusBar1.Panels.Item(1).Text = "Last Message: " & Time$

'StatusBar1.SimpleText = "Last message received at " & Time$
End If


'USER IT TYPING A MESSAGE
If InStr(1, IncData, "TypingUser: " & ThisUser) Then
StatusBar1.Panels.Item(2).Text = ThisUser & " is typing a message"
End If


'USER ACCEPTED TO CONNECT TO SWITCHBOARD
If Mid(IncData, 1, 8) = "USR 1 OK" Then SendData "CAL 2 " & ThisUser, False: StatusBar1.Panels(2).Text = "Connecting with " & ThisUser

'RINGING USER
If Mid(IncData, 1, 13) = "CAL 2 RINGING" Then StatusBar1.Panels(2).Text = "Waiting for response.."

'USER HAS JOINED THE CONVERSATION
If Mid(IncData, 1, 3) = "JOI" Then
StatusBar1.Panels(2).Text = ThisUser & " is in the conversation"
Command1.Enabled = True
Connected = True
Me.Show
Me.Text2.SetFocus

End If

'USER HAS LEFT THE CONVERSATION
If Mid(IncData, 1, 3) = "BYE" Then StatusBar1.Panels(2).Text = ThisUser & " has left the conversation": Connected = False: Command1.Enabled = False








End Sub
