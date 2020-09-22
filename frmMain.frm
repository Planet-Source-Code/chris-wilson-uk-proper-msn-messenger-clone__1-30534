VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galaxy Messenger"
   ClientHeight    =   5595
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4395
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   975
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   4440
      Width           =   4335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   3000
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   120
      Width           =   4335
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8520
      Top             =   5040
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8520
      Top             =   4560
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   4335
      TabIndex        =   5
      Top             =   4560
      Width           =   4395
      Begin VB.Image picLogo 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   120
         Picture         =   "frmMain.frx":2358
         Stretch         =   -1  'True
         Top             =   140
         Width           =   780
      End
      Begin VB.Image picLogo2 
         Height          =   495
         Left            =   1000
         Picture         =   "frmMain.frx":3DCE
         Top             =   190
         Width           =   3105
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4150
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   3600
         Top             =   3000
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "To sign in with a different name click here"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   4
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   120
         Picture         =   "frmMain.frx":8E80
         Stretch         =   -1  'True
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "username@username.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "You are not signed in"
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
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   0
         Top             =   120
         Width           =   4215
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   7646
      _Version        =   393217
      Indentation     =   353
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
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
   Begin VB.Label Label4 
      Caption         =   "ABOVE: USERLIST    - BELOW: ONLINE CONTACTS"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Image cAni3 
      Height          =   480
      Left            =   6240
      Picture         =   "frmMain.frx":974A
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image cAni2 
      Height          =   480
      Left            =   6240
      Picture         =   "frmMain.frx":A014
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image cAni1 
      Height          =   480
      Left            =   6240
      Picture         =   "frmMain.frx":A8DE
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Ani3 
      Height          =   855
      Left            =   6960
      Picture         =   "frmMain.frx":B1A8
      Top             =   1920
      Width           =   810
   End
   Begin VB.Image Ani2 
      Height          =   615
      Left            =   6960
      Picture         =   "frmMain.frx":D66E
      Top             =   1320
      Width           =   690
   End
   Begin VB.Image Ani1 
      Height          =   690
      Left            =   6960
      Picture         =   "frmMain.frx":ED1C
      Top             =   720
      Width           =   615
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Signout 
         Caption         =   "&Sign out"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "&Options"
      Begin VB.Menu Prefs 
         Caption         =   "&Prefrences"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Animation As Integer
Dim Animation2 As Integer
Dim AnimationTimer As Integer
Dim LastUsername As String
Dim LastPassword As String
Public UserName As String
Dim Password As String
Dim ConnectSection As Integer
Dim RecList As Boolean
Dim TotContacts As Integer
Dim AtLoad As Boolean
Dim iM(65535) As New Form4
Dim IMnumber As Integer
Dim RingUser As String
Dim SwitchRequest As Boolean



Public Sub StartConnect(UN As String, PW As String)
LastUsername = UN
LastPassword = PW
UserName = UN
Password = PW
Label3.Visible = False
Label2 = "Click here to cancel sign in"
Label1 = "Attempting to connect.."
Timer2.Enabled = True
ConnectSection = 0
Winsock1.Connect "64.4.13.58", 1863

End Sub

Private Sub About_Click()
Form3.Show
End Sub

Private Sub Exit_Click()
RemoveTrayIcon

End

End Sub

Private Sub Form_Load()
If AtLoad = False Then SysTray.AddTrayIcon Form1.Icon, Form1, "Galaxy Messenger": AtLoad = True
SwitchRequest = False
Text1 = ""
Text2 = ""
Text3 = ""

RecList = False
LastUsername = GetSetting("Galaxy Messenger", "Login", "Username")
LastPassword = GetSetting("Galaxy Messenger", "Login", "Password")
If LastUsername = "" Then Label2 = "Click here to sign in"
If Not LastUsername = "" Then Label2 = LastUsername
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If TrayEvent(X) = "LEFTUP" Then Form1.Show
If TrayEvent(X) = "RIGHTUP" Then
response$ = MsgBox("Are you sure you want to close Galaxy Messenger?" & vbCrLf & "All online conversations will be closed", vbQuestion + vbYesNo, "Galaxy Messenger")
If response$ = vbYes Then SysTray.RemoveTrayIcon: End
End If



End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Form1.Hide
End Sub

Private Sub Label2_Click()
If Label2 = "Click here to sign in" Then Label3_Click: Exit Sub
If Label2 = "Click here to cancel sign in" Then
Form_Load
Winsock1.Close
Label1 = "You are not signed in"
Timer2.Enabled = False
Image1.Picture = cAni1.Picture
Label3.Visible = True
Exit Sub

End If
StartConnect LastUsername, LastPassword
End Sub

'ANIMATION SUBS
Private Sub Label3_Click()
Form2.Show
End Sub

Private Sub Prefs_Click()
MsgBox "Sorry, there are no settings available yet", vbInformation, "Galaxy Messenger"

End Sub

Private Sub Signout_Click()
Form_Load
Winsock1.Close
Label1 = "You are not signed in"
Timer2.Enabled = False
Image1.Picture = cAni1.Picture
Label3.Visible = True
TreeView1.Nodes.Clear
Frame1.Visible = True
End Sub

Private Sub Timer1_Timer()
AnimationTimer = AnimationTimer + 1
Animation = Animation + 1
If Animation = 0 Then picLogo.Picture = Ani1.Picture
If Animation = 1 Then picLogo.Picture = Ani2.Picture
If Animation = 2 Then picLogo.Picture = Ani3.Picture
If Animation = 2 Then Animation = -1
If AnimationTimer = 6 Then Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Animation2 = Animation2 + 1
If Animation2 = 0 Then Image1.Picture = cAni1.Picture
If Animation2 = 1 Then Image1.Picture = cAni2.Picture
If Animation2 = 2 Then Image1.Picture = cAni3.Picture
If Animation2 = 2 Then Animation2 = -1
End Sub

Public Sub StartAni()
AnimationTimer = 0
Timer1.Enabled = True
End Sub

Private Sub TreeView1_DblClick()
If TreeView1.SelectedItem.Index = 1 Or TreeView1.SelectedItem.Index = 2 Then Exit Sub

RingUser = TreeView1.SelectedItem.Text
IMnumber = IMnumber + 1
Load iM(IMnumber)
iM(IMnumber).ThisUser = TreeView1.SelectedItem.Text
SwitchRequest = True
SendData "XFR " & IMnumber & " SB"
End Sub

Private Sub Winsock1_Close()
Form_Load
Winsock1.Close
Label1 = "You are not signed in"
Timer2.Enabled = False
Image1.Picture = cAni1.Picture
Label3.Visible = True
End Sub

'END ANIMATION SUBS
Function RemoveString(Entire As String, Word As String, Replace As String) As String
    Dim I As Integer
    I = 1
    Dim LeftPart
    Do While True
        I = InStr(1, Entire, Word)
        If I = 0 Then
            Exit Do
        Else
            LeftPart = Left(Entire, I - 1)
            Entire = LeftPart & Replace & Right(Entire, Len(Entire) - Len(Word) - Len(LeftPart))
        End If
    Loop
    
   RemoveString = Entire
      
End Function

Private Sub Winsock1_Connect()
If ConnectSection = 0 Then SendData "VER 0 MSNP5 MSNP4 CVRO": Label1 = "Preparing to log in.."
If ConnectSection = 1 Then SendData "VER 3 MSNP5 MSNP4 CVRO": Label1 = "Logging on to server.."
End Sub

Public Sub SendData(Data As String, Optional NoCRLF As Boolean = False)
If NoCRLF = False Then Winsock1.SendData Data & vbCrLf Else Winsock1.SendData Data
Debug.Print ">> Outgoing Data: " & vbCrLf & Data
StartAni
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim IncData As String
Winsock1.GetData IncData

Debug.Print "<< Incomming Data: " & vbCrLf & IncData
StartAni

'STAGE 1
If IncData = "VER 0 MSNP5 MSNP4" & vbCrLf Then SendData "INF 1": Label1 = Label1 & "."

'STAGE 2
If IncData = "INF 1 MD5" & vbCrLf Then SendData "USR 2 MD5 I " & UserName: Label1 = Label1 & "."

'STAGE 3 (receives ip address of server and connects to it)
If Mid(IncData, 1, 8) = "XFR 2 NS" Then
X = InStrRev(IncData, ":")
ConnectSection = 1
Winsock1.Close
Winsock1.Connect Mid(IncData, 10, X - 10), 1863
End If

'STAGE 4
If IncData = "VER 3 MSNP5 MSNP4" & vbCrLf Then SendData "INF  4"

'STAGE 5
If IncData = "INF 4 MD5" & vbCrLf Then SendData "USR 5 MD5 I " & UserName

'STAGE 6
If Mid(IncData, 1, 11) = "USR 5 MD5 S" Then
Label1 = "Verifying username && password.."
mt = InStrRev(IncData, "S")
xt = Right(IncData, Len(IncData) - mt)
xt = Left(xt, Len(xt) - 2)
xt = Right(xt, Len(xt) - 1)
SendData "USR 6 MD5 S " & MD5.MD5String(xt & Password)
End If
If IncData = "911 6" & vbCrLf Then
MsgBox "Your username or password is invalid", vbInformation, "Error"
Winsock1_Close
End If

'STAGE 7
If Mid(IncData, 1, 8) = "USR 6 OK" Then
X = InStrRev(IncData, " ")
uname$ = Mid(IncData, X + 1, Len(Mid(IncData, X)) - 2) & ", signing you in.."
uname$ = RemoveString(uname$, "%20", " ")
uname$ = RemoveString(uname$, "%25", "%")
Label1 = "Welcome " & uname$
SendData "CHG 7 NLN" 'sets online status and receives list of online contacts
End If

'STAGE 8 (decoding the online contacts list)
'NTS: in future versions make this detect users status
If InStr(1, IncData, "ILN 7") Then
Text2 = Text2 & IncData
End If
If Mid(IncData, 1, 9) = "CHG 7 NLN" Then SendData "LST 8 RL" 'request contact list

'Useful notes: (* = deffinatly..  NST: Find out the others)
'BUSY = BSY *
'AWAY = AWY *
'BRB = BRB *
'PHONE = PHN
'IDLE = IDL *
'LUNCH = OUT
'ONLINE = NLN *
'OFFLINE = FLN *

'\/\/\/\/\/\/\/\/\ STAGE 9 - FINAL STAGE (decoding the contact list) \/\/\/\/\/\/\/\/\
'\/\/\/\/\/\/\/\/\ STAGE 9 - FINAL STAGE (decoding the contact list) \/\/\/\/\/\/\/\/\
If RecList = True Then
Text1 = Text1 & IncData


If GetContactNumber(Text1, GetTotalContacts(Text1)) = GetTotalContacts(Text1) Then
Frame1.Visible = False
RecList = False

TreeView1.Nodes.Add , , , "Online", 3
TreeView1.Nodes.Add , , , "Offline", 2
TreeView1.Nodes.Item(1).Bold = True
TreeView1.Nodes.Item(2).Bold = True
TreeView1.Nodes.Item(1).Expanded = True
TreeView1.Nodes.Item(2).Expanded = True

Do
lists1 = lists1 + 1
If Mid(Text1, lists1, 2) = vbCrLf Then

list2 = Mid(Text1, 1, lists1 - 1)
Text1 = Mid(Text1, lists1 + 2)
lists1 = 0


Do
list3 = list3 + 1
If Mid(list2, list3, 1) = " " Then spaces2 = spaces2 + 1
Loop Until spaces2 = 6

list3 = list3 + 1
list2 = Mid(list2, list3)

Do
list4 = list4 + 1
If Mid(list2, list4, 1) = " " Then spaces3 = spaces3 + 1
Loop Until spaces3 = 1

list4 = list4 - 1
list2 = Mid(list2, 1, list4)

'ADD TO LIST
'NTS: list2 = email
If Text2 = "" Then TreeView1.Nodes.Add 2, tvwChild, , list2, 1: GoTo 55
If InStr(1, Text2.Text, list2) Then
TreeView1.Nodes.Add 1, tvwChild, , list2, 1
Else
TreeView1.Nodes.Add 2, tvwChild, , list2, 1
End If
55 ' Sleep 10



End If
56
list3 = 0
list2 = ""
list4 = 0

spaces2 = 0
Spaces = 0
spaces3 = 0

If Len(Text1) <= 8 Then GoTo 60
Loop Until Text1 = ""
60



End If
Exit Sub
End If
'END STAGE 9

If InStr(1, IncData, "LST 8 RL") Then
Label1 = "Receiving contact list.."
Dim List1 As String
RecList = True
'List1 = Mid(incData, InStr(1, incData, "LST 8 RL"))
List1 = IncData
Text1 = Text1 & List1
xx% = 0

Do
I% = I% + 1
If Mid(Text1, I%, 1) = " " Then xx% = xx% + 1
Loop Until xx% = 5
I% = I% + 1
Do
o% = o% + 1
If Mid(Text1, I% + o%, 1) = " " Then Z% = Z% + 1
Loop Until Z% = 1

TotContacts = Mid(Text1, I%, o%)
End If

'IF CONTACT COMES ONLINE
If Mid(IncData, 1, 7) = "NLN NLN" Then
nowonline = GetItem(IncData, 2)

listsearchx = 2
Do
listsearchx = listsearchx + 1
If listsearchx = TreeView1.Nodes.Count + 1 Then Exit Sub
Loop Until InStr(1, TreeView1.Nodes.Item(listsearchx).Text, nowonline)

TreeView1.Nodes.Remove (listsearchx)
TreeView1.Nodes.Add 1, tvwChild, , nowonline, 1

End If

'IF CONTACT GOES OFFLINE
If Mid(IncData, 1, 3) = "FLN" Then
nowoffline = GetItem(IncData, 1, , True)

listsearchx = 2

Do
listsearchx = listsearchx + 1
If listsearchx = TreeView1.Nodes.Count + 1 Then Exit Sub
Loop Until InStr(1, TreeView1.Nodes.Item(listsearchx).Text, nowoffline)

TreeView1.Nodes.Remove (listsearchx)
TreeView1.Nodes.Add 2, tvwChild, , nowoffline, 1
End If

'IF CONTACT GOES AWAY
If Mid(IncData, 1, 7) = "NLN AWY" Then
nowaway = GetItem(IncData, 2)

listsearchx = 2

Do
listsearchx = listsearchx + 1
If listsearchx = TreeView1.Nodes.Count + 1 Then Exit Sub
Loop Until InStr(1, TreeView1.Nodes.Item(listsearchx).Text, nowaway)

TreeView1.Nodes.Remove (listsearchx)
TreeView1.Nodes.Add 1, tvwChild, , nowaway & " (Away)", 1
End If

'IF CONTACT GOES BRB
If Mid(IncData, 1, 7) = "NLN BRB" Then
nowBRB = GetItem(IncData, 2)

listsearchx = 2

Do
listsearchx = listsearchx + 1
If listsearchx = TreeView1.Nodes.Count + 1 Then Exit Sub
Loop Until InStr(1, TreeView1.Nodes.Item(listsearchx).Text, nowBRB)

TreeView1.Nodes.Remove (listsearchx)
TreeView1.Nodes.Add 1, tvwChild, , nowBRB & " (Be Right Back)", 1
End If

'IF CONTACT GOES BUSY
If Mid(IncData, 1, 7) = "NLN BSY" Then
nowBusy = GetItem(IncData, 2)

listsearchx = 2

Do
listsearchx = listsearchx + 1
If listsearchx = TreeView1.Nodes.Count + 1 Then Exit Sub
Loop Until InStr(1, TreeView1.Nodes.Item(listsearchx).Text, nowBusy)

TreeView1.Nodes.Remove (listsearchx)
TreeView1.Nodes.Add 1, tvwChild, , nowBusy & " (Busy)", 1
End If

'IF CONTACT GOES ON THE PHONE
If Mid(IncData, 1, 7) = "NLN PHN" Then
nowPhone = GetItem(IncData, 2)

listsearchx = 2

Do
listsearchx = listsearchx + 1
If listsearchx = TreeView1.Nodes.Count + 1 Then Exit Sub
Loop Until InStr(1, TreeView1.Nodes.Item(listsearchx).Text, nowPhone)

TreeView1.Nodes.Remove (listsearchx)
TreeView1.Nodes.Add 1, tvwChild, , nowPhone & " (On the Phone)", 1
End If

'IF CONTACT GOES To Lunch
If Mid(IncData, 1, 7) = "NLN LUN" Then
nowLunch = GetItem(IncData, 2)

listsearchx = 2

Do
listsearchx = listsearchx + 1
If listsearchx = TreeView1.Nodes.Count + 1 Then Exit Sub
Loop Until InStr(1, TreeView1.Nodes.Item(listsearchx).Text, nowLunch)

TreeView1.Nodes.Remove (listsearchx)
TreeView1.Nodes.Add 1, tvwChild, , nowLunch & " (Out to Lunch)", 1
End If

'IF CONTACT GOES IDLE
If Mid(IncData, 1, 7) = "NLN IDL" Then
nowIdle = GetItem(IncData, 2)

listsearchx = 2

Do
listsearchx = listsearchx + 1
If listsearchx = TreeView1.Nodes.Count + 1 Then Exit Sub
Loop Until InStr(1, TreeView1.Nodes.Item(listsearchx).Text, nowIdle)

TreeView1.Nodes.Remove (listsearchx)
TreeView1.Nodes.Add 1, tvwChild, , nowIdle & " (Idle)", 1
End If


'WHEN USER REQUESTS TO START A CONVO
If Mid(IncData, 1, 3) = "RNG" Then
IMnumber = IMnumber + 1

Load iM(IMnumber)
iM(IMnumber).IncConnection GetItem(IncData, 1), GetItem(IncData, 2, ":"), GetItem(IncData, 4), GetItem(IncData, 5), GetItem(IncData, 6, , True), UserName
End If

'WHEN LOCAL USER REQUEST TO START A CONVO IS ACCEPTED
If Mid(IncData, 1, 3) = "XFR" Then
If SwitchRequest = False Then Exit Sub
SwitchRequest = False
temp$ = GetItem(IncData, 1)
iM(temp$).OutConnection iM(temp$).ThisUser, GetItem(IncData, 3, ":"), GetItem(IncData, 5, , True)
End If



End Sub



