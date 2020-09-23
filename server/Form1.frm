VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "File Transter Server"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog fs 
      Left            =   3840
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wrec 
      Left            =   2760
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox fNames 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6615
      TabIndex        =   3
      Top             =   2310
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "Accept"
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   5760
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock wLogin 
      Left            =   4560
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5760
      Top             =   1200
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5760
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock wServer 
      Left            =   1200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   885
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Remote:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cIP, fName, accok As Boolean

Private Sub Command1_Click()
Winsock1.SendData "Accept"
wrec.LocalPort = 1215
wrec.Bind
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
wServer.Close
wServer.Bind 7777, wServer.LocalIP
Command1.Enabled = False
End Sub

Private Sub winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Winsock1.GetData gotstring, vbString
End Sub

Private Sub wLogin_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
wLogin.GetData gotstring, vbString
End Sub

Private Sub wrec_DataArrival(ByVal bytesTotal As Long)
wrec.GetData newbuf, 1
filecontents = newbuf
fs.CancelError = True
fs.ShowSave
Open fs.FileName For Binary Access Write As #1
Put #1, , filecontents
Close #1
MsgBox "File successfully copied", vbOKOnly + vbInformation, "Server"
End Sub

Private Sub wServer_DataArrival(ByVal bytesTotal As Long)
wServer.GetData gotstring, vbString

Select Case gotstring
Case "Ready?"
Winsock1.SendData "OK"
End Select
partsting = Left(gotstring, 3)
Select Case partsting
Case "IP="
txtServer.Text = cIP
cIP = Mid(gotstring, 4, Len(gotstring) - 3)
Winsock1.Close
Winsock1.RemoteHost = cIP
Winsock1.LocalPort = 11111
Winsock1.RemotePort = 7778
Winsock1.Bind Winsock1.LocalPort
Winsock1.SendData "Yes"
End Select
partsting = Left(gotstring, 5)
Select Case partsting
Case "File="
fName = Mid(gotstring, 6, Len(gotstring) - 5)
fNames.Text = fName
End Select
partstring = Left(gotstring, 9)
Select Case partstring
Case "Password="
password = Mid(gotstring, 10, Len(gotstring) - 9)
If Text2.Text = password Then
Winsock1.SendData "OK All"
accok = True
If Len(Trim(fNames)) > 0 Then
Command1.Enabled = True
End If
Else
Winsock1.SendData "Reject"
End If

End Select
End Sub


