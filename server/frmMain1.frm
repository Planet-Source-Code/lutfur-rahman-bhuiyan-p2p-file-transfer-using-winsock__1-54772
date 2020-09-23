VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "File Transter (Client)"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wcom 
      Left            =   2880
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock wlogin 
      Left            =   1800
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wserver 
      Left            =   1080
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Server"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CDLG 
      Left            =   0
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox txtSend 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Text            =   "File to send"
      Top             =   0
      Width           =   3615
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   600
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server IP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   870
   End
   Begin VB.Label lblStatus 
      Caption         =   "Label1"
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const RPORT As Integer = 5555
Dim Buf() As Byte
Dim bufPos As Long
Dim Sendbytes As Long

Sub SendFile(strFile As String)
ReDim Buf(FileLen(strFile) - 1)

Open strFile For Binary As #1
  Get #1, 1, Buf
Close #1

With Winsock1
  .RemoteHost = txtRemote.Text
  .RemotePort = RPORT
  .Connect
End With

lblStatus.Caption = "Trying to connect..."
End Sub

Private Sub Check1_Click()
If Check1.Value = 0 Then
  Winsock1.Close
  Close #1
Else
  Winsock1.LocalPort = RPORT
  Winsock1.Listen
  Open App.Path & "\downloaded" For Binary As #1
  bufPos = 1
End If
End Sub

Private Sub cmdBrowse_Click()
CDLG.Filter = "All files (*.*)|*.*"
CDLG.ShowOpen
txtSend.Text = CDLG.FileName

lblStatus.Caption = CDLG.FileName & " is ready for take-off!"
End Sub

Private Sub cmdSend_Click()
SendFile (txtSend.Text)
End Sub

Private Sub Command1_Click()
wcom.Close
wcom.RemoteHost = txtServer.Text
wcom.LocalPort = 11111
wcom.RemotePort = 7777
wcom.Bind wcom.LocalPort
wcom.SendData "Hello"
End Sub

Private Sub Form_Load()
wserver.Close
wserver.Bind 7778, wserver.LocalIP
wlogin.Close
wlogin.Bind 7779, wserver.LocalIP
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #1
Winsock1.Close
Unload Me
End Sub

Private Sub Winsock1_Close()
Close #1
Sendbytes = 0
Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
'change label-caption and send data when connected
lblStatus.Caption = "Connected to " & txtRemote.Text
Winsock1.SendData Buf
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'let's hope only one connects, or else the 1st user will be disconnected
Winsock1.Close
Winsock1.Accept requestID 'accept one connection
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim newBuf() As Byte

Winsock1.GetData newBuf 'recieve data
Put #1, bufPos, newBuf 'put newBuf (the data received) into #1 at bufPos
bufPos = bufPos + UBound(newBuf) + 1 'get the right position
End Sub

Private Sub Winsock1_SendComplete()
lblStatus.Caption = "Done."
Winsock1.Close
Buf() = ""
End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Sendbytes = Sendbytes + bytesSent
'UBound -> Returns a long containing the largest available
'          subscript for the indicated dimension of an array
lblStatus.Caption = Int(((Sendbytes / UBound(Buf)) * 100)) & " %  completed..."
End Sub
