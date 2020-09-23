VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmServer 
   Caption         =   "TCP Server"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtfOutput 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2355
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmServer.frx":0000
   End
   Begin MSWinsockLib.Winsock tcpServer 
      Index           =   0
      Left            =   3360
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Num

Private Sub Form_Load()
Num = 0 'num of winsocks
tcpServer(0).LocalPort = 27478 'port to listen on
tcpServer(0).Listen 'start listening (for clients)
End Sub

Private Sub tcpServer_ConnectionRequest(Index As Integer, ByVal requestID As Long) 'client requests to connect
Num = Num + 1 'increase number that hold # of winsocks
Load tcpServer(Num) 'load new winsock for new client
tcpServer(Num).Accept requestID 'accept request by getting ID
tcpServer(Num).SendData "(Connected)" 'send "(Connected)" to client
End Sub

Private Sub tcpServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String 'Dim variable for data
tcpServer(Index).GetData strData 'load incoming data
rtfOutput.Text = rtfOutput.Text & vbCrLf & "To Server (" & tcpServer(Index).RemoteHostIP & ") :) 'add incoming data to richtextbox with the ip it's from with it"
End Sub

Private Sub rtfOutput_Change()
rtfOutput.SelLength = Len(rtfOutput.Text) 'keeps text showing the most recent text
End Sub
