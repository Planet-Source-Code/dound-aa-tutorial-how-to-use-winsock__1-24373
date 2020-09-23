VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmClient 
   Caption         =   "TCP Client"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtfOutput 
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2355
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmClient.frx":0000
   End
   Begin VB.CommandButton cmdSendData 
      Caption         =   "Send"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtSendData 
      Height          =   405
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   2160
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error GoTo errorguy
tcpClient.RemoteHost = tcpClient.LocalIP  'IP address here...localip=ip you're using so that you can test client/server on your own comp =D ... I recommend http://www.dyndns.org/ if you have a changing IP and need clients to connect to server
tcpClient.RemotePort = 27478 'assign port server will listen for client on
tcpClient.Connect 'try to connect to server

Exit Sub

errorguy: MsgBox Err.Description & "... Now exiting program.": Unload frmClient: End 'tell user the error and exit program
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long) 'incoming data
Dim strData As String 'dim variable for data collection
tcpClient.GetData strData 'get data

If strData = "(Connected)" Then 'server sends when it connects
rtfOutput.Text = rtfOutput.Text & vbCrLf & "Connected" 'tell user that connection was successful
txtSendData.Locked = False 'allow user to type messages to send
End If

rtfOutput.Text = rtfOutput.Text & vbCrLf & "To You: " & strData 'add incoming data to richtextbox
End Sub

Private Sub rtfOutput_Change()
rtfOutput.SelLength = Len(rtfOutput.Text) 'move focus to last typed message
End Sub

Private Sub cmdSendData_Click()
On Error GoTo errorguy
If Not txtSendData.Text = "" Then 'if not data to send = nothing
tcpClient.SendData txtSendData.Text 'send data
rtfOutput.Text = rtfOutput.Text & vbCrLf & "From You: " & txtSendData.Text  'record data sent into richtextbox
txtSendData.Text = "" 'clear textbox where message to send is typed
txtSendData.SetFocus 'set focus to send data textbox so user can immediately start typing again
End If

Exit Sub

errorguy: MsgBox Err.Description & "... Now exiting program.": Unload frmClient: End 'tell user the error and exit program
End Sub
