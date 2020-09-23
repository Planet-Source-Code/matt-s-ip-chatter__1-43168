VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form ipchat 
   Caption         =   "IP CHAT"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   Icon            =   "ipchat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmail 
      Caption         =   "Send E-mail"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "File Transfer"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5280
      Width           =   7695
   End
   Begin MSWinsockLib.Winsock CTLWinsock 
      Left            =   7800
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsChat 
      Left            =   7680
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   7200
      TabIndex        =   7
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtOut 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   6975
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtIn 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00400040&
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   960
      Width           =   6975
   End
   Begin VB.Label Label2 
      Caption         =   "IP"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nick"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "ipchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdclose_Click()
wsChat.Close
CTLWinsock.Close
cmdclose.Enabled = False
cmdSend.Enabled = False

txtName.Enabled = True
cmdlisten.Enabled = True
cmdConnect.Enabled = True

txtIn.text = "----- Connection Closed -----" & vbCrLf
End Sub

Private Sub cmdConnect_Click()
IIP = Val(txtIP.text)
If txtName.text = " " Or txtIP.text = "" Then
MsgBox "Error Fill In Boxes", vbCritical, "ERROR"

Else
wsChat.Close
wsChat.Connect txtIP.text, 9999
cmdclose.Enabled = True
cmdlisten.Enabled = False
cmdConnect.Enabled = False
txtName.Enabled = False
End If
Do
DoEvents
 Loop Until wsChat.State = sckConnected Or wsChat.State = sckError
 
 If wsChat.State = sckConnected Then
 
AddText "----- Connection Established -----" & vbCrLf, txtIn

cmdSend.Enabled = True
txtName.Enabled = False
txtOut.SetFocus

Else


AddText "----- Connection Failed -----" & vbCrLf, txtIn

End If
 
 End Sub

Private Sub AddText(ByVal text As String, ByRef Box As TextBox)
Box.text = Box.text & text & vbCrLf
Box.SelStart = Len(Box.text)
End Sub







Private Sub cmdlisten_Click()
If txtName.text = "" Then

MsgBox "You must enter an name first!", vbCritical, "Error!"

txtName.SetFocus
Exit Sub

End If


txtIP.text = wsChat.LocalIP
txtIP.text = IIP
wsChat.Close
wsChat.LocalPort = 9999
wsChat.Listen

cmdclose.Enabled = True
cmdlisten.Enabled = False
cmdConnect.Enabled = False

txtName.Enabled = False

AddText "----- Waiting for Connection -----", txtIn
End Sub

Private Sub cmdmail_Click()
SMTP.Show
End Sub

Private Sub cmdSend_Click()
If UCase(txtOut.text) = "/QUIT" Then
wsChat.Close
ElseIf UCase(txtOut.text) = "/CLEAR" Then
txtIn.text = ""
Else
wsChat.SendData "[" & txtName.text & "] " & txtOut.text

AddText "[" & txtName.text & "] " & txtOut.text, txtIn

txtOut.text = ""
txtOut.SetFocus
End If
End Sub







Private Sub Command1_Click()
Form1.Show
End Sub

   
Private Sub wschat_ConnectionRequest(ByVal requestID As Long)
wsChat.Close
wsChat.Accept requestID


AddText "----- Connection Established -----" & vbCrLf, txtIn


cmdSend.Enabled = True
txtName.Enabled = False
txtOut.SetFocus
End Sub

Private Sub wschat_DataArrival(ByVal bytesTotal As Long)

Dim incoming As String


wsChat.GetData incoming
AddText incoming, txtIn
End Sub


Private Sub wschat_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If number <> 0 Then

AddText "----- Error [" & Description & "] -----" & vbCrLf, txtIn

Call cmdclose_Click

End If

End Sub
