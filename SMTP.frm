VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form SMTP 
   Caption         =   "SMTP"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   Icon            =   "SMTP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMessage 
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3240
      Width           =   7095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4680
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtHost 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtSender 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtRecipient 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtSubject 
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Type Message Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   12
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Subject"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Recipient"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Sender"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Server"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "SMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_State
Dim strData As String


Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
txtRecipient.text = ""
txtSubject.text = ""
txtMessage.text = ""
Unload Me
SMTP.Show
End Sub



Private Sub cmdSend_Click()
Winsock1.Connect Trim$(txtHost), 25
m_State = MAIL_CONNECT

End Sub

Private Sub Form_Activate()
txtSender.SetFocus
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim strServerResponse As String
Dim strResponseCode As String

strResponseCode = Left(strServerResponse, 3)

If strData = "250" Or strData = "220" Or _
strData = "354" Then

Select Case m_State
Case MAIL_CONNECT
Case MAIL_HELO
Case MAIL_FROM
Case MAIL_RCPTTO
Case MAIL_DATA
Case MAIL_DOT
Case MAIL_QUIT
End Select

End If

Dim strDataToSend As String

Winsock1.GetData strServerResponse

Debug.Print strServerResponse

strResponseCode = Left(strServerResponse, 3)

If strResponseCode = "250" Or _
strResponseCode = "220" Or _
strResponseCode = "354" Then
       
Select Case m_State
Case MAIL_CONNECT
m_State = MAIL_HELO
strDataToSend = Trim$(txtSender)

 strDataToSend = Left$(strDataToSend, _
                                InStr(1, strDataToSend, _
                                "@") - 1)
Winsock1.SendData "HELO " & _
strDataToSend & vbCrLf

Debug.Print "HELO " & strDataToSend

Case MAIL_HELO

m_State = MAIL_FROM

Winsock1.SendData "MAIL FROM:" & _
                                  Trim$(txtSender) & vbCrLf
Debug.Print "MAIL FROM:" & Trim$(txtSender)

Case MAIL_FROM

m_State = MAIL_RCPTTO

Winsock1.SendData "RCPT TO:" & _
                                  Trim$(txtRecipient) & vbCrLf

Debug.Print "RCPT TO:" & Trim$(txtRecipient)

Case MAIL_RCPTTO

m_State = MAIL_DATA

Winsock1.SendData "DATA" & vbCrLf

Debug.Print "DATA"

Case MAIL_DATA

m_State = MAIL_DOT

Winsock1.SendData "Subject:" & txtSubject & vbLf

Debug.Print "Subject:" & txtSubject
Dim varLines As Variant
Dim varLine As Variant

varLines = Split(txtMessage, vbCrLf)
For Each varLine In varLines
Winsock1.SendData CStr(varLine) & vbLf

Debug.Print CStr(varLine)
Next

Winsock1.SendData "." & vbCrLf
Debug.Print "."
Case MAIL_DOT
m_State = MAIL_QUIT
Winsock1.SendData "QUIT" & vbCrLf
Debug.Print "QUIT"
Case MAIL_QUIT

Winsock1.Close

End Select

End If
    
End Sub















