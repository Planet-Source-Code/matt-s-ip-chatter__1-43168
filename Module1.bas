Attribute VB_Name = "Module1"
Option Explicit
Public IIP As String
Public Enum SMTP_State
    MAIL_CONNECT
    MAIL_HELO
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
End Enum

Public m_State As SMTP_State


