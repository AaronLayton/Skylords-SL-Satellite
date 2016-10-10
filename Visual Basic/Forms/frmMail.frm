VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Email"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   Icon            =   "frmMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   4080
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   4080
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3120
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim inData As String
Private change As Boolean
Private Const TIME_OUT = 30


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    mTest = False
    cSending = False
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    mTest = False
    cSending = False
    Me.Hide
End Sub

Private Sub Timer1_Timer()
    timer = timer + 1
    If timer = TIME_OUT Then
        Winsock1.Close
        MsgBox "Could not connect to host " + Winsock1.RemoteHost + vbCrLf + ", Operation timed out"
        Timer1.Enabled = False
        Timer2.Enabled = True
        mTest = False
        cSending = False
    End If
End Sub

Private Sub Timer2_Timer()
    hideIt = hideIt + 1
    If hideIt = 6 Then
        Unload Me
        Timer2.Enabled = False
    End If
    mTest = False
    cSending = False
End Sub

Private Sub Winsock1_Connect()
    Text1 = Text1 & "Connected" & vbNewLine
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    
    Dim reply As String
    Dim tmp() As String
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 220 Then           'Error occured
        MsgBox "Server returned the following error:" + vbCrLf + reply
        Winsock1.Close
        Exit Sub
    End If
    'Start the process
    Winsock1.SendData "HELO " + Winsock1.LocalHostName + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 250 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        Winsock1.Close
        Exit Sub
    End If
    'Send MAIL FROM
    Winsock1.SendData "MAIL FROM:<Skylords@SL-Satellite.co.uk>" + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 250 Then
        MsgBox "Server returned the following error:" + vbCrLf + inData
        Winsock1.Close
        Exit Sub
    End If
    'Send RCPT TO
    Winsock1.SendData "RCPT TO:<" + frmOptions.Text4 + ">" + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 250 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        Winsock1.Close
        Exit Sub
    End If
    'Send DATA
    DoEvents
    Winsock1.SendData "DATA" + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable         'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 354 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        Winsock1.Close
        Exit Sub
    End If
    'Send the E-Mail
    Winsock1.SendData "From: <Skylords@SL-Satellite.co.uk>" + vbCrLf + _
                      "To: " + frmOptions.Text4 + vbCrLf + _
                      "Subject: " + Subject + vbCrLf + _
                      "X-Mailer: anyMail v1.1" + vbCrLf + _
                      "Mime-Version: 1.0" + vbCrLf + _
                      "Content-Type: text/plain;" + vbTab + "charset=us-ascii" + vbCrLf + vbCrLf + _
                      Message
    Winsock1.SendData vbCrLf + "." + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not DataAvailable             'Wait for reply
        If Winsock1.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    DataAvailable = False
    tmp = Split(reply, " ")
    If Not Val(tmp(0)) = 250 Then               'Error occured
        MsgBox "Server returned the following error:" + vbCrLf + reply
        Winsock1.Close
        Exit Sub
    End If
    Winsock1.SendData "QUIT"
    Text1 = Text1 & "Message Sent!" & vbNewLine
    Winsock1.Close
    mTest = False
    cSending = False
    Timer2.Enabled = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim data As String
    Winsock1.GetData data, vbString
    'Add data arrived data to the already arrived data
    inData = inData + data
    'Wait till a line is recieved (with CR LF in the end)
    If StrComp(Right$(inData, 2), vbCrLf) = 0 Then DataAvailable = True
    
    Text1 = Text1 & inData & vbNewLine
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Not Number = sckSuccess Then
        MsgBox Description          'Display error
        Timer1.Enabled = False
        Winsock1.Close
        mTest = False
        cSending = False
        Timer2.Enabled = ture
    End If
End Sub
