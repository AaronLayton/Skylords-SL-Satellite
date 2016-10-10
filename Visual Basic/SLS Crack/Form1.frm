VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SL Satellite Crack!"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Code"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Here is the Registration code!"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Place the Request code on this box!"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2580
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text2 = Encode64(Text1, "SL Satellite")
End Sub

Public Function Encode64(Message, Password) As String
    asciinumber = 255
    passwordl = Password
    messagel = Message
    passlength = Len(passwordl)
    messagelength = Len(messagel)
    posistionpass = 0
    posistionmess = 0

    While posistionmess < messagelength
        posistionpass = posistionpass + 1 'move onto the next character in the password

        If posistionpass > passlength Then posistionpass = 1
    
        passchar = Mid(passwordl, posistionpass, 1)
        passcharval = Asc(passchar)
        posistionmess = posistionmess + 1
        messagechar = Mid(messagel, posistionmess, 1)
        messagecharval = Asc(messagechar)
        newchar = messagecharval + passcharval
    
        If newchar > 255 Then newchar = newchar - asciinumber
    
        newmessage = newmessage + Chr(newchar)
    Wend
    Encode64 = newmessage
End Function
