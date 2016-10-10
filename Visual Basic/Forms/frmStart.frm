VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registration"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   3120
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Registration Code"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Request Code"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "To obtain a registration code, you need to PM DAngel (in skylords) or send an Email to sls-payment@indagalaxy.co.uk."
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "To stop unauthorized copies of SL Satellite, you have to register your copy with DAngel and the online DB."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Decode64(Text2.Text, "SL Satellite") = Text1.Text Then
        SetRegValue HKEY_CURRENT_USER, "Software\D-Lords\D-Register", "Regcode", "" & Text2.Text
        MsgBox "Congratulations! You have now registered SL Satellite"
        frmMain.Show
        Unload Me
    Else
        MsgBox "Wrong registration code!"
        MsgBox "Contact DAngel on Skylords for the Registration code!"
    End If
End Sub

Private Sub Form_Load()
    CheckInstall
    If CheckRegistration Then
        frmMain.Show
        Unload Me
    End If
End Sub
