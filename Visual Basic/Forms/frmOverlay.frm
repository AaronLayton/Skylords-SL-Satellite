VERSION 5.00
Begin VB.Form frmOverlay 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9660
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   72
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   1920
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left Click to goto Skylords.com - Right Click to Dismiss"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2985
      TabIndex        =   2
      Top             =   1440
      Width           =   3915
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message to be displayed!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3525
      TabIndex        =   1
      Top             =   1080
      Width           =   2880
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1350
      Left            =   2160
      TabIndex        =   0
      Top             =   -120
      Width           =   5655
   End
End
Attribute VB_Name = "frmOverlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event MouseDown(Button As Integer)

Dim Trans
Dim start As Boolean
Dim gUp As Boolean

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Unload Me
    Else
        Shell ("explorer http://www.skylords.com/login"), vbMaximizedFocus
        Unload Me
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Unload Me
    Else
        Shell ("explorer http://www.skylords.com/login"), vbMaximizedFocus
        Unload Me
    End If
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Unload Me
    Else
        Shell ("explorer http://www.skylords.com/login"), vbMaximizedFocus
        Unload Me
    End If
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Unload Me
    Else
        Shell ("explorer http://www.skylords.com/login"), vbMaximizedFocus
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Trans = 0
    MakeTransparent Me.hWnd, (Trans)
    Me.Height = 1800
    Me.Width = Screen.Width
    Me.Left = 0
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Label1.Left = (Me.Width / 2) - (Label1.Width / 2)
    Label2.AutoSize = False
    Label2.AutoSize = True
    Label2.Left = (Me.Width / 2) - (Label2.Width / 2)
    Label3.Left = (Me.Width / 2) - (Label3.Width / 2)
    start = True
    gUp = False
    AlwaysOnTop Me, True
    If frmOptions.Check3 = 1 Then
        PlaySound (App.Path & "\overlay.wav")
    End If
End Sub

Private Sub Timer1_Timer()
    If start Then
        Trans = Trans + 2
        MakeTransparent Me.hWnd, (Trans)
    Else
        If gUp Then
            Trans = Trans + 1
            MakeTransparent Me.hWnd, (Trans)
            If Trans > 180 Then gUp = False
        Else
            Trans = Trans - 1
            MakeTransparent Me.hWnd, (Trans)
            If Trans < 130 Then gUp = True
        End If
    End If
    If Trans > 180 Then start = False
End Sub
