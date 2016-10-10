VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SL Satellite"
   ClientHeight    =   2430
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7695
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   7695
   StartUpPosition =   1  'CenterOwner
   Begin Satellite.ShellIcon ShellIcon1 
      Left            =   6360
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      Icon            =   "frmMain.frx":0E42
   End
   Begin VB.TextBox regVal 
      Height          =   285
      Left            =   6480
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   6960
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "frmMain.frx":1C94
      Top             =   6000
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "frmMain.frx":1C9A
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   4200
      Top             =   2880
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmMain.frx":1CA0
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   "News Alert"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton Command1 
         Caption         =   "Clear Alerts"
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Checking News..."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "No New Updates"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   1800
         Width           =   2535
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences"
      End
      Begin VB.Menu mnuSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuTests 
         Caption         =   "Tests"
         Begin VB.Menu mnuPopup 
            Caption         =   "Popup"
         End
         Begin VB.Menu mnuOverlay 
            Caption         =   "Overlay"
         End
         Begin VB.Menu mnuEmail 
            Caption         =   "Email"
         End
      End
      Begin VB.Menu mnuUpdates 
         Caption         =   "Check for Updates"
      End
      Begin VB.Menu mnuSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuPopup2 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu ppuTests 
         Caption         =   "Tests"
         Begin VB.Menu ppuPopup 
            Caption         =   "Popup"
         End
         Begin VB.Menu ppuOverlay 
            Caption         =   "Overlay"
         End
         Begin VB.Menu ppuEmail 
            Caption         =   "Email"
         End
      End
      Begin VB.Menu mnuSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu ppuOpen 
         Caption         =   "Open SL Satellite"
      End
      Begin VB.Menu ppuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim first As Boolean

Private Sub Command1_Click()
    List1.Clear
End Sub

Private Sub Form_Load()
    GetSettings
    If Command = "/hide" Then
        WindowState = 1
    ElseIf Command <> "" Then
        frmOptions.List1.AddItem (Command)
        SaveSettings
        End
    End If
    cSending = False
    mTest = False
    pNumber = 0
    first = True
    Sort_IE
    If frmOptions.Check2.Value = 1 Then CheckUpdates
    ShellIcon1.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    WindowState = 1
End Sub

Private Sub Form_Resize()
    If WindowState = 1 Then Hide Else Show
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuEmail_Click()
    Call SendMail("", True)
End Sub

Private Sub mnuOverlay_Click()
    frmOverlay.Label2.Caption = "You have prompted to test the Warning Overlay!"
    frmOverlay.Show
End Sub

Private Sub mnuPopup_Click()
        Set tpop = New frmPopup
        tpop.SetNumber 450
        tpop.LblText.Caption = "Popup Test"
        tpop.LblMessage.Caption = "You have prompted to test the popup"
        tpop.LblOptions.Caption = ""
        tpop.Visible = True
End Sub

Private Sub ppuEmail_Click()
    Call SendMail("", True)
End Sub

Private Sub ppuExit_Click()
    ShellIcon1.Visible = False
    End
End Sub

Private Sub ppuOpen_Click()
    WindowState = 0: Show: AppActivate Caption, wait
End Sub

Private Sub ppuOverlay_Click()
    frmOverlay.Label2.Caption = "You have prompted to test the Warning Overlay!"
    frmOverlay.Show
End Sub

Private Sub ppuPopup_Click()
        Set tpop = New frmPopup
        tpop.SetNumber 450
        tpop.LblText.Caption = "Popup Test"
        tpop.LblMessage.Caption = "You have prompted to test the popup"
        tpop.LblOptions.Caption = ""
        tpop.Visible = True
End Sub

Private Sub ShellIcon1_DblClick(Button As Integer)
    If Button = 1 Then WindowState = 0: Show: AppActivate Caption, wait
End Sub

Private Sub ShellIcon1_SingleClick(Button As Integer)
    PopupMenu mnuPopup2, 2
End Sub

Private Sub mnuExit_Click()
    ShellIcon1.Visible = False
    End
End Sub

Private Sub mnuMinimize_Click()
    WindowState = 1
End Sub

Private Sub mnuPreferences_Click()
    GetSettings
    frmOptions.Show
End Sub

Private Sub mnuUpdates_Click()
    CheckUpdates
End Sub

Private Sub Timer1_Timer()
    Me.Caption = "SL Satellite - Checking News"
    Text3 = OldNews
    Label2 = "...................."
    Text1 = "...................."
    Text2 = "*********************"
    DoEvents
    Text1 = GetNews
    DoEvents
    Text2 = GetNew
    If Not Text1 = "Offline" Then
        If first Then
            first = False
        Else
            If Text2 <> "" Then
                temp = Split(Text2, vbNewLine)
                
                For i = UBound(temp) To 0 Step -1
                    SortNews (temp(i))
                Next
            End If
        End If
        Me.Caption = "SL Satellite "
        Label2 = "Retrieved news successfuly"
    Else
        Me.Caption = "SL Satellite - Offline"
        Label2 = "Offline"
    End If
End Sub
