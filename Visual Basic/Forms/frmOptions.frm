VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4950
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   3960
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Options"
      TabPicture(0)   =   "frmOptions.frx":1472
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Watch List"
      TabPicture(1)   =   "frmOptions.frx":148E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Alerts"
      TabPicture(2)   =   "frmOptions.frx":14AA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Logging"
      TabPicture(3)   =   "frmOptions.frx":14C6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Username"
      TabPicture(4)   =   "frmOptions.frx":14E2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame6 
         Caption         =   "Skylords Username"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   44
         Top             =   360
         Width           =   4455
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   480
            TabIndex        =   46
            Top             =   1080
            Width           =   3855
         End
         Begin VB.Label Label18 
            Caption         =   "Enter your Skylords username, for checking news events. This was we can alert you of high priority alerts."
            Height          =   615
            Left            =   480
            TabIndex        =   45
            Top             =   360
            Width           =   3735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Loggin Options"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   4455
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   480
            TabIndex        =   43
            Text            =   "0"
            Top             =   1440
            Width           =   3615
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   285
            Left            =   4080
            TabIndex        =   42
            Top             =   1440
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Browse"
            Height          =   255
            Left            =   3480
            TabIndex        =   40
            Top             =   2760
            Width           =   855
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Check9"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Check8"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   480
            TabIndex        =   35
            Top             =   2400
            Width           =   3855
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Time Offset"
            Height          =   195
            Left            =   480
            TabIndex        =   41
            Top             =   1200
            Width           =   810
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Records ALL News"
            Height          =   195
            Left            =   480
            TabIndex        =   39
            Top             =   750
            Width           =   1380
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Records Alerts"
            Height          =   195
            Left            =   480
            TabIndex        =   38
            Top             =   270
            Width           =   1035
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Save Location"
            Height          =   255
            Left            =   480
            TabIndex        =   34
            Top             =   2040
            Width           =   1035
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Alert Options"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   4455
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   480
            TabIndex        =   31
            Top             =   2040
            Width           =   3855
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   480
            TabIndex        =   27
            Top             =   1440
            Width           =   3855
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Check6"
            Height          =   255
            Left            =   2400
            TabIndex        =   24
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Check5"
            Height          =   255
            Left            =   2400
            TabIndex        =   20
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Check4"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Check3"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Email Address"
            Height          =   195
            Left            =   480
            TabIndex        =   32
            Top             =   1800
            Width           =   990
         End
         Begin VB.Label Label12 
            Height          =   615
            Left            =   480
            TabIndex        =   30
            Top             =   2520
            Width           =   3855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "SMTP Server"
            Height          =   195
            Left            =   480
            TabIndex        =   26
            Top             =   1200
            Width           =   960
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Send Emails"
            Height          =   195
            Left            =   2760
            TabIndex        =   25
            Top             =   750
            Width           =   870
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Allow Window Overlay"
            Height          =   195
            Left            =   2760
            TabIndex        =   23
            Top             =   270
            Width           =   1590
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Allow Popup Alert"
            Height          =   195
            Left            =   480
            TabIndex        =   22
            Top             =   750
            Width           =   1245
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Play sounds"
            Height          =   195
            Left            =   480
            TabIndex        =   21
            Top             =   270
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3255
         Left            =   -72480
         TabIndex        =   11
         Top             =   360
         Width           =   2055
         Begin VB.Label Label5 
            Height          =   2175
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Users"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   10
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton Command3 
            Caption         =   "Remove"
            Height          =   255
            Left            =   1080
            TabIndex        =   16
            Top             =   2880
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Add"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   2880
            Width           =   975
         End
         Begin VB.ListBox List1 
            Height          =   2400
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Main Options"
         Height          =   3255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4455
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   255
            Left            =   4080
            TabIndex        =   47
            Top             =   1920
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Check7"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "5"
            Top             =   1920
            Width           =   3615
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Check2"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Filter your own actions from alerts"
            Height          =   195
            Left            =   480
            TabIndex        =   29
            Top             =   1230
            Width           =   2340
         End
         Begin VB.Label Label4 
            Caption         =   $"frmOptions.frx":14FE
            Height          =   675
            Left            =   480
            TabIndex        =   9
            Top             =   2400
            Width           =   3720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Refresh rate (seconds)"
            Height          =   195
            Left            =   480
            TabIndex        =   7
            Top             =   1680
            Width           =   1605
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Check for updates on startup"
            Height          =   195
            Left            =   480
            TabIndex        =   6
            Top             =   750
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Start with windows"
            Height          =   195
            Left            =   480
            TabIndex        =   5
            Top             =   270
            Width           =   1320
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_BROWSEFORCOMPUTER = &H1000

Private Const MAX_PATH = 260


Private Declare Function SHBrowseForFolder Lib "shell32" _
    (lpbi As BrowseInfo) As Long


Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long


Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal _
    lpString2 As String) As Long


Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
    End Type
    
    
'***************************************************************
'***************************************************************

Private Sub Check6_Click()
    If Check6.Value = 0 Then
        Text3.Enabled = False
        Text4.Enabled = False
    End If
    If Check6.Value = 1 Then
        Text3.Enabled = True
        Text4.Enabled = True
    End If
End Sub

Private Sub Check8_Click()
    If Check8.Value = 0 Then
        Text5.Enabled = False
        Command4.Enabled = False
        Check9.Value = 0
        End If
    If Check8.Value = 1 Then
        Text5.Enabled = True
        Command4.Enabled = True
    End If
End Sub

Private Sub Check9_Click()
    If Check9.Value = 1 Then
        Check8.Value = 1
        Text5.Enabled = True
        Command4.Enabled = True
    End If
End Sub

Private Sub Command1_Click()
    If Text1 = "" Then
        MsgBox "Enter a Refresh rate in seconds"
        Exit Sub
    End If
    If Check6.Value = 1 Then
        If InStr(Text3, ".") = 0 Or InStr(Text3, ":") = 0 Then
            MsgBox "mail.server.com:port"
            Exit Sub
        End If
        If InStr(Text4.Text, "@") = 0 Then
            MsgBox "you@yourplace.com"
            Exit Sub
        End If
    End If
    If Text5 = "" Then Text5 = App.Path & "\News"
    If Text7 = "" Then
        MsgBox "Please enter your Skylords Username"
        Exit Sub
    End If
    frmMain.Timer1.Interval = Text1 * 1000
    SaveSettings
    Me.Hide
End Sub

Private Sub Command2_Click()
    If Text2 <> "" Then
        List1.AddItem Text2, 0
        Text2 = ""
    Else
        MsgBox "Cant add a blank value"
    End If
End Sub

Private Sub Command3_Click()
    Text2 = List1.List(List1.ListIndex)
    If List1.ListIndex = -1 Then Exit Sub
    List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub Command4_Click()
    
    'Opens a Treeview control that displays
    '     the directories in a computer
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = "This is the title"


    With tBrowseInfo
        .hwndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS _
        + BIF_DONTGOBELOWDOMAIN
        
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)


    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        Text5 = sBuffer
    Else
        If Check8.Value = 1 Then
            If Text5 = "" Then
                MsgBox "Please select a Save path!"
            End If
        End If
    End If
    

End Sub

Private Sub Command5_Click()
MsgBox frmOptions.List1.ListCount
End Sub

Private Sub Form_Load()
    If Check6.Value = 0 Then
        Text3.Enabled = False
        Text4.Enabled = False
        Command4.Enabled = False
    End If
    If Check6.Value = 1 Then
        Text3.Enabled = True
        Text4.Enabled = True
    End If
    If Check8.Value = 0 Then Text5.Enabled = False
    'GetSettings
    Label5.Caption = "Add users to the 'Watch List' to be alerted when they are in the news. Perfect for watching watching over noobie's." & vbNewLine & vbNewLine & "Integrated 'Right Click' in Internet Explorer only."
    Label12.Caption = "Normal alerts will be show if you are at your computer." & vbNewLine & "If you are away from your computer an email will be sent (Perfect if your at work)."
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.Hide
End Sub

Private Sub UpDown1_DownClick()
    Text6 = Text6 - 1
    If Text6 < -12 Then Text6 = -12
End Sub

Private Sub UpDown1_UpClick()
    Text6 = Text6 + 1
    If Text6 > 12 Then Text6 = 12
End Sub

Private Sub UpDown2_DownClick()
    Text1.Text = Text1.Text - 1
    If Text1 < 5 Then Text1 = 5
End Sub

Private Sub UpDown2_UpClick()
    Text1 = Text1 + 1
    If Text1 > 30 Then Text1 = 30
End Sub
