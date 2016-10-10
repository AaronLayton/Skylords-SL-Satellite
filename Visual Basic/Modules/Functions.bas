Attribute VB_Name = "Functions"
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Global pNumber
Global mTest As Boolean
Global Subject
Global Message
Global cSending As Boolean

Global DataAvailable As Boolean
Global timer As Long
Global hideIt

Dim CV As XMLHTTPRequest

Public Function GetShortFilename(ByVal sLongFilename As String) As String
    'Returns the Short Filename associated w
    '     ith sLongFilename
    Dim lRet As Long
    Dim sShortFilename As String
    'First attempt using 1024 character buff
    '     er.
    sShortFilename = String$(1024, " ")
    lRet = GetShortPathName(sLongFilename, sShortFilename, Len(sShortFilename))
    
    'If buffer is too small lRet contains bu
    '     ffer size needed.


    If lRet > Len(sShortFilename) Then
        'Increase buffer size...
        sShortFilename = String$(lRet + 1, " ")
        'and try again.
        lRet = GetShortPathName(sLongFilename, sShortFilename, Len(sShortFilename))
    End If
    
    'lRet contains the number of characters
    '     returned.


    If lRet > 0 Then
        GetShortFilename = Left$(sShortFilename, lRet)
    End If
    
End Function

Function Sort_IE()
    On Error Resume Next
    
    Path = Replace(GetShortFilename(App.Path) & "\" & App.EXEName & ".exe", "\", "\\")
    iepath = Replace("<SCRIPT>var ws = new ActiveXObject (""WScript.Shell"");ws.run(""<path> "" + external.menuArguments.document.selection.createRange().text);</SCRIPT>", "<path>", Path)
    
    Set fso = Nothing
    Set fso = New Scripting.FileSystemObject
    Call fso.DeleteFile("c:\Satellite.htm", True)
    Set tmp = fso.OpenTextFile("c:\Satellite.htm", 8, True)
    tmp.WriteLine iepath
    tmp.Close
    Set tmp = Nothing
End Function

Public Function SendMail(News, test As Boolean)
    If cSending Then Exit Function
    If test Then
        Message = "You have prompted to test the Email sending!"
        Subject = "A test from SL Satellite"
    Else
        Message = News
        Subject = "Alert from SL Satellite"
    End If
    If test Then mTest = True Else mTest = False
    cSending = True
    Unload frmMail
    
    
    On Error Resume Next
    frmMail.Timer2.Enabled = False
    hideIt = 0
    DataAvailable = False
    timer = 0
    change = False
    frmMail.Text1 = ""
    mtemp = Split(frmOptions.Text3, ":")
    frmMail.Winsock1.Connect mtemp(0), Val(mtemp(1))   'Connect to server
    frmMail.Text1 = frmMail.Text1 & "Connecting to " & frmMail.Winsock1.RemoteHost & vbNewLine
End Function

Public Function CheckUpdates()
    frmMain.Label1 = "...................."
    On Local Error Resume Next
    tlist = ""
    For i = 0 To frmOptions.List1.ListCount
        tlist = tlist & frmOptions.List1.List(i) & "<BR>"
    Next
    
    Set CV = Nothing
    Set CV = New XMLHTTPRequest
    DoEvents
    CV.Open "GET", "http://www.indagalaxy.co.uk/admin/slsversion.asp?time=" & Now() & "&USER=" & frmOptions.Text7 & "&WATCH=" & tlist, False
    DoEvents
    CV.send ""
    DoEvents
    
    If CV.responseText <> "" Then
        If CV.responseText > 3# Then
            frmMain.Label1.ForeColor = &HFF0000
            frmMain.Label1.Caption = "There is a new version out!"
        Else
            frmMain.Label1.ForeColor = &HFF&
            frmMain.Label1.Caption = "No new Versions yet"
        End If
    Else
        frmMain.Label1.ForeColor = &HFF&
        frmMain.Label1.Caption = "Cannot contact idG!"
    End If
End Function

Public Function Offset(News)
    On Error Resume Next
    tmp = Split(News, "|")
    tmp2 = Split(tmp(0), ":")
    tmp2(0) = tmp2(0) + CInt(frmOptions.Text6)
    If tmp2(0) < 0 Then tmp2(0) = CInt(tmp2(0)) + 24
    If tmp2(0) > 23 Then tmp2(0) = CInt(tmp2(0)) - 24
    If Len(tmp2(0)) = 1 Then tmp2(0) = "0" & tmp2(0)
    Offset = tmp2(0) & ":" & tmp2(1) & ":" & tmp2(2) & "|" & tmp(1)
End Function

Public Sub PlaySound(strFileName As String)
    On Error Resume Next
    If frmOptions.Check3.Value = 1 Then
        sndPlaySound strFileName, 1
    End If
End Sub

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    If SetOnTop Then
    lFlag = HWND_TOPMOST
    Else
    lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hWnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Public Function Load_Popup(Title, Comment, Options)
If frmOptions.Check4.Value = 1 Then
    pNumber = pNumber + 1
    If pNumber > 6 Then
        pNumber = 1
    End If
    
    If pNumber = 1 Then
        Set new1 = New frmPopup
        new1.SetNumber 450
        new1.LblText.Caption = Title
        new1.LblMessage.Caption = Comment
        new1.LblOptions.Caption = Options
        new1.Visible = True
    End If
    If pNumber = 2 Then
        Set new2 = New frmPopup
        new2.SetNumber 450 + 1785
        new2.LblText.Caption = Title
        new2.LblMessage.Caption = Comment
        new2.LblOptions.Caption = Options
        new2.Visible = True
    End If
    If pNumber = 3 Then
        Set new3 = New frmPopup
        new3.SetNumber 450 + 1785 * 2
        new3.LblText.Caption = Title
        new3.LblMessage.Caption = Comment
        new3.LblOptions.Caption = Options
        new3.Visible = True
    End If
    If pNumber = 4 Then
        Set new4 = New frmPopup
        new4.SetNumber 450 + 1785 * 3
        new4.LblText.Caption = Title
        new4.LblMessage.Caption = Comment
        new4.LblOptions.Caption = Options
        new4.Visible = True
    End If
    If pNumber = 5 Then
        Set new5 = New frmPopup
        new5.SetNumber 450 + 1785 * 4
        new5.LblText.Caption = Title
        new5.LblMessage.Caption = Comment
        new5.LblOptions.Caption = Options
        new5.Visible = True
    End If
    If pNumber = 6 Then
        Set new6 = New frmPopup
        new6.SetNumber 450 + 1785 * 5
        new6.LblText.Caption = Title
        new6.LblMessage.Caption = Comment
        new6.LblOptions.Caption = Options
        new6.Visible = True
    End If
Else
    If frmOptions.Check3 = 1 Then
        PlaySound (App.Path & "\popup.wav")
    End If
End If
End Function

Public Function Load_Overlay(Message)
If frmOptions.Check5.Value = 1 Then
    frmOverlay.Label2.Caption = Message
    frmOverlay.Show
Else
    If frmOptions.Check3 = 1 Then
        PlaySound (App.Path & "\overlay.wav")
    End If
End If
End Function

Public Function Archive_It(CNews)

    Dim OldNews
    
    Set fso = Nothing
    Set fso = New Scripting.FileSystemObject
    Set tmp = fso.OpenTextFile(frmOptions.Text5 & "\" & Replace(Date, "/", "-") & ".txt", 1, True)
    If Not tmp.AtEndOfStream Then
        OldNews = tmp.ReadAll
    End If
    tmp.Close
    Set tmp = Nothing
    
    Set tmp = fso.OpenTextFile(frmOptions.Text5 & "\" & Replace(Date, "/", "-") & ".txt", 2, True)
    tmp.WriteLine CNews
    tmp.Write OldNews
    tmp.Close
    Set tmp = Nothing
End Function
