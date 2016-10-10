Attribute VB_Name = "News"
'Functions are copywrited to indaGalaxy.co.uk
'you are not permitted to use these in other projects

Dim XHT As XMLHTTPRequest
Global OldNews
Global NewNews
Global CurrentNews
Global isOnline As Boolean
Dim lStart
Dim nLength


Public Function GetNews()

    On Local Error Resume Next
    
    Set XHT = Nothing
    Set XHT = New XMLHTTPRequest
    DoEvents
    XHT.Open "GET", "http://www.skylords.com/news?time=" & Now(), False
    DoEvents
    XHT.send ""
    DoEvents
    
    If XHT.responseText <> "" And InStr(XHT.responseText, "SkyLords Online multiplayer strategy game: real time global, clan and personal events, archived news") Then
        isOnline = True
        CurrentNews = Replace(XHT.responseText, "<br />", vbNewLine)
        lStart = InStr(CurrentNews, "shown.</p>") + 10
        nLength = InStr(CurrentNews, "<p id=""pager"">") - lStart - 2
    
        CurrentNews = Mid(CurrentNews, lStart, nLength)
        
        GetNews = CurrentNews
    Else
        DoEvents
        XHT.Open "GET", "http://www.skylords.net/news?time=" & Now(), False
        DoEvents
        XHT.send ""
        DoEvents
        
        If XHT.responseText <> "" And InStr(XHT.responseText, "SkyLords Online multiplayer strategy game: real time global, clan and personal events, archived news") Then
            isOnline = True
            CurrentNews = Replace(XHT.responseText, "<br />", vbNewLine)
            lStart = InStr(CurrentNews, "shown.</p>") + 10
            nLength = InStr(CurrentNews, "<p id=""pager"">") - lStart - 2
            
            CurrentNews = Mid(CurrentNews, lStart, nLength)
            
            GetNews = CurrentNews
        Else
            isOnline = False
            GetNews = "Offline"
        End If
    End If
End Function

'This function returns the NEW news
Public Function GetNew()
    NewNews = ""
    tmp = Split(CurrentNews, vbNewLine)
    
    For X = 0 To UBound(tmp)
        If InStr(OldNews, tmp(X)) = 0 Then
            NewNews = NewNews & tmp(X) & vbNewLine
        End If
    Next X
    
    If InStr(NewNews, vbNewLine) Then
        NewNews = Left(NewNews, Len(NewNews) - 2)
    End If
    
    OldNews = CurrentNews
    GetNew = NewNews
End Function

'Function to sort out the new News
Public Function SortNews(News)
    News = Offset(News)
    
    'Run checks for the username
    If InStr(News, frmOptions.Text7) Then
        If Is_Bad(News) Then
            'Temporary load all warnings
            Call Load_Popup("Warning", News, "")
            Call Load_Overlay(News)
            Call SendMail(News, False)
            If frmOptions.Check8 = 1 Then
                Call Archive_It(News)
            End If
            frmMain.List1.AddItem News, 0
            Exit Function
        Else
            If frmOptions.Check7 = 0 Then
                Call Load_Popup("Warning", News, "")
                frmMain.List1.AddItem News, 0
                If frmOptions.Check8 = 1 Then
                    Call Archive_It(News)
                End If
                Exit Function
            End If
        End If
    End If
    
    'Run checks for the watchlist
    For X = 0 To frmOptions.List1.ListCount - 1
        If InStr(News, frmOptions.List1.List(X)) Then
            'Add to main window
            frmMain.List1.AddItem News, 0
            
            'Call the popup.
            Call Load_Popup("Alert", News, "")
            
            'Check to see if it needs archiving.
            'Only need to check this one, if rec all
            'then check8 is on by default.
            If frmOptions.Check8 = 1 Then
                Call Archive_It(News)
            End If
            Exit Function
        End If
    Next
    
    'If its not in the above checks then
    'check to see if the user wants to
    'record ALL news
    If frmOptions.Check9 = 1 Then
        Call Archive_It(News)
    End If
    
    

End Function

Function Is_Bad(News) As Boolean
    Is_Bad = False
    'Check for certain news types
    If InStr(News, "has found the planet of " & frmOptions.Text7 & ".") Then
        Is_Bad = True
    End If
    
    If InStr(News, frmOptions.Text7 & " planet") And InStr(News, "has been captured by") Then
        Is_Bad = True
    End If
    
    If InStr(News, frmOptions.Text7 & " has been defeated by") Then
        Is_Bad = True
    End If
    
    If InStr(News, frmOptions.Text7 & " planet") And InStr(News, "has been detected.") Then
        Is_Bad = True
    End If
    
    If InStr(News, frmOptions.Text7 & " ship") And InStr(News, "has been destroyed") Then
        Is_Bad = True
    End If
    
    If InStr(News, frmOptions.Text7 & " planet") And InStr(News, "has been attacked by") Then
        Is_Bad = True
    End If
End Function


