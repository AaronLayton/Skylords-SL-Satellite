Attribute VB_Name = "Settings"
'This function will get all of the settings and enter
'them into the options pannel
Public Function GetSettings()
    On Error Resume Next
    
    'Gets the Watch List settings
    frmMain.regVal = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "wList")
    If Not frmMain.regVal = "!Empty!" Then
        frmOptions.List1.Clear
        If InStr(frmMain.regVal, "[+]") Then
            tmp = Split(frmMain.regVal, "[+]")
            For i = 0 To UBound(tmp)
                frmOptions.List1.AddItem (tmp(i))
            Next
        Else
            frmOptions.List1.AddItem frmMain.regVal
        End If
    End If
    '***************************************
    
    'Checks if the program starts with window
    frmMain.regVal = GetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "SL Satellite")
    If frmMain.regVal <> "" Then frmOptions.Check1.Value = 1 Else frmOptions.Check1.Value = 0
    '***************************************
    
    'Check to see if the user wants the updates on startup
    frmMain.regVal = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "cUpdates")
    If frmMain.regVal = "True" Then frmOptions.Check2.Value = 1 Else frmOptions.Check2.Value = 0
    '***************************************
    
    'Check if the user wants to filter their own events
    frmMain.regVal = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "fAlerts")
    If frmMain.regVal = "True" Then frmOptions.Check7.Value = 1 Else frmOptions.Check7.Value = 0
    '***************************************
    
    'Get the refresh rate
    frmOptions.Text1 = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Refresh")
    frmMain.Timer1.Interval = CInt(frmOptions.Text1) * 1000
    '***************************************
    
    'Checks to see if the user wants to play sounds
    frmMain.regVal = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Sounds")
    If frmMain.regVal = "True" Then frmOptions.Check3.Value = 1 Else frmOptions.Check3.Value = 0
    '***************************************
    
    'Checks to see if the user wants Popups
    frmMain.regVal = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Popup")
    If frmMain.regVal = "True" Then frmOptions.Check4.Value = 1 Else frmOptions.Check4.Value = 0
    '***************************************
    
    'Checks to see if the user wants Overlays
    frmMain.regVal = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Overlay")
    If frmMain.regVal = "True" Then frmOptions.Check5.Value = 1 Else frmOptions.Check5.Value = 0
    '***************************************
    
    'Checks to see if the user wants to send EMails
    frmMain.regVal = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "sMail")
    If frmMain.regVal = "True" Then frmOptions.Check6.Value = 1 Else frmOptions.Check6.Value = 0
    '***************************************
    
    'Gets the SMTP address
    frmOptions.Text3 = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "SMTP")
    If frmOptions.Text3 = "!Empty!" Then frmOptions.Text3 = ""
    '***************************************
    
    'Gets the EMail address
    frmOptions.Text4 = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "EMail")
    If frmOptions.Text4 = "!Empty!" Then frmOptions.Text4 = ""
    '***************************************
    
    'Checks to see if the user wants to record their Alerts
    frmMain.regVal = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "rAlerts")
    If frmMain.regVal = "True" Then frmOptions.Check8.Value = 1 Else frmOptions.Check8.Value = 0
    '***************************************
    
    'Checks to see if the user wants to record ALL Alerts
    frmMain.regVal = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "rAll")
    If frmMain.regVal = "True" Then frmOptions.Check9.Value = 1 Else frmOptions.Check9.Value = 0
    '***************************************
    
    'Gets the Time Offset
    frmOptions.Text6 = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Offset")
    '***************************************
    
    'Gets the EMail address
    frmOptions.Text5 = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "sPath")
    If frmOptions.Text5 = "!Empty!" Then frmOptions.Text5 = ""
    '***************************************
    
    'Gets the current Username
    frmOptions.Text7 = GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Username")
    If frmOptions.Text7 = "!Empty!" Then frmOptions.Text7 = ""
    '***************************************
End Function

Public Function SaveSettings()
    'Saves the Watch List settings
    tmp = ""
    For i = 0 To frmOptions.List1.ListCount - 1
        tmp = tmp & frmOptions.List1.List(i) & "[+]"
    Next
    
    If tmp = "" Then
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "wList", "!Empty!"
    Else
        tmp = Left(tmp, Len(tmp) - 3)
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "wList", "" & tmp
    End If
    '***************************************
    
    'Saves the program to registry (or deletes it from)
    If frmOptions.Check1.Value = 1 Then
        SetRegValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "SL Satellite", """" & App.Path & "\" & App.EXEName & ".exe"" /hide"
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "SL Satellite"
    End If
    '***************************************
    
    'Saves checking updates
    If frmOptions.Check2.Value = 1 Then
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "cUpdates", "True"
    Else
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "cUpdates", "False"
    End If
    '***************************************
    
    'Saves wether the user wants to filter their own events
    If frmOptions.Check7.Value = 1 Then
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "fAlerts", "True"
    Else
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "fAlerts", "False"
    End If
    '***************************************
    
    'Save the refresh rate
    SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Refresh", frmOptions.Text1
    '***************************************
    
    'Saves wether the user wants to play sounds
    If frmOptions.Check3.Value = 1 Then
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Sounds", "True"
    Else
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Sounds", "False"
    End If
    '***************************************
    
    'Saves wether the user wants Popups
    If frmOptions.Check4.Value = 1 Then
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Popup", "True"
    Else
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Popup", "False"
    End If
    '***************************************
    
    'Saves wether the user wants Overlays
    If frmOptions.Check5.Value = 1 Then
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Overlay", "True"
    Else
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Overlay", "False"
    End If
    '***************************************
    
    'Saves wether the user wants to send Emails
    If frmOptions.Check6.Value = 1 Then
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "sMail", "True"
    Else
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "sMail", "False"
    End If
    '***************************************
    
    'Saves the SMTP value
    If frmOptions.Text3 <> "" Then
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "SMTP", frmOptions.Text3
    Else
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "SMTP", "!Empty!"
    End If
    '***************************************
    
    'Saves the EMail value
    If frmOptions.Text4 <> "" Then
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "EMail", frmOptions.Text4
    Else
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "EMail", "!Empty!"
    End If
    '***************************************
    
    'Saves wether the user wants to record their Alerts
    If frmOptions.Check8.Value = 1 Then
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "rAlerts", "True"
    Else
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "rAlerts", "False"
    End If
    '***************************************
    
    'Saves wether the user wants to record ALL Alerts
    If frmOptions.Check9.Value = 1 Then
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "rAll", "True"
    Else
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "rAll", "False"
    End If
    '***************************************
    
    'Saves the time offset
    SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Offset", frmOptions.Text6
    '***************************************
    
    'Saves the Save Path
    If frmOptions.Text5 <> "" Then
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "sPath", frmOptions.Text5
    Else
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "sPath", "!Empty!"
    End If
    '***************************************
    
    'Saves the current Username
    If frmOptions.Text7 <> "" Then
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Username", frmOptions.Text7
    Else
        SetRegValue HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Username", "!Empty!"
    End If
    '***************************************
End Function

