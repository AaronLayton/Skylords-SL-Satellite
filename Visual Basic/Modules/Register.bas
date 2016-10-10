Attribute VB_Name = "Register"
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Function CheckInstall()
    'Code to check if the install is ok
    If GetRegValue(HKEY_CURRENT_USER, "Software\indaGalaxy\SL Satellite", "Username") = "" Then
        MsgBox "Please check your installation!"
        End
    End If
End Function

Public Function CheckRegistration() As Boolean
    CheckRegistration = False
    If GetRegValue(HKEY_CURRENT_USER, "Software\D-Lords\D-Register", "Request") = "" Then
        MakePass
        EncryptIt
        frmStart.Text1 = Decode64(GetRegValue(HKEY_CURRENT_USER, "Software\D-Lords\D-Register", "Request"), GetTheName)
    Else
        strName = GetTheName
        tmp = GetRegValue(HKEY_CURRENT_USER, "Software\D-Lords\D-Register", "Request")
        frmStart.Text1.Text = Decode64(tmp, strName)
    End If
    
    If GetRegValue(HKEY_CURRENT_USER, "Software\D-Lords\D-Register", "Regcode") = "" Then
        CheckRegistration = False
    Else
        tmp = GetRegValue(HKEY_CURRENT_USER, "Software\D-Lords\D-Register", "Regcode")
        If Decode64(tmp, "SL Satellite") = frmStart.Text1.Text Then CheckRegistration = True
    End If
    
End Function

Public Function MakePass() As String
frmStart.Text4.Text = ""
d = 10
Do
    Randomize
        c = Int(Rnd * 2)
        If c = 1 Then
            a = (Int(Rnd * 9))
 
            frmStart.Text4.SelStart = Len(frmStart.Text4.Text)
            frmStart.Text4.SelText = a
            frmStart.Text4.SelStart = Len(frmStart.Text4.Text)
        ElseIf c = 0 Then
            a = Int((Rnd * 25 + 65))
            a = Chr(a)
            a = StrConv(a, vbLowerCase)
            frmStart.Text4.Text = frmStart.Text4.Text + a
        End If
Loop Until Len(frmStart.Text4.Text) = d
End Function

Public Function GetTheName() As String
    'This part gets the Computer Username to encrpt with
    Dim strName As String
    Dim lngBuffer As Long
  
    strName = String$(255, 0)
    lngBuffer = GetUserName(strName, Len(strName))
    strName = Trim(strName)
    If strName = "" Then strName = "google"
    
    GetTheName = strName
End Function

Public Function EncryptIt()
    strName = GetTheName
    'Do the encoding
    theval = Encode64(frmStart.Text4.Text, strName)
    'Save it to the registry
    CreateRegistryKey HKEY_CURRENT_USER, "Software\D-Lords\D-Register"
    SetRegValue HKEY_CURRENT_USER, "Software\D-Lords\D-Register", "Request", "" & theval
End Function

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

Public Function Decode64(Message, Password) As String
    
    asciinumber = 255
    passwordl = Password
    messagel = Message
    passlength = Len(passwordl)
    messagelength = Len(messagel)
    posistionpass = 0
    posistionmess = 0

    While posistionmess < messagelength
        posistionpass = posistionpass + 1
        
        If posistionpass > passlength Then posistionpass = 1
        passchar = Mid(passwordl, posistionpass, 1)
        passcharval = Asc(passchar)
        posistionmess = posistionmess + 1
        messagechar = Mid(messagel, posistionmess, 1)
        messagecharval = Asc(messagechar)
        
        If messagecharval < passcharval Then messagecharval = messagecharval + asciinumber
        newchar = messagecharval - passcharval
        newmessage = newmessage + Chr(newchar)
    Wend
 
    Decode64 = newmessage
End Function
