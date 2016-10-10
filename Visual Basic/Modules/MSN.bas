Attribute VB_Name = "MSN"
Public Declare Function receivedmessage Lib "user32" Alias "ReceivedA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal param As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function receivedmessagelong& Lib "user32" Alias "ReceivedA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal param As Long)
Public Const WM_SETTEXT = &HC
Public Const WM_COMMAND = &H111
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const VK_PAUSE = &H13
Public Const VK_SPACE = &H20
Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
    Public Const SW_NORMAL = 1

Public Function MSN_FindMSN()
Dim bah As Long
bah = FindWindow("MSNMSBLClass", vbNullString)
MSN_FindMSN = bah
End Function

'PostMessage MSN_FindMSN, WM_COMMAND, 40210, 0
'offline niet status
'PostMessage MSN_FindMSN, WM_COMMAND, 40229, 0
'mijn ontvangen bestanden
'PostMessage MSN_FindMSN, WM_COMMAND, 40287, 0
'contactpersonen opslaan
'PostMessage MSN_FindMSN, WM_COMMAND, 40256, 0
'contactpersonen van hotmail openen
'PostMessage MSN_FindMSN, WM_COMMAND, 40158, 0
'controle of hotmail eruit ligt...
'PostMessage MSN_FindMSN, WM_COMMAND, 40274, 0
'videogesprek starten..
'PostMessage MSN_FindMSN, WM_COMMAND, 40205, 0
'een tekstbericht verzenden
'PostMessage MSN_FindMSN, WM_COMMAND, 40257, 0
'mijn aangepaste emoticons
'PostMessage MSN_FindMSN, WM_COMMAND, 40268, 0
'nickname veranderen
'PostMessage MSN_FindMSN, WM_COMMAND, 40039, 0
'nog niets gezien niemand is online..:s denk voor flooder
'PostMessage MSN_FindMSN, WM_COMMAND, 40198, 0
'postvak in
'PostMessage MSN_FindMSN, WM_COMMAND, 40229, 0
'mijn ontvagen bestanden..:s 2x
'PostMessage MSN_FindMSN, WM_COMMAND, 40271, 0
'groep toevoegen
'PostMessage MSN_FindMSN, WM_COMMAND, 40282, 0
'contactpersoon toevoegen
'PostMessage MSN_FindMSN, WM_COMMAND, 40301, 0
'msncontact persoon zoeken...
'PostMessage MSN_FindMSN, WM_COMMAND, 40003, 0
'alle gesprek venster sluiten en msn minimalisren
'PostMessage MSN_FindMSN, WM_COMMAND, 40266, 0
'telefoon in opties
'PostMessage MSN_FindMSN, WM_COMMAND, 40289, 0
'een contactpersoon verwijderen
'PostMessage MSN_FindMSN, WM_COMMAND, 40284, 0
'eigenschappen van een contact presoon bekijken..
'PostMessage MSN_FindMSN, WM_COMMAND, 40283, 0
'profiel van een contact presoon bekijken..
'PostMessage MSN_FindMSN, WM_COMMAND, 40168, 0
'bezet
'PostMessage MSN_FindMSN, WM_COMMAND, 40169, 0
'zo terug
'PostMessage MSN_FindMSN, WM_COMMAND, 40170, 0
'afwezig
'PostMessage MSN_FindMSN, WM_COMMAND, 40172, 0
'lunchpauze
'PostMessage MSN_FindMSN, WM_COMMAND, 40171, 0
'aan de telefoon
'PostMessage MSN_FindMSN, WM_COMMAND, 40166, 0
'online
'PostMessage MSN_FindMSN, WM_COMMAND, 40167, 0
'offline weergeven
'PostMessage MSN_FindMSN, WM_COMMAND, 40254, 0
'whitboard starten
'PostMessage MSN_FindMSN, WM_COMMAND, 40269, 0
'verberg balk in msn om contactpersoon toetevoegen
'PostMessage MSN_FindMSN, WM_COMMAND, 40273, 0
'een audio gesprek starten
'PostMessage MSN_FindMSN, WM_COMMAND, 40274, 0
'een video gesprek starten
'PostMessage MSN_FindMSN, WM_COMMAND, 40288, 0
'contactpersoonlijst importeren
'PostMessage MSN_FindMSN, WM_COMMAND, 40292, 0
'berichtgeschiedenis open..
'PostMessage MSN_FindMSN, WM_COMMAND, 40308, 0
'contactpresonen sorteren op email
'PostMessage MSN_FindMSN, WM_COMMAND, 40307, 0
'contactpersonen sorteren op nickname
