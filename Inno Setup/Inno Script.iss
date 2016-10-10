;---------------- SL Satellite Setup -----------------
;- Copyright indaGalaxy.co.uk, Aaron (DAngel) Layton -
;-----------------------------------------------------

[Setup]
AppName=SL Satellite v2
AppVerName=SL Satellite v2
DefaultDirName={pf}\indaGalaxy\SL Satellite v2
DefaultGroupName=SL Satellite v2
UninstallDisplayIcon={app}\Satellite.exe

[Files]
Source: "Satellite.exe"; DestDir: "{app}"
Source: "Satellite.htm"; DestDir: "c:\"
Source: "overlay.wav"; DestDir: "{app}"
Source: "popup.wav"; DestDir: "{app}"
Source: "Readme.txt"; DestDir: "{app}\News"
Source: "Skylords.com.url"; DestDir: "{app}"
Source: "indaGalaxy.co.uk.url"; DestDir: "{app}"
Source: "License Readme.txt"; DestDir: "{app}"; Flags: isreadme
Source: "comdlg32.ocx"; DestDir: "{app}"
Source: "mscomct2.ocx"; DestDir: "{app}"
Source: "scrrun.dll"; DestDir: "{app}"
Source: "tabctl32.ocx"; DestDir: "{app}"

[Icons]
Name: "{group}\SL Satellite v2"; Filename: "{app}\Satellite.exe"
Name: "{group}\Skylords.com"; Filename: "{app}\Skylords.com.url"
Name: "{group}\indaGalaxy.co.uk"; Filename: "{app}\indaGalaxy.co.uk.url"

[Registry]
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\indaGalaxy"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "SMTP"; ValueData: "mail.hotpop.com:25"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "EMail"; ValueData: "You@yourplace.com"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "sMail"; ValueData: "False"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "wList"; ValueData: "SP Confederation"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "sPath"; ValueData: "{app}\News"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "cUpdates"; ValueData: "True"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "Username"; ValueData: "Your Username"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "fAlerts"; ValueData: "False"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "Refresh"; ValueData: "8"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "Sounds"; ValueData: "True"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "Overlay"; ValueData: "True"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "Popup"; ValueData: "True"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "rAlerts"; ValueData: "True"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "rAll"; ValueData: "False"
Root: HKCU; Subkey: "Software\indaGalaxy\SL Satellite"; ValueType: string; ValueName: "Offset"; ValueData: "0"
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "SL Satellite"; ValueData: """{app}\Satellite.exe"" /hide"
Root: HKCU; Subkey: "Software\Microsoft\Internet Explorer\MenuExt\Send to SL Satellite"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Internet Explorer\MenuExt\Send to SL Satellite"; ValueType: string; ValueName: ""; ValueData: "c:\Satellite.htm"
Root: HKCU; Subkey: "Software\Microsoft\Internet Explorer\MenuExt\Send to SL Satellite"; ValueType: dword; ValueName: "Contexts"; ValueData: 48

