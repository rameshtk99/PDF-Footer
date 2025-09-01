; -- Inno Setup Script for PdfFooterAutomation --

[Setup]
AppName=PDF Footer Automation
AppVersion=1.0
DefaultDirName={pf}\PdfFooterAutomation
DefaultGroupName=PdfFooterAutomation
OutputBaseFilename=PdfFooterAutomation_Setup
Compression=lzma
SolidCompression=yes
SetupIconFile=logo.ico  
WizardImageFile=logo.bmp
WizardSmallImageFile=logo_small.bmp



[Files]
Source: "dist\PdfFooterAutomation.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "preeti.ttf"; DestDir: "{app}"; Flags: ignoreversion
Source: "logo.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\PDF Footer Automation"; Filename: "{app}\PdfFooterAutomation.exe"; IconFilename: "{app}\logo.ico"
Name: "{commondesktop}\PDF Footer Automation"; Filename: "{app}\PdfFooterAutomation.exe"; IconFilename: "{app}\logo.ico"

[Run]
; Install font (Preeti.ttf)
Filename: "cmd"; Parameters: "/c copy ""{app}\preeti.ttf"" ""%WINDIR%\Fonts"""; Flags: runhidden
; Register font (not always needed, but safe)
Filename: "regsvr32"; Parameters: "/s ""{app}\preeti.ttf"""; Flags: runhidden

; Launch app after install
Filename: "{app}\PdfFooterAutomation.exe"; Description: "Launch PDF Footer Automation"; Flags: nowait postinstall skipifsilent
