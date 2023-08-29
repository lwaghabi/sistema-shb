
[Setup]

#define nomeSistema = "PSHB MS"
#define versaoSistema = "6.0"
AppName={#nomeSistema}
AppVerName={#nomeSistema} - Vs.: {#versaoSistema}
AppVersion={#versaoSistema}

DefaultDirName=C:\SEMIHERMATICS\{#nomeSistema}
DefaultGroupName=SEMIHERMATICS\{#nomeSistema}
;OutputDir=C:\Users\lwagh\OneDrive\Documentos\OneDrive\PSHBCompartilhado\
;OutputDir=C:\Users\lwagh\OneDrive\PSHBCompartilhado
OutputDir=C:\Users\lwagh\OneDrive\Documentos\OneDrive\PSHBCompartilhado
Appcopyright=SEMIHERMATICS DO BRASIL LTDA
OutputBaseFileName = {#nomeSistema}
WindowVisible = no
WindowShowCaption = no
WindowStartMaximized = no
WizardImageFile = C:\Program Files (x86)\HcSoftware\HcMaker 3.0\MODELOS\MICROSOFTMODERN01.BMP
;WizardSmallImageFile = 
AllowNoIcons = True
ShowLanguageDialog = no
;SetupIconFile=
UninstallDisplayIcon={app}\PROJETO SHB MS 2.2.EXE
UninstallDisplayName={#nomeSistema}
VersionInfoVersion=1.0.0
VersionInfoCompany=SEMIHERMATICS DO BRASIL LTDA
VersionInfoDescription={#nomeSistema}
;UserInfoPage = False
DisableDirPage = True
DisableProgramGroupPage = True
AllowUNCPath = False
AppendDefaultDirName = False
AlwaysShowDirOnReadyPage = True
AlwaysShowGroupOnReadyPage = True
ShowTasksTreeLines = True
InternalCompressLevel = Max
SolidCompression = True
VersionInfoTextVersion={#versaoSistema}
AlwaysShowComponentsList = True
AppPublisherURL = '
VersionInfoCopyright=© SEMIHERMATICS DO BRASIL LTDA

;AppSupportPhone=
;IDIOMAS DE INSTALAÇÃO
[Languages]
Name:po; MessagesFile: C:\Program Files (x86)\HcSoftware\HcMaker 3.0\Languages\BrazilianPortuguese.isl
;COMPONENTES E ARQUIVOS DO SISTEMA
[Files]
;Source: C:\MEUS DOCUMENTOS\SISTEMA SHB\PROJETO SHB MS\PROJETO SHB MS 2.2.EXE; DestDir: {app}; Flags: ignoreversion restartreplace allowunsafefiles
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\PSHBCompartilhado\PROJETO SHB MS 2.2.EXE; DestDir: {app}; Flags: ignoreversion restartreplace allowunsafefiles
Source: C:\Program Files (x86)\HcSoftware\HcMaker 3.0\Skins\ISSkin.dll; DestDir: {app}; Flags: dontcopy
Source: C:\Program Files (x86)\HcSoftware\HcMaker 3.0\Skins\Styles\WinXP.Luna.cjstyles; DestDir: {tmp}; Flags: dontcopy
Source: C:\WINDOWS\SYSTEM32\MSMASK32.OCX; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\WINDOWS\SYSTEM32\TABCTL32.OCX; DestDir: {sys}; Flags: sharedfile regserver noregerror
Source: C:\WINDOWS\SYSTEM32\MSFLXGRD.OCX; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\WINDOWS\SYSTEM32\MSCOMCT2.OCX; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\WINDOWS\SYSTEM32\COMCTL32.OCX; DestDir: {sys}; Flags: restartreplace allowunsafefiles regserver sharedfile uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror
Source: C:\Program Files (x86)\HcSoftware\HcMaker 3.0\VBRunTime\MSVBVM60.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace allowunsafefiles regserver sharedfile uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror
Source: C:\Program Files (x86)\HcSoftware\HcMaker 3.0\VBRunTime\OLEAUT32.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace allowunsafefiles regserver sharedfile uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror
Source: C:\Program Files (x86)\HcSoftware\HcMaker 3.0\VBRunTime\OLEPRO32.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace allowunsafefiles regserver sharedfile uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror
Source: C:\Program Files (x86)\HcSoftware\HcMaker 3.0\VBRunTime\ASYCFILT.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace allowunsafefiles regserver sharedfile uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror
Source: C:\Program Files (x86)\HcSoftware\HcMaker 3.0\VBRunTime\COMCAT.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace allowunsafefiles regserver sharedfile uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror
;REFERÊNCIAS DO PROJETO Adicionado pelo professor Mário Andrade em 14/09/2021

Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\MSCOMCTL.OCX; DestDir: {sys}; Flags: sharedfile regserver noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\MSCOMCT2.OCX; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\COMDLG32.OCX; DestDir: {sys}; Flags: restartreplace allowunsafefiles regserver sharedfile uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\GLXPBUTTONZ.OCX; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\CHAMALEONBTN.OCX; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\MSVBVM60.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace allowunsafefiles regserver sharedfile uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\OLEAUT32.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace allowunsafefiles regserver sharedfile uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\OLEPRO32.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace allowunsafefiles regserver sharedfile uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\ASYCFILT.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace allowunsafefiles regserver sharedfile uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\COMCAT.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace allowunsafefiles regserver sharedfile uninsneveruninstall onlyifdoesntexist allowunsafefiles noregerror

Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\TABCTL32.OCX; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\CoolXPFrame.ocx; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\CoolXPText.ocx; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\CoolXPCheck.ocx; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\STDOLE2.TLB; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\MSBIND.DLL; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\MIGV2.DLL; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\MSDBRPTR.DLL; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\MSSTDFMT.DLL; DestDir: {sys}; Flags: sharedfile regserver noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\ARPRO2.DLL; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\MSADO60_BACKCOMPAT.TLB; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\SCRRUN.DLL; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror

;Para o atualizador
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\MSINET.OCX; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror

;para Microsoft windows form 2.0 obrigatório
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\FM20.DLL; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\FM20.OCA; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\FM20PTB.DLL; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\FM20ENU.DLL; DestDir: {sys}; Flags: allowunsafefiles sharedfile regserver uninsneveruninstall noregerror 

Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\vcredist.exe; DestDir: {tmp}\; Flags: onlyifdoesntexist
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\mysql-connector-odbc-5.2.7-win32.MSI; DestDir: {tmp}\; Flags: onlyifdoesntexist
Source: C:\Users\lwagh\OneDrive\Documentos\OneDrive\Sistema\Dll\mysql-connector-odbc-3.51.30-win32.MSI; DestDir: {tmp}\; Flags: onlyifdoesntexist
;<========================================================================================================================== 

;ARQUIVOS ADICIONAIS
;ICONES DO PROJETO
[Icons]
Name: {group}\{#nomeSistema}; Filename: {app}\PROJETO SHB MS 2.2.EXE; WorkingDir: {app} IconFilename: {app}\PROJETO SHB MS 2.2.EXE; IconIndex: 0
Name: {userdesktop}\{#nomeSistema}; Filename: {app}\PROJETO SHB MS 2.2.EXE; WorkingDir: {app}; IconIndex: 0; Tasks: desktopicon
Name: {userappdata}\Microsoft\Internet Explorer\Quick Launch\{#nomeSistema}; Filename: {app}\PROJETO SHB MS 2.2.EXE; WorkingDir: {app}; IconIndex: 0; Tasks: quicklaunchicon; MinVersion: 4.1.1998,4
Name: {group}\{cm:UninstallProgram,{#nomeSistema}}; Filename: {uninstallexe}
;EXECUÇÃO
[Run]
Filename: {app}\PROJETO SHB MS 2.2.EXE; Description: "Executar: {#nomeSistema}"; Flags: postinstall nowait skipifsilent unchecked
Filename: {tmp}\vcredist.exe; Description: Instalando Arquivos Adicionais; Check: not fwExists 
Filename: msiexec.exe; Parameters: "/i ""{tmp}\mysql-connector-odbc-3.51.30-win32.msi"""; Check: not msq3Exists
Filename: msiexec.exe; Parameters: "/i ""{tmp}\mysql-connector-odbc-5.2.7-win32.msi"""; Check: not msq5Exists
;TASKS
[Tasks]
Name: desktopicon; Description: Criar um icone na Área de Trabalho ( &desktop ); GroupDescription: Atalhos adicionais:
Name: quicklaunchicon; Description: Criar um ícone na Barra Quick Lanch; GroupDescription: Atalhos adicionais:; MinVersion: 4.1.1998,4

;ARQUIVO INI
[INI]
Filename: {app}\Instalação.ini; Section: {#nomeSistema}; Key: Compilado; String: 14/09/2021
Filename: {app}\Instalação.ini; Section: {#nomeSistema}; Key: Versão; String: {#versaoSistema}
Filename: {app}\Instalação.ini; Section: {#nomeSistema}; Key: Sistema; String: {#versaoSistema}
;CÓDIGO PASCAL
[Code]
procedure LoadSkin(lpszPath: String; lpszIniFileName: String);
external 'LoadSkin@files:isskin.dll stdcall';
procedure UnloadSkin();
external 'UnloadSkin@files:isskin.dll stdcall';
function ShowWindow(hWnd: Integer; uType: Integer): Integer;
external 'ShowWindow@user32.dll stdcall';
function InitializeSetup(): Boolean;
begin
  ExtractTemporaryFile('WinXP.Luna.cjstyles');
  LoadSkin(ExpandConstant('{tmp}\WinXP.Luna.cjstyles'), '');
  Result := True;
end;
function fwExists(): Boolean;
  begin
  if RegKeyExists(HKLM,'SOFTWARE\Classes\Installer\Products\D04BB691875110D32B98EBCF771AA1E1\') = false then
  begin
     Result :=   false;      
  end else
  begin       
      Result := true;      
  end;
end;
function msq3Exists(): Boolean;
  begin
  if RegKeyExists(HKLM,'SOFTWARE\Wow6432Node\MySQL AB\MySQL Connector/ODBC 3.51') = false then
  begin
     Result :=   false;
     if RegKeyExists(HKLM,'SOFTWARE\MySQL AB\MySQL Connector/ODBC 3.51') = true then
     begin
       Result :=   true;
     end;         
  end else
  begin
      Result := true;      
  end;
end;
function msq5Exists(): Boolean;
  begin
  if RegKeyExists(HKLM,'SOFTWARE\Wow6432Node\MySQL AB\MySQL Connector/ODBC 5.2') = false then
  begin
     Result :=   false; 
     if RegKeyExists(HKLM,'SOFTWARE\MySQL AB\MySQL Connector/ODBC 5.2') = true then
     begin
       Result :=   true;
     end;    
  end else
  begin
      Result := true;      
  end;
end;
procedure DeinitializeSetup();
begin
  ShowWindow(StrToInt(ExpandConstant('{wizardhwnd}')), 0);
  UnloadSkin();
end;

