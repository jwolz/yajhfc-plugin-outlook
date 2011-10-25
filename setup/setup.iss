#ifndef VERSION
 #define public VERSION "0.5.1"
#endif
#define JACOBDIR "C:\Users\jonas\java\jacob-1.16-M1"

[Files]
Source: ..\build\yajhfc-outlook-pb-plugin.jar; DestDir: {app}; Components: base
Source: ..\dist\README-outlook-plugin.txt; DestDir: {app}; Components: base

Source: {#JACOBDIR}\jacob.jar; DestDir: {app}\lib; Components: base
Source: {#JACOBDIR}\jacob-1.16-M1-x64.dll; DestDir: {app}\lib; Components: base
Source: {#JACOBDIR}\jacob-1.16-M1-x86.dll; DestDir: {app}\lib; Components: base
Source: {#JACOBDIR}\README.txt; DestDir: {app}\lib; DestName: jacob-README.txt; Components: base
Source: {#JACOBDIR}\LICENSE.TXT; DestDir: {app}\lib; DestName: jacob-LICENSE.txt; Components: base


[Setup]
AppCopyright=© 2005-2011 by Jonas Wolz
AppName=YajHFC Outlook Plugin
AppVerName=YajHFC Outlook Plugin {#VERSION}
AppVersion={#VERSION}
;InfoBeforeFile=temp\README.txt
LicenseFile=..\COPYING
DefaultDirName={reg:HKLM\Software\YajHFC,instpath|{pf}\YajHFC}
DefaultGroupName=YajHFC
AppPublisher=Jonas Wolz
AppPublisherURL=http://www.yajhfc.de/
AppID={{087FBDA8-29D7-49F3-8C67-D8221819808B}
UninstallDisplayIcon={app}\yajhfc.ico
UninstallDisplayName=YajHFC Outlook Plugin {#VERSION}
OutputBaseFilename=setup-outlookplugin
ArchitecturesInstallIn64BitMode=x64
DisableDirPage=yes

[Registry]
Root: HKLM; Subkey: Software\YajHFC; ValueType: string; ValueName: addLaunchArg.outlookplugin; ValueData: "--loadplugin=""{app}\yajhfc-outlook-pb-plugin.jar"""; Flags: uninsdeletekeyifempty uninsdeletevalue
Root: HKLM; Subkey: Software\YajHFC; ValueType: string; ValueName: outlookplugin-version; ValueData: {#VERSION}; Flags: uninsdeletekeyifempty uninsdeletevalue

[Components]
Name: Base; Description: All files; Flags: fixed; Types: custom compact full

[Languages]
Name: en; MessagesFile: compiler:Default.isl
Name: de; MessagesFile: compiler:Languages\German.isl 
Name: fr; MessagesFile: compiler:Languages\French.isl
Name: es; MessagesFile: compiler:Languages\Spanish.isl
Name: it; MessagesFile: compiler:Languages\Italian.isl
Name: pl; MessagesFile: compiler:Languages\Polish.isl
Name: ru; MessagesFile: compiler:Languages\Russian.isl
;Name: tr; MessagesFile: compiler:Languages\Turkish.isl

[Code]
procedure splitversion(ver: string; var major, minor, build: integer);
var
  p, p2, i: integer;
  s:string;
begin
    major:=0;
    minor:=0;
    build:=0;

    p := Pos('.', ver);
		if (p > 0) then
		begin
			major := strtointdef(copy(ver, 1, p-1), 0);
      s:= copy(ver, p+1, length(ver)-p);
      p2 := Pos('.', s);
      if (p2 > 0) then
		  begin
			  minor := strtointdef(copy(s, 1, p2-1), 0);
        for i:=p2+1 to length(s) do
        begin
          if (ord(s[i]) < ord('0')) or (ord(s[i]) > ord('9')) then
            break;
        end;
        if (i > p2+1) then
        begin
          build := strtointdef(copy(s, p2+1, i-p2-1), 0);
        end
      end
      else
      begin
        minor := strtointdef(s, 0)
      end;
		end
end;

function InitializeSetup(): Boolean;
var
  yajhfcver: string;
  ymaj, ymin, ybld, pmaj, pmin, pbld: integer;
begin
  if RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\YajHFC', 'version', yajhfcver) then
	begin
    splitversion('{#VERSION}', pmaj, pmin, pbld);
    splitversion(yajhfcver, ymaj, ymin, ybld);
    if (ymaj = pmaj) and (ymin = pmin) and (ybld = pbld) then
    begin
      result := true;
    end
    else
    begin
      case MsgBox(
		  'The installed YajHFC version ' + yajhfcver + ' differs from the plugin version {#VERSION}. This may cause problems.'#13#10#13#10
      'Install the plugin anyway?',
		  mbError, MB_YESNO) of
		IDYES:
			result := true;
		IDNO:
			result := false;
		end;
    end
  end
  else
  begin
    MsgBox('To install the YajHFC Outlook plugin, the YajHFC main application must have been installed before.'#13#10#13#10
           'Please install YajHFC and try again.',
           mbCriticalError, MB_OK);
    result := false;
  end
end;










