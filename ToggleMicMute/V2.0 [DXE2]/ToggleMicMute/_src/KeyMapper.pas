{ **************************************************************************** }
//
// Mapping of virtual key codes to key names.
//
// This code is part of ToggleMicMute, a Windows system tray application to mute
// and unmute your microphone by clicking its icon or using a keyboard shortcut.
//
// Written in Delphi XE2.
//
// Target OS: MS Windows (Vista and later)
// Author   : Andreas Heim, 2011
//
//
// This program is free software; you can redistribute it and/or modify it
// under the terms of the GNU General Public License version 3 as published
// by the Free Software Foundation.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License along
// with this program; if not, write to the Free Software Foundation, Inc.,
// 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
//
{ **************************************************************************** }

unit KeyMapper;

interface

uses
  Windows, Classes, StrUtils;


type
  TKeyMapper = class(TObject)
  private
    FModifiers    : Array [0..3]   of Cardinal;
    FModifiersStr : Array [0..3]   of String;

    FVK           : Array [0..126] of Word;
    FVKStr        : Array [0..126] of String;

    FWinKeyPressed: Boolean;

    function    CheckForWinKey(var Key: Word): Boolean;

    function    MapModStrToModifier(AModStr: String): Cardinal;
    function    MapKeyStrToKey(AKeyStr: String): Cardinal;

  public
    constructor Create;

    procedure   CheckForWinKeyPressed(var Key: Word);
    procedure   CheckForWinKeyReleased(var Key: Word);

    function    MapModifiersToString(AShift: TShiftState): String;
    function    MapKeyToString(AKey: Word): String;

    procedure   ParseHotKeyStr(AHotKeyStr: String; var Modifiers, Key: Cardinal);

  end;



implementation


constructor TKeyMapper.Create;
begin
  inherited;

  FWinKeyPressed := False;

  FModifiers[0]    := MOD_SHIFT;
  FModifiers[1]    := MOD_CONTROL;
  FModifiers[2]    := MOD_ALT;
  FModifiers[3]    := MOD_WIN;

  FModifiersStr[0] := 'SHIFT';
  FModifiersStr[1] := 'CTRL';
  FModifiersStr[2] := 'ALT';
  FModifiersStr[3] := 'WIN';

  FVK[0]   := 48; //0
  FVK[1]   := 49;
  FVK[2]   := 50;
  FVK[3]   := 51;
  FVK[4]   := 52; //bis
  FVK[5]   := 53;
  FVK[6]   := 54;
  FVK[7]   := 55;
  FVK[8]   := 56;
  FVK[9]   := 57; //9
  FVK[10]  := 65; //A
  FVK[11]  := 66;
  FVK[12]  := 67;
  FVK[13]  := 68;
  FVK[14]  := 69;
  FVK[15]  := 70;
  FVK[16]  := 71;
  FVK[17]  := 72;
  FVK[18]  := 73;
  FVK[19]  := 74;
  FVK[20]  := 75;
  FVK[21]  := 76;
  FVK[22]  := 77; //bis
  FVK[23]  := 78;
  FVK[24]  := 79;
  FVK[25]  := 80;
  FVK[26]  := 81;
  FVK[27]  := 82;
  FVK[28]  := 83;
  FVK[29]  := 84;
  FVK[30]  := 85;
  FVK[31]  := 86;
  FVK[32]  := 87;
  FVK[33]  := 88;
  FVK[34]  := 89;
  FVK[35]  := 90; //Z
  FVK[36]  := VK_CANCEL;
  FVK[37]  := VK_BACK; //Backspace
  FVK[38]  := VK_TAB;
  FVK[39]  := VK_CLEAR;
  FVK[40]  := VK_RETURN;
  FVK[41]  := VK_SHIFT;
  FVK[42]  := VK_CONTROL;
  FVK[43]  := VK_MENU; //ALT
  FVK[44]  := VK_PAUSE;
  FVK[45]  := VK_CAPITAL; //CapsLock
  FVK[46]  := VK_FINAL;
  FVK[47]  := VK_CONVERT;
  FVK[48]  := VK_NONCONVERT;
  FVK[49]  := VK_ACCEPT;
  FVK[50]  := VK_MODECHANGE;
  FVK[51]  := VK_ESCAPE;
  FVK[52]  := VK_SPACE;
  FVK[53]  := VK_PRIOR; //Bild Auf
  FVK[54]  := VK_NEXT; //Bild Ab
  FVK[55]  := VK_END;
  FVK[56]  := VK_HOME;
  FVK[57]  := VK_LEFT;
  FVK[58]  := VK_UP;
  FVK[59]  := VK_RIGHT;
  FVK[60]  := VK_DOWN;
  FVK[61]  := VK_SELECT;
  FVK[62]  := VK_PRINT;
  FVK[63]  := VK_EXECUTE;
  FVK[64]  := VK_SNAPSHOT;
  FVK[65]  := VK_INSERT;
  FVK[66]  := VK_DELETE;
  FVK[67]  := VK_HELP;
  FVK[68]  := VK_LWIN;
  FVK[69]  := VK_RWIN;
  FVK[70]  := VK_APPS;
  FVK[71]  := VK_SLEEP;
  FVK[72]  := VK_NUMPAD0;
  FVK[73]  := VK_NUMPAD1;
  FVK[74]  := VK_NUMPAD2;
  FVK[75]  := VK_NUMPAD3;
  FVK[76]  := VK_NUMPAD4;
  FVK[77]  := VK_NUMPAD5;
  FVK[78]  := VK_NUMPAD6;
  FVK[79]  := VK_NUMPAD7;
  FVK[80]  := VK_NUMPAD8;
  FVK[81]  := VK_NUMPAD9;
  FVK[82]  := VK_MULTIPLY;
  FVK[83]  := VK_ADD;
  FVK[84]  := VK_SEPARATOR;
  FVK[85]  := VK_SUBTRACT;
  FVK[86]  := VK_DECIMAL;
  FVK[87]  := VK_DIVIDE;
  FVK[88]  := VK_F1;
  FVK[89]  := VK_F2;
  FVK[90]  := VK_F3;
  FVK[91]  := VK_F4;
  FVK[92]  := VK_F5;
  FVK[93]  := VK_F6;
  FVK[94]  := VK_F7;
  FVK[95]  := VK_F8;
  FVK[96]  := VK_F9;
  FVK[97]  := VK_F10;
  FVK[98]  := VK_F11;
  FVK[99]  := VK_F12;
  FVK[100] := VK_F13;
  FVK[101] := VK_F14;
  FVK[102] := VK_F15;
  FVK[103] := VK_F16;
  FVK[104] := VK_F17;
  FVK[105] := VK_F18;
  FVK[106] := VK_F19;
  FVK[107] := VK_F20;
  FVK[108] := VK_F21;
  FVK[109] := VK_F22;
  FVK[110] := VK_F23;
  FVK[111] := VK_F24;
  FVK[112] := VK_NUMLOCK;
  FVK[113] := VK_SCROLL;
  FVK[114] := VK_OEM_1;
  FVK[115] := VK_OEM_PLUS;
  FVK[116] := VK_OEM_COMMA;
  FVK[117] := VK_OEM_MINUS;
  FVK[118] := VK_OEM_PERIOD;
  FVK[119] := VK_OEM_2;
  FVK[120] := VK_OEM_3;
  FVK[121] := VK_OEM_4;
  FVK[122] := VK_OEM_5;
  FVK[123] := VK_OEM_6;
  FVK[124] := VK_OEM_7;
  FVK[125] := VK_OEM_8;
  FVK[126] := VK_OEM_102;

  FVKStr[0]   := '0';
  FVKStr[1]   := '1';
  FVKStr[2]   := '2';
  FVKStr[3]   := '3';
  FVKStr[4]   := '4';
  FVKStr[5]   := '5';
  FVKStr[6]   := '6';
  FVKStr[7]   := '7';
  FVKStr[8]   := '8';
  FVKStr[9]   := '9';
  FVKStr[10]  := 'A';
  FVKStr[11]  := 'B';
  FVKStr[12]  := 'C';
  FVKStr[13]  := 'D';
  FVKStr[14]  := 'E';
  FVKStr[15]  := 'F';
  FVKStr[16]  := 'G';
  FVKStr[17]  := 'H';
  FVKStr[18]  := 'I';
  FVKStr[19]  := 'J';
  FVKStr[20]  := 'K';
  FVKStr[21]  := 'L';
  FVKStr[22]  := 'M';
  FVKStr[23]  := 'N';
  FVKStr[24]  := 'O';
  FVKStr[25]  := 'P';
  FVKStr[26]  := 'Q';
  FVKStr[27]  := 'R';
  FVKStr[28]  := 'S';
  FVKStr[29]  := 'T';
  FVKStr[30]  := 'U';
  FVKStr[31]  := 'V';
  FVKStr[32]  := 'W';
  FVKStr[33]  := 'X';
  FVKStr[34]  := 'Y';
  FVKStr[35]  := 'Z';
  FVKStr[36]  := '';
  FVKStr[37]  := 'BACKSPACE';
  FVKStr[38]  := 'TAB';
  FVKStr[39]  := '';
  FVKStr[40]  := 'ENTER';
  FVKStr[41]  := ''; //SHIFT
  FVKStr[42]  := ''; //CTRL
  FVKStr[43]  := ''; //ALT
  FVKStr[44]  := 'PAUSE';
  FVKStr[45]  := 'CAPSLOCK';
  FVKStr[46]  := '';
  FVKStr[47]  := '';
  FVKStr[48]  := '';
  FVKStr[49]  := '';
  FVKStr[50]  := '';
  FVKStr[51]  := 'ESC';
  FVKStr[52]  := 'LEER';
  FVKStr[53]  := 'BILD AUF';
  FVKStr[54]  := 'BILD AB';
  FVKStr[55]  := 'ENDE';
  FVKStr[56]  := 'POS1';
  FVKStr[57]  := 'LINKS';
  FVKStr[58]  := 'AUF';
  FVKStr[59]  := 'RECHTS';
  FVKStr[60]  := 'AB';
  FVKStr[61]  := '';
  FVKStr[62]  := '';
  FVKStr[63]  := '';
  FVKStr[64]  := '';
  FVKStr[65]  := 'EINFG';
  FVKStr[66]  := 'ENTF';
  FVKStr[67]  := '';
  FVKStr[68]  := ''; //LWIN
  FVKStr[69]  := ''; //RWIN
  FVKStr[70]  := 'MENÜ';
  FVKStr[71]  := '';
  FVKStr[72]  := 'NUM0';
  FVKStr[73]  := 'NUM1';
  FVKStr[74]  := 'NUM2';
  FVKStr[75]  := 'NUM3';
  FVKStr[76]  := 'NUM4';
  FVKStr[77]  := 'NUM5';
  FVKStr[78]  := 'NUM6';
  FVKStr[79]  := 'NUM7';
  FVKStr[80]  := 'NUM8';
  FVKStr[81]  := 'NUM9';
  FVKStr[82]  := 'NUM*';
  FVKStr[83]  := 'NUM+';
  FVKStr[84]  := '';
  FVKStr[85]  := 'NUM-';
  FVKStr[86]  := 'NUM,';
  FVKStr[87]  := 'NUM/';
  FVKStr[88]  := 'F1';
  FVKStr[89]  := 'F2';
  FVKStr[90]  := 'F3';
  FVKStr[91]  := 'F4';
  FVKStr[92]  := 'F5';
  FVKStr[93]  := 'F6';
  FVKStr[94]  := 'F7';
  FVKStr[95]  := 'F8';
  FVKStr[96]  := 'F9';
  FVKStr[97]  := 'F10';
  FVKStr[98]  := 'F11';
  FVKStr[99]  := ''; //F12, Reserviert für Kernel-Debugger
  FVKStr[100] := 'F13';
  FVKStr[101] := 'F14';
  FVKStr[102] := 'F15';
  FVKStr[103] := 'F16';
  FVKStr[104] := 'F17';
  FVKStr[105] := 'F18';
  FVKStr[106] := 'F19';
  FVKStr[107] := 'F20';
  FVKStr[108] := 'F21';
  FVKStr[109] := 'F22';
  FVKStr[110] := 'F23';
  FVKStr[111] := 'F24';
  FVKStr[112] := 'NUMLOCK';
  FVKStr[113] := 'ROLLEN';
  FVKStr[114] := 'Ü';
  FVKStr[115] := '+';
  FVKStr[116] := ',';
  FVKStr[117] := '-';
  FVKStr[118] := '.';
  FVKStr[119] := '#';
  FVKStr[120] := 'Ö';
  FVKStr[121] := 'ß';
  FVKStr[122] := '^';
  FVKStr[123] := '´';
  FVKStr[124] := 'Ä';
  FVKStr[125] := '';
  FVKStr[126] := '<';
end;


function TKeyMapper.MapModifiersToString(AShift: TShiftState): String;
begin
  Result := '';

  if ssShift in AShift then
    Result := FModifiersStr[0];

  if ssCtrl in AShift then
  begin
    if Result <> '' then Result := Result + '+';
    Result := Result + FModifiersStr[1];
  end;

  if ssAlt in AShift then
  begin
    if Result <> '' then Result := Result + '+';
    Result := Result + FModifiersStr[2];
  end;

  if FWinKeyPressed then
  begin
    if Result <> '' then Result := Result + '+';
    Result := Result + FModifiersStr[3];
  end;
end;


function TKeyMapper.MapKeyToString(AKey: Word): String;
var
  I: Integer;

begin
  Result := '';

  for I := 0 to 126 do
    if (FVK[I] = AKey) then
    begin
      Result := FVKStr[I];
      break;
    end;
end;


procedure TKeyMapper.ParseHotKeyStr(AHotKeyStr: String; var Modifiers: Cardinal; var Key: Cardinal);
var
  Part    : String;
  Start,
  DelimIdx: Integer;
  Modifier: Cardinal;

begin
  Modifiers := 0;
  Key       := 0;

  if (Length(AHotKeyStr) > 0) then
  begin
    Start    := 1;
    DelimIdx := PosEx('+', AHotKeyStr, Start);

    if (DelimIdx = 0) and (Start <= Length(AHotKeyStr)) then
      DelimIdx := Length(AHotKeyStr) + 1;

    while (DelimIdx > 0) do
    begin
      Part     := MidStr(AHotKeyStr, Start, DelimIdx - Start);
      Modifier := MapModStrToModifier(Part);

      if (Modifier <> 0) then
        Inc(Modifiers, Modifier)
      else
        Key := MapKeyStrToKey(Part);

      Start    := DelimIdx + 1;
      DelimIdx := PosEx('+', AHotKeyStr, Start);

      if (DelimIdx = 0) and (Start <= Length(AHotKeyStr)) then
        DelimIdx := Length(AHotKeyStr) + 1;
    end;
  end;
end;


function TKeyMapper.MapModStrToModifier(AModStr: String): Cardinal;
var
  I: Integer;

begin
  Result := 0;

  for I := 0 to 3 do
    if (FModifiersStr[I] = AModStr) then
    begin
      Result := FModifiers[I];
      break;
    end;
end;


function TKeyMapper.MapKeyStrToKey(AKeyStr: String): Cardinal;
var
  I: Integer;

begin
  Result := 0;

  for I := 0 to 126 do
    if (FVKStr[I] = AKeyStr) then
    begin
      Result := FVK[I];
      break;
    end;
end;


procedure TKeyMapper.CheckForWinKeyPressed(var Key: Word);
begin
  if not FWinKeyPressed then
    FWinKeyPressed := CheckForWinKey(Key);
end;


procedure TKeyMapper.CheckForWinKeyReleased(var Key: Word);
begin
  if FWinKeyPressed then
    FWinKeyPressed := not CheckForWinKey(Key);
end;


function TKeyMapper.CheckForWinKey(var Key: Word): Boolean;
begin
  Result := False;

  if (Key = VK_LWIN) or (Key = VK_RWIN) then
  begin
    Result := True;
    Key    := 0;
  end;
end;


end.
