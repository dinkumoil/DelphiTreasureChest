{ **************************************************************************** }
//
// ToggleMicMute, a Windows system tray application to mute and unmute your
// microphone by clicking its icon or using a configurable keyboard shortcut.
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

program ToggleMicMute;

uses
  Forms,
  Main in 'Main.pas' {frmMain},
  ConfigHotKey in 'ConfigHotKey.pas' {frmConfigHotKey},
  KeyMapper in 'KeyMapper.pas',
  CoreAudio in 'lib\CoreAudio.pas',
  PropSys2 in 'lib\PropSys2.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.ShowMainForm := False;
  Application.MainFormOnTaskbar := False;
  Application.Title := 'ToggleMicMute';
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TfrmConfigHotKey, frmConfigHotKey);
  Application.Run;
end.
