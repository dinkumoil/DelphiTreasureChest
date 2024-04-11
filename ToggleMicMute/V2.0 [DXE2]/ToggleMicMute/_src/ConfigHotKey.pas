{ **************************************************************************** }
//
// Keyboard shortcut configuration dialog for ToggleMicMute.
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

unit ConfigHotKey;


interface


uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, StdCtrls,
  ExtCtrls, Buttons, Forms, Dialogs;


type
  TfrmConfigHotKey = class(TForm)
  private
    FLastKey: Word;

  published
    lblActHotKey       : TLabel;
    pnlEditContainer   : TPanel;
    edtActHotKey       : TEdit;

    sbtnDeleteActHotKey: TSpeedButton;

    lblNewHotKey       : TLabel;
    edtNewHotKey       : TEdit;

    btnOK              : TButton;
    btnAbort           : TButton;
    btnApply           : TButton;

    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);

    procedure edtNewHotKeyKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure edtNewHotKeyKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure edtNewHotKeyKeyPress(Sender: TObject; var Key: Char);

    procedure sbtnDeleteActHotKeyClick(Sender: TObject);

    procedure btnOKClick(Sender: TObject);
    procedure btnAbortClick(Sender: TObject);
    procedure btnApplyClick(Sender: TObject);

  end;


var
  frmConfigHotKey: TfrmConfigHotKey;




implementation

{$R *.dfm}


uses
  Main;



procedure TfrmConfigHotKey.FormCreate(Sender: TObject);
var
  SysMenu: HMenu;

begin
  Caption := frmMain.AppName + Caption;
  Icon    := Application.Icon;

  SysMenu := GetSystemMenu(Handle, FALSE) ;

  DeleteMenu(SysMenu, SC_SIZE, MF_BYCOMMAND) ;
  DeleteMenu(SysMenu, SC_MINIMIZE, MF_BYCOMMAND) ;
  DeleteMenu(SysMenu, SC_MAXIMIZE, MF_BYCOMMAND) ;
  DeleteMenu(SysMenu, SC_RESTORE, MF_BYCOMMAND) ;

  edtActHotKey.Text := frmMain.HotKey;
end;


procedure TfrmConfigHotKey.FormShow(Sender: TObject);
begin
  FLastKey                    := 0;
  edtActHotKey.Text           := frmMain.HotKey;
  edtNewHotKey.Text           := '';
  sbtnDeleteActHotKey.Enabled := (edtActHotKey.Text <> '');
  btnApply.Enabled            := False;
  edtNewHotKey.SetFocus;
end;


procedure TfrmConfigHotKey.edtNewHotKeyKeyDown(Sender: TObject; var Key: Word;
                                               Shift: TShiftState);
var
  ModifiersStr: String;
  KeyStr      : String;
  HotKeyStr   : String;

begin
  if (Key <> FLastKey) then
  begin
    FLastKey := Key;

    objKeyMapper.CheckForWinKeyPressed(Key);
    ModifiersStr := objKeyMapper.MapModifiersToString(Shift);
    KeyStr       := objKeyMapper.MapKeyToString(Key);

    HotKeyStr := ModifiersStr;
    if (ModifiersStr <> '') and (KeyStr <> '') then HotKeyStr := HotKeyStr + '+';
    HotKeyStr := HotKeyStr + KeyStr;

    if (HotKeyStr <> '') then
    begin
			edtNewHotKey.Text     := HotKeyStr;
			edtNewHotKey.SelStart := Length(edtNewHotKey.Text);
			btnApply.Enabled      := True;
    end;
  end;

  Key := 0;
end;


procedure TfrmConfigHotKey.edtNewHotKeyKeyUp(Sender: TObject; var Key: Word;
                                             Shift: TShiftState);
begin
  if (Key = FLastKey) then FLastKey := 0;
  objKeyMapper.CheckForWinKeyReleased(Key);
end;


procedure TfrmConfigHotKey.edtNewHotKeyKeyPress(Sender: TObject; var Key: Char);
begin
  Key := #0;
end;


procedure TfrmConfigHotKey.sbtnDeleteActHotKeyClick(Sender: TObject);
begin
  edtActHotKey.Text           := '';
  sbtnDeleteActHotKey.Enabled := False;
  btnApply.Enabled            := True;
end;


procedure TfrmConfigHotKey.btnOKClick(Sender: TObject);
begin
  if (btnApply.Enabled) then
  begin
    edtActHotKey.Text := edtNewHotKey.Text;
    frmMain.HotKey    := edtActHotKey.Text;
  end;

  Close;
end;


procedure TfrmConfigHotKey.btnAbortClick(Sender: TObject);
begin
  Close;
end;


procedure TfrmConfigHotKey.btnApplyClick(Sender: TObject);
begin
  edtActHotKey.Text           := edtNewHotKey.Text;
  edtNewHotKey.Text           := '';
  sbtnDeleteActHotKey.Enabled := (edtActHotKey.Text <> '');
  btnApply.Enabled            := False;
  frmMain.HotKey              := edtActHotKey.Text;
  edtActHotKey.Text           := frmMain.HotKey;
end;


end.
