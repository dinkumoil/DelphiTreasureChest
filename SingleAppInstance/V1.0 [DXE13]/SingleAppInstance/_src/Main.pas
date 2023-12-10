{ **************************************************************************** }
//
// Demo code for single app instance component.
//
// Written in Delphi 10.4 Sydney.
//
// Target OS: MS Windows (Vista and later)
// Author   : Andreas Heim, 2020-06
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

unit Main;


interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Math, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.StdCtrls, Vcl.Forms, Vcl.Dialogs,

  class_SingleAppInstance;


type
  TMainForm = class(TForm)
    btnClose: TButton;
    memOutput: TMemo;

    procedure FormCreate(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);

  private
    procedure OtherInstanceSentData(const Data: string);

  end;


var
  MainForm: TMainForm;



implementation

{$R *.dfm}


procedure TMainForm.FormCreate(Sender: TObject);
begin
  if TSingleAppInstance.InstanceState = aisMainInstance then
  begin
    TSingleAppInstance.OnOtherInstanceSentData := OtherInstanceSentData;
    TSingleAppInstance.AcceptIncomingData      := true;
  end;
end;


procedure TMainForm.OtherInstanceSentData(const Data: string);
begin
  memOutput.Lines.Add(Data);
end;


procedure TMainForm.btnCloseClick(Sender: TObject);
begin
  Close;
end;


end.
