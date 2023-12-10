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
// This program is free software; you can redistribute it and/or modify
// it under the terms of the Mozilla Public License Version 2.0.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
//
{ **************************************************************************** }

program SingleAppInstance;

uses
  System.SysUtils,
  Vcl.Forms,
  Main in 'Main.pas' {MainForm},
  class_SingleAppInstance in 'lib\class_SingleAppInstance.pas',
  class_SingleAppInstance.PipeServer in 'lib\class_SingleAppInstance.PipeServer.pas',
  class_SingleAppInstance.PipeClient in 'lib\class_SingleAppInstance.PipeClient.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;

  if TSingleAppInstance.IsAnotherInstanceRunning then
  begin
    TSingleAppInstance.SendCommandLineData();
    exit;
  end;

  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
