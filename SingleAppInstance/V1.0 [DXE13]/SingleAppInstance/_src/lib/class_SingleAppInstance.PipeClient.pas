{ **************************************************************************** }
//
// Named pipe client. Intended to be used in conjunction with single app
// instance component in unit class_SingleAppInstance.pas.
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

unit class_SingleAppInstance.PipeClient;


interface

uses
  Winapi.Windows, System.SysUtils, System.Types, System.Classes;


type
  TPipeClient = class(TObject)
  private
    FPipe:              THandle;
    FMainInstWndHandle: THandle;

  public
    constructor Create;
    destructor  Destroy; override;

    function    ConnectToServer(const APipeName: string): boolean;
    function    TalkToServer: boolean;
    procedure   CloseConnection;

    property    MainInstanceWndHandle: THandle read FMainInstWndHandle;

  end;



implementation


constructor TPipeClient.Create;
begin
  inherited;

  FPipe              := INVALID_HANDLE_VALUE;
  FMainInstWndHandle := 0;
end;


destructor TPipeClient.Destroy;
begin
  if FPipe <> INVALID_HANDLE_VALUE then
    CloseHandle(FPipe);

  inherited;
end;


function TPipeClient.ConnectToServer(const APipeName: string): boolean;
var
  PipeReadMode: DWORD;
  IsSuccess:    BOOL;
  PipeName:     string;

begin
  PipeName := '\\.\pipe\' + APipeName;

  // Try to open a named pipe; wait for it, if necessary.
  while true do
  begin
    FPipe := CreateFile(PChar(PipeName),      // pipe name
                        GENERIC_READ          // read and write access
                          or GENERIC_WRITE,
                        0,                    // no sharing
                        nil,                  // default security attributes
                        OPEN_EXISTING,        // opens existing pipe
                        0,                    // default attributes
                        0);                   // no template file

    // Break if the pipe handle is valid.
    if FPipe <> INVALID_HANDLE_VALUE then
      break;

    // Exit if an error other than ERROR_PIPE_BUSY occurs.
    if GetLastError() <> ERROR_PIPE_BUSY then
      exit(false);

    // All pipe instances are busy, so wait some time.
    if not WaitNamedPipe(PChar(PipeName), NMPWAIT_USE_DEFAULT_WAIT) then
      exit(false);
  end;

  // The pipe connected; change to message-read mode.
  PipeReadMode := PIPE_READMODE_MESSAGE;
  IsSuccess    := SetNamedPipeHandleState(FPipe,         // pipe handle
                                          PipeReadMode,  // new pipe mode
                                          nil,           // don't set maximum bytes
                                          nil);          // don't set maximum time

  if not IsSuccess then
    CloseConnection;

  Result := IsSuccess;
end;


function TPipeClient.TalkToServer: boolean;
var
  Data:         string;
  Cnt:          integer;
  IsSuccess:    boolean;
  BytesRead:    DWORD;
  BytesToWrite: DWORD;
  ByesWritten:  DWORD;

begin
  if ParamCount = 0 then exit(true);

  // ...........................................................................
  // Read window handle of main instance from the pipe.
  // ...........................................................................

  BytesRead := 0;

  IsSuccess := ReadFile(FPipe,               // pipe handle
                        FMainInstWndHandle,  // buffer to receive reply
                        SizeOf(THandle),     // size of buffer
                        BytesRead,           // number of bytes read
                        nil);                // not overlapped

  if not IsSuccess then  // This terminates communication if there was an error
  begin                  // or if the server wants to send too much data
    CloseConnection;
    exit(false);
  end;

  // ...........................................................................
  // Write command line of other instance to the pipe
  // ...........................................................................

  Data := AnsiQuotedStr(ParamStr(1), '"');

  // Concat application's command line parameters
  for Cnt := 2 to ParamCount do
    Data := Format('%s "%s"', [Data, ParamStr(Cnt)]);

  // Send message to the pipe server.
  BytesToWrite := Length(Data) * SizeOf(char);
  ByesWritten  := 0;

  IsSuccess := WriteFile(FPipe,         // pipe handle
                         Data[1],       // message
                         BytesToWrite,  // message length
                         ByesWritten,   // bytes written
                         nil);          // not overlapped

  if not IsSuccess then
  begin
    CloseConnection;
    exit(false);
  end;

  Result := IsSuccess;
end;


procedure TPipeClient.CloseConnection;
begin
   CloseHandle(FPipe);
   FPipe := INVALID_HANDLE_VALUE;
end;


end.
