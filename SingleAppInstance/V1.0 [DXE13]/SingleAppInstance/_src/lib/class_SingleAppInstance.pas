{ **************************************************************************** }
//
// Single app instance component. This component is intended to be used
// as a singleton class to be able to implement single instance mode for
// an arbitrary application.
//
// It only implements the control logic for this purpose. The necessary
// communication between the main application instance and subsequently
// started instances in order to exchange command line arguments is done
// via a client-server architecture using named pipes. The required units
// are listed below.
//
// Written in Delphi 10.4 Sydney.
//
// Target OS: MS Windows (Vista and later)
// Author   : Andreas Heim, 2020-06
//
//
// This component requires the following units to function properly:
//
//   - class_SingleAppInstance.PipeServer.pas
//   - class_SingleAppInstance.PipeClient.pas
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

unit class_SingleAppInstance;


interface

uses
  Winapi.Windows, System.SysUtils, System.Types, System.Classes, Vcl.Forms,
  Vcl.Dialogs,

  class_SingleAppInstance.PipeServer,
  class_SingleAppInstance.PipeClient;


type
  // App instance states
  TAppInstanceState = (
    aisNone,
    aisUnknown,
    aisMainInstance,
    aisOtherInstance
  );


  // Event type for handler of incoming data of subsequently started app instance
  TOtherInstanceSentDataEvent = procedure(const Data: string) of object;


  // This is actually a singleton class since it can not be instantiated
  TSingleAppInstance = class(TObject)
  private type
    TSwitchFunc = procedure(hwnd: HWND; fAltTab: BOOL); stdcall;

  private
    class var DllHandle:                THandle;
    class var SwitchToThisWindow:       TSwitchFunc;

    class var FAppMutex:                THandle;
    class var FPipeServer:              TPipeServerThread;
    class var FPipeClient:              TPipeClient;

    class var FAcceptIncomingData:      boolean;
    class var FInstanceState:           TAppInstanceState;

    class var FOnOtherInstanceSentData: TOtherInstanceSentDataEvent;

    class constructor Create();
    class destructor  Destroy;

    class function    CheckForAnotherInstance(): TAppInstanceState; static;
    class procedure   StartNamedPipeServer(); static;
    class procedure   StopNamedPipeServer(); static;
    class procedure   PipeServerClientSentData(Sender: TPipeServerThread; const Data: string); static;

    class function    GetInstanceState: TAppInstanceState; static;
    class function    GetIsAnotherInstanceRunning(): boolean; static;

  public
    constructor       Create;

    class procedure   SendCommandLineData(); static;

    class property    InstanceState:            TAppInstanceState           read GetInstanceState;
    class property    IsAnotherInstanceRunning: boolean                     read GetIsAnotherInstanceRunning;
    class property    AcceptIncomingData:       boolean                     read FAcceptIncomingData      write FAcceptIncomingData;

    class property    OnOtherInstanceSentData:  TOtherInstanceSentDataEvent read FOnOtherInstanceSentData write FOnOtherInstanceSentData;

  end;



implementation

const
  // Names for application mutex and named pipe
  APP_MUTEX: string = 'SingleAppInstanceMutex_74D30CE1-3C97-4C42-9BC7-E6C2205CFEB8';
  PIPE_NAME: string = 'SingleAppInstancePipe_BD705031-62BF-488E-929D-3AB8B7375CBD';


{ TSingleAppInstance }

constructor TSingleAppInstance.Create;
begin
  raise ENoConstructException.Create('Creating instances of ' + Self.ClassName + ' is not allowed!');
end;


class constructor TSingleAppInstance.Create();
begin
  // Dynamically load missing Win32 API function
  DllHandle                 := LoadLibrary(USER32);
  @SwitchToThisWindow       := GetProcAddress(DllHandle, 'SwitchToThisWindow');

  FAcceptIncomingData       := false;
  FInstanceState            := aisNone;
  FAppMutex                 := 0;
  FPipeServer               := nil;
  FPipeClient               := nil;
  FOnOtherInstanceSentData  := nil;
end;


class destructor TSingleAppInstance.Destroy;
begin
  FPipeClient.Free;
  StopNamedPipeServer();

  if FAppMutex <> 0 then
    ReleaseMutex(FAppMutex);

  if DllHandle <> 0 then
    FreeLibrary(DllHandle);
end;


class function TSingleAppInstance.GetInstanceState: TAppInstanceState;
begin
  if FInstanceState = aisNone then
    GetIsAnotherInstanceRunning();

  Result := FInstanceState;
end;


class function TSingleAppInstance.GetIsAnotherInstanceRunning(): boolean;
begin
  Result := false;

  FInstanceState := CheckForAnotherInstance();

  case FInstanceState of
    aisMainInstance:  StartNamedPipeServer();
    aisOtherInstance: Result := true;
    else              Result := true;
  end;
end;


// -----------------------------------------------------------------------------
// Pipe server part
// -----------------------------------------------------------------------------
class procedure TSingleAppInstance.StartNamedPipeServer;
begin
  // Current instance is the main instance. Start a background thread running
  // a named pipe server to receive seubsequently started instances' command
  // line data.
  FPipeServer := TPipeServerThread.Create(PIPE_NAME, PipeServerClientSentData);
  FPipeServer.Start;

  // Release sync mutex to unblock other instances waiting to send command line
  // data
  ReleaseMutex(FAppMutex);
  FAppMutex := 0;
end;


class procedure TSingleAppInstance.StopNamedPipeServer();
begin
  if Assigned(FPipeServer) then
    FPipeServer.Terminate;

  FPipeServer := nil;
end;


class procedure TSingleAppInstance.PipeServerClientSentData(Sender: TPipeServerThread; const Data: string);
begin
  // Check semaphore
  if not FAcceptIncomingData then exit;

  // Hand over command line data received from other instance
  // to handler method of main program
  if Assigned(FOnOtherInstanceSentData) then
    FOnOtherInstanceSentData(Data);
end;


// -----------------------------------------------------------------------------
// Pipe client part
// -----------------------------------------------------------------------------
class function TSingleAppInstance.CheckForAnotherInstance(): TAppInstanceState;
begin
  // This is a simpler approach that relies on the pipe itself to decide if the
  // current app instance has to work as the main instance or if it should act
  // as a subsequently started instance. But this is an unreliable approach.
  // If the main instance is terminated during a bunch of clients trying to
  // connect, we may end up with multiple "main instances". To overcome that,
  // the current approach is to use a mutex to decide to work as the main
  // instance or as a subsequently started instance.
  //
  //  FPipeClient := TPipeClient.Create;
  //
  //  if FPipeClient.ConnectToServer(PIPE_NAME) then
  //    Result := aisOtherInstance
  //  else
  //  begin
  //    FreeAndNil(FPipeClient);
  //    Result := aisMainInstance;
  //  end;

  Result := aisUnknown;

  // Create sync mutex and take its ownership
  FAppMutex := CreateMutex(nil, true, PChar(APP_MUTEX));
  if FAppMutex = 0 then exit;

  // If the mutex doesn't exist, the current instance is the main instance that
  // should receive command line data from subsequently started instances
  if GetLastError() <> ERROR_ALREADY_EXISTS then
    Result := aisMainInstance
  else
  begin
    // The mutex already exists, thus the current instance is a subsequently
    // started instance that should send command line data to the main instance.

    // Wait until the current instance gets ownership of the sync mutex in order
    // to get the exclusive right to communicate with the main instance's pipe
    // server.
    WaitForSingleObject(FAppMutex, INFINITE);

    // Create a pipe client ...
    FPipeClient := TPipeClient.Create;

    // ... and connect to the server.
    // On success set instance state to according value. Otherwise free pipe
    // client object.
    if FPipeClient.ConnectToServer(PIPE_NAME) then
      Result := aisOtherInstance
    else
      FreeAndNil(FPipeClient);
  end;
end;


class procedure TSingleAppInstance.SendCommandLineData();
begin
  // If current instance's pipe client has been successfully connected, ...
  if Assigned(FPipeClient) then
  begin
    // ... start data exchange with main instance's pipe server ...
    FPipeClient.TalkToServer;

    // ... and activate its window.
    if FPipeClient.MainInstanceWndHandle <> 0 then
      SwitchToThisWindow(FPipeClient.MainInstanceWndHandle, true);

    // At this point the pipe client is no longer needed.
    FreeAndNil(FPipeClient);
  end;
end;


end.
