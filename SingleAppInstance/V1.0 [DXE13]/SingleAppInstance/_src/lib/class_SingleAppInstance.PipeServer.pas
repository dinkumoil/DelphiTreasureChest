{ **************************************************************************** }
//
// Named pipe server running in a background thread, using I/O completion
// routines. Intended to be used in conjunction with single app instance
// component in unit class_SingleAppInstance.pas.
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

unit class_SingleAppInstance.PipeServer;


interface

uses
  Winapi.Windows, System.SysUtils, System.Types, System.Classes, Vcl.Forms;


type
  // Forward declaration
  TPipeServerThread = class;

  // Event type for assigning a class procedure
  TClientDataEvent = reference to procedure(Sender: TPipeServerThread; const ClientData: string);


  // This class can only be used as a singleton class
  TPipeServerThread = class(TThread)
  private type
    // Storage for managing per-client connection data
    PPipeInstance = ^TPipeInstance;

    TPipeInstance = record
      SyncRec:             TOverlapped;
      PipeHandle:          THandle;
      RequestBuf:          PByte;
      RequestBufTotalSize: Cardinal;
      RequestBufUsedSize:  Cardinal;
      ReplyBuf:            PByte;
      ReplyBufTotalSize:   Cardinal;
      ReplyBufUsedSize:    Cardinal;
      BytesToWrite:        Cardinal;
    end;

  private
    FPipeName:     string;
    FPipe:         THandle;
    FOnClientData: TClientDataEvent;

    class var FCurInstance: TPipeServerThread;

    function    CreateAndConnectInstance(var SyncRec: TOverlapped): boolean;
    function    ConnectToNewClient(var SyncRec: TOverlapped; PipeHandle: THandle): boolean;
    procedure   DisconnectAndClose(var PipeInst: PPipeInstance);
    procedure   ProcessClientData(const ClientData: string);

  protected
    procedure   Execute; override;
    procedure   TerminatedSet; override;

  public
    constructor Create(const APipeName: string; ClientDataHandler: TClientDataEvent);

    class property CurInstance: TPipeServerThread read FCurInstance;

  end;



implementation

const
  // Missing Win32 API constant
  PIPE_REJECT_REMOTE_CLIENTS: DWORD = $00000008;

  // Client connect timeout in ms
  PIPE_TIMEOUT: DWORD = 200;

  // Pipe input/output buffer sizes, in bytes
  READ_BUF_SIZE  = 1024;
  WRITE_BUF_SIZE = 16;


// I/O completion routines. Since these are Win32 API callback functions with a
// certain signature, they can not be object instance methods. They could also
// be static class methods but that approach has some disadvantages.
procedure SendMainInstanceData(lpOverlap: POverlapped); forward;
procedure RequestOtherInstanceData(dwErr: DWORD; cbBytesWritten: DWORD; lpOverlap: POverlapped); stdcall; forward;
procedure ReceiveOtherInstanceData(dwErr: DWORD; cbBytesRead: DWORD; lpOverlap: POverlapped); stdcall; forward;
function  CheckForServerThreadTermination(PipeInst: TPipeServerThread.PPipeInstance): boolean; forward;
procedure DisconnectAndClose(PipeInst: TPipeServerThread.PPipeInstance); forward;


{ TPipeServerThread }

constructor TPipeServerThread.Create(const APipeName: string; ClientDataHandler: TClientDataEvent);
begin
  inherited Create(true);

  FPipeName       := APipeName;
  FOnClientData   := ClientDataHandler;
  FPipe           := INVALID_HANDLE_VALUE;
  FCurInstance    := Self;

  // Auto-free object when thread is terminated
  FreeOnTerminate := true;
end;


// Overridden TThread method.
// This function is the pipe server's thread working routine. It implements the
// main loop of the pipe server.
procedure TPipeServerThread.Execute;
var
  ConnectEvent: THandle;
  ConnectSync:  TOverlapped;
  PipeInst:     PPipeInstance;
  WaitResult:   DWORD;
  BytesRead:    DWORD;
  IsSuccess:    BOOL;
  IsPendingIO:  boolean;

begin
  // Create single event object for client connect operations.
  ConnectEvent := CreateEvent(nil,    // default security attribute
                              true,   // manual reset event
                              true,   // initial state = signaled
                              nil);   // unnamed event object

  if ConnectEvent = 0 then
    exit;

  // Initialize single OVERLAPPED struct for establishing a client connection
  // and store event object's handle
  ZeroMemory(@ConnectSync, SizeOf(TOverlapped));
  ConnectSync.hEvent := ConnectEvent;

  // Create pipe instance and wait for a client to connect.
  IsPendingIO := CreateAndConnectInstance(ConnectSync);

  // Pipe server's main loop
  while true do
  begin
    // Wait for a client to connect, or for a read or write
    // operation to be completed, which causes a completion
    // routine to be queued for execution.
    WaitResult := WaitForSingleObjectEx(ConnectEvent,  // event object to wait for
                                        INFINITE,      // waits indefinitely
                                        true);         // alertable wait enabled

    // Process wait result
    case WaitResult of
      // The wait conditions are satisfied by a completed connect operation
      // or if the connect operation has been canceled.
      0:
      begin
        // If a connect operation is pending, get its result.
        if IsPendingIO then
        begin
          BytesRead := 0;

          IsSuccess := GetOverlappedResult(FPipe,        // pipe handle
                                           ConnectSync,  // OVERLAPPED structure
                                           BytesRead,    // bytes transferred
                                           false);       // do not wait

          // Break main loop on error or if connect operation has been canceled.
          if not IsSuccess then
            break;
        end;

        // Allocate storage for pipe instance.
        New(PipeInst);
        ZeroMemory(PipeInst, SizeOf(TPipeInstance));

        // Allocate memory for output buffer
        GetMem(PipeInst.RequestBuf, WRITE_BUF_SIZE);
        PipeInst.RequestBufTotalSize := WRITE_BUF_SIZE;
        PipeInst.RequestBufUsedSize  := 0;

        // Allocate memory for input buffer
        GetMem(PipeInst.ReplyBuf, READ_BUF_SIZE);
        PipeInst.ReplyBufTotalSize := READ_BUF_SIZE;
        PipeInst.ReplyBufUsedSize  := 0;

        // Save pipe handle in instance's storage
        PipeInst.PipeHandle := FPipe;

        // Start the I/O sequence for communicating with the client.
        SendMainInstanceData(@PipeInst.SyncRec);

        // Drop responsibility for pipe's data storage and its handle
        PipeInst := nil;
        FPipe    := INVALID_HANDLE_VALUE;

        // Create new pipe instance for the next client.
        IsPendingIO := CreateAndConnectInstance(ConnectSync);
      end;

      // The wait conditions are satisfied by execution of an I/O completion
      // routine. This allows async I/O via APC.
      WAIT_IO_COMPLETION:
        ;

      // An error occurred in the wait function
      // or the I/O operation has been cancelled
      else
        break;
    end;
  end;

  // Free pipe resources
  CloseHandle(ConnectEvent);
  DisconnectAndClose(PipeInst);
end;


// Overridden TThread method.
// This function cancels all pending I/O operations for the currently active pipe
procedure TPipeServerThread.TerminatedSet;
begin
  if FPipe <> INVALID_HANDLE_VALUE then
    CancelIoEx(FPipe, nil);

  inherited;
end;


// This function creates a pipe instance and connects to the client.
// It returns TRUE if the connect operation is pending, and FALSE if
// the connection has been completed.
function TPipeServerThread.CreateAndConnectInstance(var SyncRec: TOverlapped): boolean;
var
  PipeName: string;

begin
  PipeName := '\\.\pipe\' + FPipeName;

  FPipe := CreateNamedPipe(PChar(PipeName),                 // pipe name
                           PIPE_ACCESS_DUPLEX               // read/write access
                             or FILE_FLAG_OVERLAPPED,       // overlapped mode
                           PIPE_TYPE_MESSAGE                // message-type pipe
                             or PIPE_READMODE_MESSAGE       // message read mode
                             or PIPE_WAIT                   // blocking mode
                             or PIPE_REJECT_REMOTE_CLIENTS, // no remote clients
                           PIPE_UNLIMITED_INSTANCES,        // unlimited instances
                           WRITE_BUF_SIZE,                  // output buffer size
                           READ_BUF_SIZE,                   // input buffer size
                           PIPE_TIMEOUT,                    // client time-out
                           nil);                            // default security attributes

  if FPipe = INVALID_HANDLE_VALUE then
    exit(false);

  // Call a subroutine to connect to the new client.
  Result := ConnectToNewClient(SyncRec, FPipe);
end;


// This function establishes the connection between the server and a client.
// It returns TRUE if the connect operation is pending, and FALSE if the
// connection has been completed.
// The event object whose handle is part of the OVERLAPPED structure is set
// to the non-signalled state due to calling ConnectNamedPipe. But if a client
// connected between the calls to CreateNamedPipe and ConnectNamedPipe, the
// event object is set to the signaled state.
function TPipeServerThread.ConnectToNewClient(var SyncRec: TOverlapped; PipeHandle: THandle): boolean;
var
  IsConnected: BOOL;
  IsPendingIO: boolean;

begin
  IsPendingIO := false;

  // Start an overlapped connection for this pipe instance.
  IsConnected := ConnectNamedPipe(PipeHandle, @SyncRec);

  // Overlapped ConnectNamedPipe should return zero.
  if IsConnected then
    exit(false);

  case GetLastError() of
    // The server-side pipe end is ready to accept a client connection.
    ERROR_IO_PENDING:
      IsPendingIO := true;

    // Client is already connected, so signal an event.
    ERROR_PIPE_CONNECTED:
      SetEvent(SyncRec.hEvent);
  end;

  Result := IsPendingIO;
end;


// This routine is called when
//  - the main thread terminates the server thread
//  - an error occurs during opening a pipe or establishing a connection between
//    the pipe server and a client
procedure TPipeServerThread.DisconnectAndClose(var PipeInst: PPipeInstance);
begin
  if FPipe <> INVALID_HANDLE_VALUE then
  begin
    // Disconnect pipe instance.
    DisconnectNamedPipe(FPipe);

    // Close handle of pipe instance.
    CloseHandle(FPipe);
  end;

  if Assigned(PipeInst) then
  begin
    // Release storage of pipe instance.
    FreeMem(PipeInst.RequestBuf);
    FreeMem(PipeInst.ReplyBuf);
    Dispose(PipeInst);
  end;

  // House keeping
  FPipe    := INVALID_HANDLE_VALUE;
  PipeInst := nil;
end;


// This routine transmits the client data received by the server
// to the main thread using the server thread's Synchronize method
procedure TPipeServerThread.ProcessClientData(const ClientData: string);
begin
  Synchronize(procedure
    begin
      if Assigned(FOnClientData) then
        FOnClientData(Self, ClientData);
    end
  );
end;



// -----------------------------------------------------------------------------
// I/O completion routines stuff
// -----------------------------------------------------------------------------

// This routine is called as a normal routine. It runs in the context of the
// server thread. It performs step 1 of the client-server communication.
procedure SendMainInstanceData(lpOverlap: POverlapped);
var
  MainFormHandle: THandle;
  PipeInst:       TPipeServerThread.PPipeInstance;
  IsSuccess:      BOOL;

begin
  IsSuccess := false;

  // lpOverlap points to storage for this instance.
  PipeInst := TPipeServerThread.PPipeInstance(lpOverlap);

  // Write main form window handle to other instance (if there is no
  // pending request to terminate server thread).
  if not TPipeServerThread.CurInstance.Terminated then
  begin
    // Send application's main form handle to other instance.
    // Thus, this instance can activate main instance's window.
    MainFormHandle := Application.MainFormHandle;

    PipeInst.BytesToWrite := SizeOf(THandle);
    move(MainFormHandle, PipeInst.RequestBuf[0], PipeInst.BytesToWrite);

    IsSuccess := WriteFileEx(PipeInst.PipeHandle,
                             PipeInst.RequestBuf,
                             PipeInst.BytesToWrite,
                             PipeInst.SyncRec,
                             @RequestOtherInstanceData);
  end;

  // Disconnect if an error occurred.
  if not IsSuccess then
    DisconnectAndClose(PipeInst);
end;


// This routine is called as an I/O completion routine.
// It performs step 2 of the client-server communication.
procedure RequestOtherInstanceData(dwErr: DWORD; cbBytesWritten: DWORD; lpOverlap: POverlapped); stdcall;
var
  PipeInst:  TPipeServerThread.PPipeInstance;
  IsSuccess: BOOL;

begin
  IsSuccess := false;

  // lpOverlap points to storage for this instance.
  PipeInst := TPipeServerThread.PPipeInstance(lpOverlap);

  // The write operation has finished, so read the next request (if there is
  // no pending request to terminate server thread and no error occured).
  if CheckForServerThreadTermination(PipeInst)                and
     (dwErr = 0) and (cbBytesWritten = PipeInst.BytesToWrite) then
  begin
    // Read other instance's command line
    IsSuccess := ReadFileEx(PipeInst.PipeHandle,
                            PipeInst.ReplyBuf + PipeInst.ReplyBufUsedSize,
                            PipeInst.ReplyBufTotalSize - PipeInst.ReplyBufUsedSize,
                            @PipeInst.SyncRec,
                            @ReceiveOtherInstanceData);
  end;

  // Disconnect if an error occurred.
  if not IsSuccess then
    DisconnectAndClose(PipeInst)
end;


// This routine is called as an I/O completion routine.
// It performs step 3 of the client-server communication.
procedure ReceiveOtherInstanceData(dwErr: DWORD; cbBytesRead: DWORD; lpOverlap: POverlapped); stdcall;
var
  PipeInst:   TPipeServerThread.PPipeInstance;
  IsSuccess:  BOOL;
  BytesRead:  DWORD;
  ClientData: string;

begin
  // lpOverlap points to storage for this instance.
  PipeInst := TPipeServerThread.PPipeInstance(lpOverlap);

  // Read next data packet (if there is no pending request
  // to terminate server thread and no error occured).
  if CheckForServerThreadTermination(PipeInst) and
     (dwErr = 0) and (cbBytesRead > 0)         then
  begin
    // Update used size of buffer
    PipeInst.ReplyBufUsedSize := PipeInst.ReplyBufUsedSize + cbBytesRead;

    // Retrieve last I/O operation result
    BytesRead := 0;

    IsSuccess := GetOverlappedResult(PipeInst.PipeHandle,  // pipe handle
                                     PipeInst.SyncRec,     // OVERLAPPED structure
                                     BytesRead,            // bytes transferred
                                     false);               // do not wait

    // If pipe buffer was too small we have to read more data
    if not IsSuccess and (GetLastError() = ERROR_MORE_DATA) then
    begin
      // If buffer size has to be increased...
      if PipeInst.ReplyBufUsedSize = PipeInst.ReplyBufTotalSize then
      begin
        // enlarge buffer and remember new size
        ReAllocMem(PipeInst.ReplyBuf, PipeInst.ReplyBufTotalSize + READ_BUF_SIZE);
        Inc(PipeInst.ReplyBufTotalSize, READ_BUF_SIZE);
      end;

      // Write operation is done, update related data
      PipeInst.BytesToWrite := 0;

      // Request another async read operation and exit
      RequestOtherInstanceData(dwErr, 0, lpOverlap);
      exit;
    end
    else
    begin
      // Retrieve read data from buffer ...
      SetString(ClientData, PChar(PipeInst.ReplyBuf), PipeInst.ReplyBufUsedSize div SizeOf(char));

      // ... and hand it over to the thread object
      TPipeServerThread.CurInstance.ProcessClientData(ClientData);
    end;
  end;

  // Communication finished, free pipe resources
  DisconnectAndClose(PipeInst);
end;


// Check if pipe server's thread has been terminated externally.
// If yes, cancel pending I/O operation of pipe and tell caller to break.
function CheckForServerThreadTermination(PipeInst: TPipeServerThread.PPipeInstance): boolean;
var
  Dummy: DWORD;

begin
  // Default return value is telling caller to continue
  Result := true;

  // Check for pipe server thread's termination
  if TPipeServerThread.CurInstance.Terminated then
  begin
    // Request cancelling pipe I/O operation
    CancelIoEx(PipeInst.PipeHandle, @PipeInst.SyncRec);

    // Wait until cancelling has been done
    Dummy := 0;

    GetOverlappedResult(PipeInst.PipeHandle,  // pipe handle
                        PipeInst.SyncRec,     // OVERLAPPED structure
                        Dummy,                // unused value
                        true);                // wait

    // Tell caller to break
    Result := false;
  end;
end;


// This routine is called when
//  - the server has received all client data
//  - the client closes its handle to the pipe
//  - the main thread terminates the server thread
//  - an error occurs during execution of an I/O completion routine
procedure DisconnectAndClose(PipeInst: TPipeServerThread.PPipeInstance);
begin
  if not Assigned(PipeInst) then exit;

  // Disconnect pipe instance.
  DisconnectNamedPipe(PipeInst.PipeHandle);

  // Close handle of pipe instance.
  CloseHandle(PipeInst.PipeHandle);

  // Release storage of pipe instance.
  FreeMem(PipeInst.RequestBuf);
  FreeMem(PipeInst.ReplyBuf);
  Dispose(PipeInst);
end;


end.
