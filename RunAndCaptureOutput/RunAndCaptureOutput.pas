function RunAndCaptureOutput(const AFile, Params: string): string;
var
  SA:              TSecurityAttributes;
  SI:              TStartupInfo;
  PI:              TProcessInformation;
  StdOutPipeRead:  THandle;
  StdOutPipeWrite: THandle;
  WasOK:           LongBool;
  CmdLine:         string;
  WorkDir:         string;
  ResultCode:      cardinal;
  Encoding:        TEncoding;
  Buffer:          TBytes;
  BytesRead:       cardinal;

begin
  Result := '';

  with SA do
  begin
    nLength              := SizeOf(SA);
    bInheritHandle       := True;
    lpSecurityDescriptor := nil;
  end;

  // Create output pipe for child process
  if not CreatePipe(StdOutPipeRead, StdOutPipeWrite, @SA, 0) then
    RaiseLastOSError;

  try
    // Prevent inheritance of handle to pipe's read end
    if not SetHandleInformation(StdOutPipeRead, HANDLE_FLAG_INHERIT, 0) then
      RaiseLastOSError;

    // Fill StartupInfo struct
    ZeroMemory(@SI, SizeOf(SI));

    with SI do
    begin
      cb          := SizeOf(SI);
      dwFlags     := STARTF_USESHOWWINDOW or STARTF_USESTDHANDLES;
      wShowWindow := SW_HIDE;
      hStdInput   := GetStdHandle(STD_INPUT_HANDLE); // Do not redirect STDIN
      hStdOutput  := StdOutPipeWrite;                // Redirect STDOUT to pipe
      hStdError   := GetStdHandle(STD_ERROR_HANDLE); // Do not redirect STDERR
    end;

    // Retrieve working directory from executable's path
    CmdLine := AnsiQuotedStr(AFile, '"') + ' ' + Params;
    WorkDir := ExtractFileDir(AFile);

    // If executable is provided without a path, set current directory
    // as working directory
    if WorkDir = '' then
      WorkDir := GetCurrentDir();

    // Run executable
    WasOK := CreateProcess(nil, PChar(CmdLine),
                           nil, nil, True, 0, nil,
                           PChar(WorkDir), SI, PI);

    // The write end of the pipe is now controlled by the child process,
    // so its handle can be closed and marked as invalid
    CloseHandle(StdOutPipeWrite);
    StdOutPipeWrite = INVALID_HANDLE_VALUE;

    // Exit if creating the child process has failed
    if not WasOK then
    begin
      ShowMessage(SysErrorMessage(GetLastError));
      exit;
    end;

    // Read data from pipe sent by the child process
    SetLength(Buffer, 256);
    Encoding := TEncoding.GetEncoding(GetConsoleOutputCP());

    try
      // Read data in 256 byte chunks
      repeat
        WasOK := ReadFile(StdOutPipeRead, Buffer[0], Length(Buffer), BytesRead, nil);

        // Do character encoding conversion
        if BytesRead > 0 then
          Result := Result + Encoding.GetString(Buffer, 0, BytesRead);
      until (not WasOK) or (BytesRead = 0);

      // Wait until child process terminates
      WaitForSingleObject(PI.hProcess, INFINITE);

      // Retrieve exit code of child process. If an error has been reported,
      // return an empty string to the caller
      GetExitCodeProcess(PI.hProcess, ResultCode);
      if ResultCode <> 0 then Result := '';

    finally
      // Clean up
      Encoding.Free;
      SetLength(Buffer, 0);
      CloseHandle(PI.hThread);
      CloseHandle(PI.hProcess);
    end;

  finally
    // Free pipe handles
    CloseHandle(StdOutPipeRead);

    if StdOutPipeWrite <> INVALID_HANDLE_VALUE then
      CloseHandle(StdOutPipeWrite);
  end;
end;
