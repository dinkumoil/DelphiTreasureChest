uses
  WinApi.Windows;
  
 
function IsConsoleApp(const Path: String): boolean;
var
  handle:     cardinal;
  bytesread:  DWORD;
  signature:  DWORD;
  dos_header: _IMAGE_DOS_HEADER;
  pe_header:  _IMAGE_FILE_HEADER;
  opt_header: _IMAGE_OPTIONAL_HEADER;

begin
  Result := false;

  handle := CreateFile(PChar(Path),
                       GENERIC_READ,
                       FILE_SHARE_READ,
                       nil,
                       OPEN_EXISTING,
                       FILE_ATTRIBUTE_NORMAL,
                       0);

  if handle <> INVALID_HANDLE_VALUE then
  begin
    ReadFile(Handle, dos_header, sizeof(dos_header), bytesread, nil);
    SetFilePointer(Handle, dos_header._lfanew, nil, 0);
    ReadFile(Handle, signature, sizeof(signature), bytesread, nil);
    ReadFile(Handle, pe_header, sizeof(pe_header), bytesread, nil);
    ReadFile(Handle, opt_header, sizeof(opt_header), bytesread, nil);

    Result := (opt_header.Subsystem = IMAGE_SUBSYSTEM_WINDOWS_CUI);
  end;

  CloseHandle(handle);
end;
