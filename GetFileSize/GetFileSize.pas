uses
  WinApi.Windows;


function FileSize(const AFilename: String): Int64;
var
  Info: TWin32FileAttributeData;

begin
  Result := -1;

  if NOT GetFileAttributesEx(PWideChar(AFileName), GetFileExInfoStandard, @Info) then
    EXIT;

  Result := Int64(Info.nFileSizeLow) or Int64(Info.nFileSizeHigh shl 32);
end;
