uses
  Winapi.Windows, System.SysUtils, System.IOUtils;


function ExpandFilePath(const AFilePath: string): string;
var
  FileNameLen: cardinal;
  TmpFilePath: string;

begin
  Result := '';

  FileNameLen := ExpandEnvironmentStrings(PChar(AFilePath), PChar(TmpFilePath), 0);
  if FileNameLen = 0 then exit;

  SetLength(TmpFilePath, Pred(FileNameLen));

  if ExpandEnvironmentStrings(PChar(AFilePath), PChar(TmpFilePath), FileNameLen) = FileNameLen then
  begin
    if TPath.IsUNCPath(TmpFilePath) then
      Result := ExpandUNCFileName(TmpFilePath)
    else
      Result := ExpandFileName(TmpFilePath);
  end;
end;
