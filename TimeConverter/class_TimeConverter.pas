unit class_TimeConverter;


interface

uses
  Winapi.Windows, System.SysUtils, System.DateUtils, Soap.XSBuiltIns;


type
  TTimeConverter = class
  public
    class function DateTimeAsUnixTimestamp(const DT: TDateTime): Int64; static;
    class function DateTimeAsUnixUTCTimestamp(const DT: TDateTime): Int64; static;
    class function DateTimeAsSystemTime(const DT: TDateTime): TSystemTime; static;
    class function DateTimeAsFileTime(const DT: TDateTime): TFileTime; static;
    class function DateTimeAsXmlUTCTime(const DT: TDateTime): string; static;
    class function DateTimeAsXmlTime(const DT: TDateTime): string; static;

    class function UnixTimestampAsDateTime(const TS: Int64): TDateTime; static;
    class function UnixUTCTimestampAsDateTime(const UTCTS: Int64): TDateTime; static;
    class function SystemTimeAsDateTime(const ST: TSystemTime): TDateTime; static;
    class function FileTimeAsDateTime(const FT: TFileTime): TDateTime; static;
    class function XmlTimeAsDateTime(const XT: string): TDateTime; static;

  end;



implementation

{ TTimeConverter }

// -----------------------------------------------------------------------------
// Convert from TDateTime to other formats
// -----------------------------------------------------------------------------

class function TTimeConverter.DateTimeAsUnixTimestamp(const DT: TDateTime): Int64;
begin
  // Input: Local time
  // Output: Unix time als Local time
  Result := DateTimeToUnix(DT);
end;


class function TTimeConverter.DateTimeAsUnixUTCTimestamp(const DT: TDateTime): Int64;
var
  UTCDT: TDateTime;

begin
  // Input: Local time
  // Output: Unix time as UTC
  UTCDT  := TTimeZone.Local.ToUniversalTime(DT);
  Result := DateTimeToUnix(UTCDT);
end;


class function TTimeConverter.DateTimeAsSystemTime(const DT: TDateTime): TSystemTime;
begin
  // Input: Local time
  // Output: System time as local time
  DateTimeToSystemTime(DT, Result);
end;


class function TTimeConverter.DateTimeAsFileTime(const DT: TDateTime): TFileTime;
var
  ST: TSystemTime;

begin
  // Input: Local time
  // Output: File time as local time
  ST := DateTimeAsSystemTime(DT);
  SystemTimeToFileTime(ST, Result);
end;


class function TTimeConverter.DateTimeAsXmlTime(const DT: TDateTime): string;
begin
  // Input: Local time
  // Output: Local time incl. UTC offset according to ISO 8601
  Result := DateTimeToXMLTime(DT, true);
end;


class function TTimeConverter.DateTimeAsXmlUTCTime(const DT: TDateTime): string;
begin
  // Input: Local time
  // Output: UTC time according to ISO 8601
  Result := DateTimeToXMLTime(UnixTimestampAsDateTime(DateTimeAsUnixUTCTimestamp(DT)), false);
end;


// -----------------------------------------------------------------------------
// Convert from other formats to TDateTime
// -----------------------------------------------------------------------------

class function TTimeConverter.UnixTimestampAsDateTime(const TS: Int64): TDateTime;
begin
  // Input: Unix time as local time
  // Output: Local time
  Result := UnixToDateTime(TS);
end;


class function TTimeConverter.UnixUTCTimestampAsDateTime(const UTCTS: Int64): TDateTime;
var
  UTCDT: TDateTime;

begin
  // Input: Unix time as UTC time
  // Output: Local time
  UTCDT  := UnixToDateTime(UTCTS);
  Result := TTimeZone.Local.ToLocalTime(UTCDT);
end;


class function TTimeConverter.SystemTimeAsDateTime(const ST: TSystemTime): TDateTime;
begin
  // Input: System time as local time
  // Output: Local time
  Result := SystemTimeToDateTime(ST);
end;


class function TTimeConverter.FileTimeAsDateTime(const FT: TFileTime): TDateTime;
var
  ST: TSystemTime;

begin
  // Input: File time as local time
  // Output: Local time
  FileTimeToSystemTime(@FT, ST);
  Result := SystemTimeAsDateTime(ST);
end;


class function TTimeConverter.XmlTimeAsDateTime(const XT: string): TDateTime;
begin
  // Input: UTC time or local time incl. UTC offset according to ISO 8601
  // Output: Local time
  if not XT.IsEmpty then
    Result := XMLTimeToDateTime(XT, false)
  else
    Result := Default(TDateTime);
end;


end.
