unit Settings;


interface

uses
  Winapi.Windows, System.SysUtils, System.IOUtils, System.Classes, Vcl.Graphics,

  System.IniFiles.Persistence;


type
  TGetFilePathEvent = function: string of object;
  TGetFilePathFunc  = TFunc<string>;


  TSettings = class(TObject)
  public type
    TDynStrArray   = array of string;                  // Traditional dynamic array; RTTI for array and element type available
    TStrArrayIndex = 0..2;                             // Subrange type; RTTI available
    TStrArrayFail  = array[0..2] of string;            // Static array; RTTI for array and element type available, no RTTI for inplace index type
    TStrArray      = array[TStrArrayIndex] of string;  // Static array; RTTI for array, element and index type available

  private
    FVersion:      string;
    FActive:       boolean;
    FDateTimeVal:  TDateTime;
    FDateTimeVal2: TDateTime;
    FDateTimeVal3: TDateTime;
    FDateTimeVal4: TDateTime;
    FColorVal:     TColor;
    FEnumVal:      TTextFormats;
    FSetValue:     TTextFormat;
    FGUIDValue:    TGUID;
    FCfgFileNames: TArray<string>;           // Preferred common case: modern dynamic array
    FDatFileNames: TDynStrArray;             // Common case: traditional dynamic array
    FTmpFileNames: TStrArray;                // Uncommon case: static array
    // FTmpFileNames: array of string;       => Inplace type declaration of traditional dynamic array, does not compile
    // FTmpFileNames: array[0..2] of string; => Inplace type declaration of static array, does not compile
    // FTmpFileNames: TStrArrayFail;         => Static array with inplace index type declaration, causes runtime error
    FGUIDs:        TArray<TGUID>;

  private
    class var FInstance:      TSettings;
    class var FFilePath:      string;
    class var FOnGetFilePath: TGetFilePathEvent;
    class var FFnGetFilePath: TGetFilePathFunc;

    class function GetInstance: TSettings; static;
    class function GetFilePath: string; static;

  public
    [IniIntValue('Test', 'IntVal', -2147483648, true)]
    IntVal: integer;

    [IniIntValue('Test', 'IntVal2', 2147483647, true)]
    IntVal2: integer;

    [IniUIntValue('Test', 'UIntVal', 4294967295)]
    UIntVal: cardinal;

    [IniInt64Value('Test', 'Int64Val', -9223372036854775808, true)]
    Int64Val: int64;

    [IniInt64Value('Test', 'Int64Val2', 9223372036854775807, true)]
    Int64Val2: int64;

    [IniUInt64Value('Test', 'UInt64Val', 18446744073709551615)]
    UInt64Val: uint64;

    [IniFloatValue('Test', 'FloatVal', 100.5)]
    FloatVal: double;

    [IniStrValue('Test', 'InFile')]
    InFileNames: TArray<string>;  // Modern dynamic array of strings

    [IniStrValue('Test', 'OutFile', '')]
    OutFileNames: TDynStrArray;  // Traditional dynamic array of strings

    [IniStrValue('Test', 'LogFile', 'C:\Foo\Bar Baz.log')]
    LogFileNames: TStrArray;  // Static array of strings

    [IniSetValue('Test', 'Set', '[tfCalcRect, tfWordBreak]'), IniArrayLength(4)]  // Provide sets as string
    Sets: TArray<TTextFormat>;  // Modern dynamic array of sets with at least 4 elements

    constructor Create; overload;
    constructor Create(const AFilePath: string); overload;
    constructor Create(AFilePathGetter: TGetFilePathEvent); overload;
    constructor Create(AFilePathGetter: TGetFilePathFunc); overload;

    destructor  Destroy; override;

    class procedure Load; static;
    class procedure Save; static;

    class property  Instance: TSettings read GetInstance;
    class property  FilePath: string    read GetFilePath;

    class property  OnGetFilePath: TGetFilePathEvent read FOnGetFilePath write FOnGetFilePath;
    class property  FnGetFilePath: TGetFilePathFunc  read FFnGetFilePath write FFnGetFilePath;

    [IniStrValue('Header', 'Version', '1.0')]
    property Version: string read FVersion write FVersion;

    [IniBoolValue('Settings', 'Active', false)]
    property Active: boolean read FActive write FActive;

    [IniDateTimeValue('Test', 'DateTimeVal', 0)]  // Resolves to UTC equivalent of local time 1899-12-30T00:00:00.000 (start of Delphi epoch)
    property DateTimeVal: TDateTime read FDateTimeVal write FDateTimeVal;

    [IniDateTimeValue('Test', 'DateTimeVal2', 0, true)]  // Resolves to UTC 1899-12-30T00:00:00.000Z (start of Delphi epoch)
    property DateTimeVal2: TDateTime read FDateTimeVal2 write FDateTimeVal2;

    [IniDateTimeValue('Test', 'DateTimeVal3', '1899-12-30T01:00:00.000+01:00')]  // Start of Delphi epoch as local time in CET
    property DateTimeVal3: TDateTime read FDateTimeVal3 write FDateTimeVal3;

    [IniDateTimeValue('Test', 'DateTimeVal4', '1899-12-30T00:00:00.000Z')]  // Start of Delphi epoch as UTC
    property DateTimeVal4: TDateTime read FDateTimeVal4 write FDateTimeVal4;

    [IniColorValue('Test', 'ColorVal', clNone)]
    property ColorVal: TColor read FColorVal write FColorVal;

    [IniEnumValue('Test', 'EnumVal', Ord(tfPathEllipsis))]
    property EnumVal: TTextFormats read FEnumVal write FEnumVal;

    [IniSetValue('Test', 'SetVal', '[tfCalcRect, tfWordBreak]')]  // Provide sets as string
    property SetVal: TTextFormat read FSetValue write FSetValue;

    [IniGUIDValue('Test', 'GUIDVal', '{00000000-0000-0000-0000-000000000000}')]  // Provide GUIDs as string
    property GUIDVal: TGUID read FGUIDValue write FGUIDValue;

    [IniStrValue('Test', 'CfgFile', ''), IniArrayLength(3)]
    property CfgFileNames: TArray<string> read FCfgFileNames write FCfgFileNames;

    [IniStrValue('Test', 'DatFile', 'C:\Foo\BarBaz.dat'), IniArrayLength(3)]
     property DatFileNames: TDynStrArray read FDatFileNames write FDatFileNames;

    [IniStrValue('Test', 'TmpFile')]
    // property TmpFileNames: array of string       read FTmpFileNames write FTmpFileNames; => Does not compile
    // property TmpFileNames: array[0..2] of string read FTmpFileNames write FTmpFileNames; => Does not compile
    // property TmpFileNames: TStrArrayFail         read FTmpFileNames write FTmpFileNames; => Causes runtime error
    property TmpFileNames: TStrArray read FTmpFileNames write FTmpFileNames;

    [IniGUIDValue('Test', 'GUID'), IniArrayLength(3)]  // Default value for each array element is TGUID.Empty
    property GUIDs: TArray<TGUID> read FGUIDs write FGUIDs;

  end;


var
  SettingsAutoLoad: boolean = true;



implementation

{ TSettings }

constructor TSettings.Create;
begin
  if Assigned(FInstance) then
    raise EInvalidOperation.CreateFmt('Class %s is a singleton. Multiple instantiation is not allowed.', [ClassName]);

  inherited;

  FInstance := Self;
end;


constructor TSettings.Create(const AFilePath: string);
begin
  Create;

  FFilePath := AFilePath;
  Load;
end;


constructor TSettings.Create(AFilePathGetter: TGetFilePathFunc);
begin
  FFnGetFilePath := AFilePathGetter;
  Create('');
end;


constructor TSettings.Create(AFilePathGetter: TGetFilePathEvent);
begin
  FOnGetFilePath := AFilePathGetter;
  Create('');
end;


destructor TSettings.Destroy;
begin
  Save;

  inherited;
end;


class function TSettings.GetInstance: TSettings;
begin
  if not Assigned(FInstance) then
    TSettings.Create;

  Result := FInstance;
end;


class function TSettings.GetFilePath: string;
var
  ExeName:      string;
  AppDataPath:  string;
  SettingsPath: string;

begin
  if Assigned(FOnGetFilePath) then
    FFilePath := FOnGetFilePath()

  else if Assigned(FFnGetFilePath) then
    FFilePath := FFnGetFilePath()

  else if FFilePath = '' then
  begin
    SetLength(AppDataPath, ExpandEnvironmentStrings('%AppData%', nil, 0) - 1);
    ExpandEnvironmentStrings('%AppData%', PChar(AppDataPath), Length(AppDataPath) + 1);

    ExeName      := TPath.GetFileNameWithoutExtension(ParamStr(0));
    SettingsPath := TPath.Combine(AppDataPath, ExeName);
    FFilePath    := TPath.Combine(SettingsPath, ExeName + '.ini');
  end;

  Result := FFilePath;
end;


class procedure TSettings.Load;
begin
  TIniPersistence.Load(FilePath, Instance);
end;


class procedure TSettings.Save;
var
  SettingsPath: string;

begin
  SettingsPath := TPath.GetDirectoryName(FilePath);

  if not TDirectory.Exists(SettingsPath) then
    ForceDirectories(SettingsPath);

  TIniPersistence.Save(FilePath, Instance);
end;



initialization

if SettingsAutoLoad then
  TSettings.Load;



finalization

TSettings.Instance.Free;


end.
