// MIT License
//
// Original Author: Robert Love, 2009
// Extended:        Andreas Heim, 2015 - 2024
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE

unit System.IniFiles.Persistence;


interface

uses
  System.SysUtils, System.DateUtils, System.Classes, System.UITypes, System.TypInfo,
  System.Rtti, Vcl.Graphics, Vcl.Dialogs;


type
  // ***************************************************************************
  // Forward declarations
  // ***************************************************************************

  IniValueAttribute         = class;
  IniStrValueAttribute      = class;
  IniBoolValueAttribute     = class;
  IniIntValueAttribute      = class;
  IniUIntValueAttribute     = class;
  IniInt64ValueAttribute    = class;
  IniUInt64ValueAttribute   = class;
  IniFloatValueAttribute    = class;
  IniDateTimeValueAttribute = class;
  IniColorValueAttribute    = class;
  IniEnumValueAttribute     = class;
  IniSetValueAttribute      = class;
  IniGUIDValueAttribute     = class;
  TIniPersistence           = class;


  // ***************************************************************************
  // Options for storing values to INI file
  // ***************************************************************************

  TIniStorageOption = (
    isoAsHex
  );

  TIniStorageOptions = set of TIniStorageOption;


  // ***************************************************************************
  // Interface for storing numerical values as hex string
  // ***************************************************************************

  IAsHex = interface(IInterface)
  ['{D2974DB2-2316-425F-B98B-B9ADC5A43A5F}']
    function GetAsHex: boolean;
    property AsHex: boolean read GetAsHex;
  end;


  // ***************************************************************************
  // Attribute classes
  // ***************************************************************************

  // Base class for all attribute classes, should NOT be used directly
  IniValueAttribute = class(TCustomAttribute)
  strict protected
    FSection:      string;
    FName:         string;
    FDefaultValue: string;

  public
    constructor Create(const aSection, aName: string; const aDefaultValue: string = '');

    function    DefaultValue(const aTypeInfo: PTypeInfo): string; virtual;

    property    Section: string read FSection write FSection;
    property    Name:    string read FName    write FName;

  end;


  // Base class for all attribute classes supporting interfaces, should NOT be used directly
  InterfacedIniValueAttribute = class(IniValueAttribute, IInterface)
  protected
    // Methods of IInterface, reference counting is deactivated
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;

  end;


  // This attribute class can also be used as IniStrValue
  IniStrValueAttribute = class(IniValueAttribute)
  end;


  // This attribute class can also be used as IniBoolValue
  IniBoolValueAttribute = class(IniValueAttribute)
  public
    constructor Create(const aSection, aName: string; const aDefaultValue: boolean = false); overload;
  end;


  // Base class for numeric attribute classes, should NOT be used directly
  IniNumericValueAttribute = class(InterfacedIniValueAttribute, IAsHex)
  strict private
    FAsHex: boolean;

    function    GetAsHex: boolean;

  public
    constructor Create(const aSection, aName: string; const aDefaultValue: string; const aAsHex: boolean); overload;

    property    AsHex: boolean read GetAsHex;

  end;


  // This attribute class can also be used as IniIntValue
  IniIntValueAttribute = class(IniNumericValueAttribute)
  public
    constructor Create(const aSection, aName: string; const aDefaultValue: integer = 0; const aAsHex: boolean = false); overload;
  end;


  // This attribute class can also be used as IniUIntValue
  IniUIntValueAttribute = class(IniNumericValueAttribute)
  public
    constructor Create(const aSection, aName: string; const aDefaultValue: cardinal = 0; const aAsHex: boolean = false); overload;
  end;


  // This attribute class can also be used as IniInt64Value
  IniInt64ValueAttribute = class(IniNumericValueAttribute)
  public
    constructor Create(const aSection, aName: string; const aDefaultValue: int64 = 0; const aAsHex: boolean = false); overload;
  end;


  // This attribute class can also be used as IniUInt64Value
  IniUInt64ValueAttribute = class(IniNumericValueAttribute)
  public
    constructor Create(const aSection, aName: string; const aDefaultValue: uint64 = 0; const aAsHex: boolean = false); overload;
  end;


  // This attribute class can also be used as IniFloatValue
  IniFloatValueAttribute = class(IniValueAttribute)
  public
    constructor Create(const aSection, aName: string; const aDefaultValue: extended = 0.0); overload;
  end;


  // This attribute class can also be used as IniDateTimeValue
  IniDateTimeValueAttribute = class(IniValueAttribute)
  public
    constructor Create(const aSection, aName: string; const aDefaultValue: string = ''); overload;
    constructor Create(const aSection, aName: string; const aDefaultValue: TDateTime = 0; const aDefaultValueIsUTC: boolean = false); overload;
  end;


  // This attribute class can also be used as IniColorValue
  IniColorValueAttribute = class(IniValueAttribute)
  public
    constructor Create(const aSection, aName: string; const aDefaultValue: TColor = clNone); overload;
  end;


  // This attribute class can also be used as IniEnumValue
  IniEnumValueAttribute = class(IniValueAttribute)
  public
    constructor Create(const aSection, aName: string; const aDefaultValue: integer = 0); overload;

    function    DefaultValue(const aTypeInfo: PTypeInfo): string; override;

  end;


  // This attribute class can also be used as IniSetValue
  IniSetValueAttribute = class(IniValueAttribute)
  public
    constructor Create(const aSection, aName: string; const aDefaultValue: string = '[]'); overload;
  end;


  // This attribute class can also be used as IniGUIDValue
  IniGUIDValueAttribute = class(IniValueAttribute)
  public
    constructor Create(const aSection, aName: string; const aDefaultValue: TGUID); overload;
  end;


  // ***************************************************************************
  // Generic INI file handling
  // ***************************************************************************

  TIniPersistence = class (TObject)
  strict private
    class function  SetValue(var aValue: TValue; const aData: string): boolean;
    class function  GetValue(const aValue: TValue; const Options: TIniStorageOptions): string;
    class function  GetIniAttribute(Obj: TRttiObject): IniValueAttribute;

  public
    class procedure Load(const FilePath: string; Obj: TObject);
    class procedure Save(const FilePath: string; Obj: TObject);

  end;


var
  IPFormatSettings: TFormatSettings;



implementation

uses
  System.IniFiles;


// *****************************************************************************
// Helper and converter functions
// *****************************************************************************

function StrSurround(const AStr, StartStr, EndStr: string): string;
begin
  Result := AStr;

  if not Result.StartsWith(StartStr) then Result := StartStr + Result;
  if not Result.EndsWith  (EndStr)   then Result := Result   + EndStr;
end;


function To8DigitHexString(const aValue: cardinal): string;
begin
  Result := Format('$%.8x', [aValue]);
end;


function To16DigitHexString(const aValue: uint64): string;
begin
  Result := Format('$%.16x', [aValue]);
end;


function ToLocalTime(const aDateTime: TDateTime): TDateTime;
begin
  Result := TTimeZone.Local.ToLocalTime(aDateTime);
end;


function ToUniversalTime(const aDateTime: TDateTime; const ForceDaylight: boolean = false): TDateTime;
begin
  Result := TTimeZone.Local.ToUniversalTime(aDateTime, ForceDaylight);
end;



// *****************************************************************************
// IniValueAttribute
// *****************************************************************************

constructor IniValueAttribute.Create(const aSection, aName: string; const aDefaultValue: string = '');
begin
  inherited Create;

  FSection      := aSection;
  FName         := aName;
  FDefaultValue := aDefaultValue;
end;


function IniValueAttribute.DefaultValue(const aTypeInfo: PTypeInfo): string;
begin
  Result := FDefaultValue;
end;



// *****************************************************************************
// InterfacedIniValueAttribute
// *****************************************************************************

function InterfacedIniValueAttribute.QueryInterface(const IID: TGUID; out Obj): HResult;
begin
  if GetInterface(IID, Obj) then
    Result := S_OK
  else
    Result := E_NOINTERFACE;
end;


function InterfacedIniValueAttribute._AddRef: Integer;
begin
  Result := -1;  // -1 indicates no reference counting is taking place
end;


function InterfacedIniValueAttribute._Release: Integer;
begin
  Result := -1;  // -1 indicates no reference counting is taking place
end;



// *****************************************************************************
// IniBoolValueAttribute
// *****************************************************************************

constructor IniBoolValueAttribute.Create(const aSection, aName: string; const aDefaultValue: boolean = false);
begin
  inherited Create(aSection, aName, BoolToStr(aDefaultValue, true));
end;



// *****************************************************************************
// IniNumericValueAttribute
// *****************************************************************************

constructor IniNumericValueAttribute.Create(const aSection, aName: string; const aDefaultValue: string; const aAsHex: boolean);
begin
  inherited Create(aSection, aName, aDefaultValue);

  FAsHex := aAsHex;
end;


function IniNumericValueAttribute.GetAsHex: boolean;
begin
  Result := FAsHex;
end;



// *****************************************************************************
// IniIntValueAttribute
// *****************************************************************************

constructor IniIntValueAttribute.Create(const aSection, aName: string; const aDefaultValue: integer = 0; const aAsHex: boolean = false);
begin
  inherited Create(aSection, aName, IntToStr(aDefaultValue), aAsHex);
end;



// *****************************************************************************
// IniUIntValueAttribute
// *****************************************************************************

constructor IniUIntValueAttribute.Create(const aSection, aName: string; const aDefaultValue: cardinal = 0; const aAsHex: boolean = false);
begin
  inherited Create(aSection, aName, UIntToStr(aDefaultValue), aAsHex);
end;



// *****************************************************************************
// IniInt64ValueAttribute
// *****************************************************************************

constructor IniInt64ValueAttribute.Create(const aSection, aName: string; const aDefaultValue: int64 = 0; const aAsHex: boolean = false);
begin
  inherited Create(aSection, aName, IntToStr(aDefaultValue), aAsHex);
end;



// *****************************************************************************
// IniUInt64ValueAttribute
// *****************************************************************************

constructor IniUInt64ValueAttribute.Create(const aSection, aName: string; const aDefaultValue: uint64 = 0; const aAsHex: boolean = false);
begin
  inherited Create(aSection, aName, UIntToStr(aDefaultValue), aAsHex);
end;



// *****************************************************************************
// IniFloatValueAttribute
// *****************************************************************************

constructor IniFloatValueAttribute.Create(const aSection, aName: string; const aDefaultValue: extended = 0.0);
begin
  inherited Create(aSection, aName, FloatToStr(aDefaultValue, IPFormatSettings));
end;



// *****************************************************************************
// IniDateTimeValueAttribute
// *****************************************************************************

constructor IniDateTimeValueAttribute.Create(const aSection, aName: string; const aDefaultValue: string = '');
begin
  if aDefaultValue <> '' then
    Create(aSection, aName, ISO8601ToDate(aDefaultValue), true)
  else
    Create(aSection, aName, 0, false);
end;


constructor IniDateTimeValueAttribute.Create(const aSection, aName: string; const aDefaultValue: TDateTime = 0; const aDefaultValueIsUTC: boolean = false);
begin
  if not aDefaultValueIsUTC then
    inherited Create(aSection, aName, DateToISO8601(ToUniversalTime(aDefaultValue)))
  else
    inherited Create(aSection, aName, DateToISO8601(aDefaultValue));
end;



// *****************************************************************************
// IniColorValueAttribute
// *****************************************************************************

constructor IniColorValueAttribute.Create(const aSection, aName: string; const aDefaultValue: TColor = clNone);
begin
  inherited Create(aSection, aName, To8DigitHexString(ColorToRGB(aDefaultValue)));
end;



// *****************************************************************************
// IniEnumValueAttribute
// *****************************************************************************

constructor IniEnumValueAttribute.Create(const aSection, aName: string; const aDefaultValue: integer = 0);
begin
  inherited Create(aSection, aName, IntToStr(aDefaultValue));
end;


function IniEnumValueAttribute.DefaultValue(const aTypeInfo: PTypeInfo): string;
begin
  Result := GetEnumName(aTypeInfo, StrToIntDef(FDefaultValue, 0));
end;



// *****************************************************************************
// IniSetValueAttribute
// *****************************************************************************

constructor IniSetValueAttribute.Create(const aSection, aName: string; const aDefaultValue: string = '[]');
begin
  inherited Create(aSection, aName, StrSurround(aDefaultValue, '[', ']'));
end;



// *****************************************************************************
// IniGUIDValueAttribute
// *****************************************************************************

constructor IniGUIDValueAttribute.Create(const aSection, aName: string; const aDefaultValue: TGUID);
begin
  inherited Create(aSection, aName, aDefaultValue.ToString);
end;



// *****************************************************************************
// TIniPersistence
// *****************************************************************************

class procedure TIniPersistence.Load(const FilePath: string; Obj: TObject);
var
  ctx:      TRttiContext;
  objType:  TRttiType;
  Field:    TRttiField;
  Prop:     TRttiProperty;
  IniValue: IniValueAttribute;
  Value:    TValue;
  Ini:      TIniFile;
  Data:     string;

begin
  ctx := TRttiContext.Create;

  try
    Ini := TIniFile.Create(FilePath);

    try
      objType := ctx.GetType(Obj.ClassInfo);

      // Load values of properties
      for Prop in objType.GetProperties do
      begin
        IniValue := GetIniAttribute(Prop);

        if not Assigned(IniValue) then
          continue;

        Value := Prop.GetValue(Obj);
        Data  := Ini.ReadString(IniValue.Section, IniValue.Name, IniValue.DefaultValue(Value.TypeInfo));

        if SetValue(Value, Data) then
          Prop.SetValue(Obj, Value);
      end;

      // Load values of field variables
      for Field in objType.GetFields do
      begin
        IniValue := GetIniAttribute(Field);

        if not Assigned(IniValue) then
          continue;

        Value := Field.GetValue(Obj);
        Data  := Ini.ReadString(IniValue.Section, IniValue.Name, IniValue.DefaultValue(Value.TypeInfo));

        if SetValue(Value, Data) then
          Field.SetValue(Obj, Value);
      end;

    finally
      Ini.Free;
    end;

  finally
    ctx.Free;
  end;
end;


class procedure TIniPersistence.Save(const FilePath: string; Obj: TObject);
var
  ctx:         TRttiContext;
  objType:     TRttiType;
  Field:       TRttiField;
  Prop:        TRttiProperty;
  IniValue:    IniValueAttribute;
  Value:       TValue;
  IniValueIf:  IAsHex;
  Options:     TIniStorageOptions;
  Ini:         TIniFile;
  IniSections: TStringList;
  IniSection:  string;
  Data:        string;

begin
  ctx := TRttiContext.Create;

  try
    Ini         := TIniFile.Create(FilePath);
    IniSections := TStringList.Create;

    try
      // Clear content of INI file
      Ini.ReadSections(IniSections);

      for IniSection in IniSections do
        Ini.EraseSection(IniSection);

      objType := ctx.GetType(Obj.ClassInfo);

      // Save values of properties
      for Prop in objType.GetProperties do
      begin
        IniValue := GetIniAttribute(Prop);
        Options  := [];

        if Assigned(IniValue) then
        begin
          if IniValue is InterfacedIniValueAttribute then
          begin
            if Supports(IniValue, IAsHex, IniValueIf) and
               IniValueIf.AsHex                       then
              Include(Options, isoAsHex);
          end;

          Value := Prop.GetValue(Obj);
          Data  := GetValue(Value, Options);
          Ini.WriteString(IniValue.Section, IniValue.Name, Data)
        end;
      end;

      // Save values of field variables
      for Field in objType.GetFields do
      begin
        IniValue := GetIniAttribute(Field);
        Options  := [];

        if Assigned(IniValue) then
        begin
          if IniValue is InterfacedIniValueAttribute then
          begin
            if Supports(IniValue, IAsHex, IniValueIf) and
               IniValueIf.AsHex                       then
              Include(Options, isoAsHex);
          end;

          Value := Field.GetValue(Obj);
          Data  := GetValue(Value, Options);
          Ini.WriteString(IniValue.Section, IniValue.Name, Data);
        end;
      end;

    finally
      IniSections.Free;
      Ini.Free;
    end;

  finally
    ctx.Free;
  end;
end;


class function TIniPersistence.GetIniAttribute(Obj: TRttiObject): IniValueAttribute;
var
  Attr: TCustomAttribute;

begin
  for Attr in Obj.GetAttributes do
  begin
    if Attr is IniValueAttribute then
      exit(IniValueAttribute(Attr));
  end;

  Result := nil;
end;


// Called when a string read from the INI file has to be converted to a certain value
class function TIniPersistence.SetValue(var aValue: TValue; const aData: string): boolean;
var
  ASet:  TBytes;
  AGUID: TGUID;

begin
  try
    case aValue.Kind of
      tkWChar,
      tkLString,
      tkWString,
      tkString,
      tkChar,
      tkUString:
      begin
        aValue := aData;
        exit(true);
      end;

      tkInteger:
      begin
        try    aValue := StrToInt(aData);
        except aValue := StrToUInt(aData); end;
        exit(true);
      end;

      tkInt64:
      begin
        try    aValue := StrToInt64(aData);
        except aValue := StrToUInt64(aData); end;
        exit(true);
      end;

      tkFloat:
      begin
        // Special treatment of TDateTime values
        if aValue.TypeInfo.Name = PTypeInfo(TypeInfo(TDateTime)).Name then
          aValue := ToLocalTime(ISO8601ToDate(aData))
        else
          aValue := StrToFloat(aData, IPFormatSettings);

        exit(true);
      end;

      tkEnumeration:
      begin
        aValue := TValue.FromOrdinal(aValue.TypeInfo, GetEnumValue(aValue.TypeInfo, aData));
        exit(true);
      end;

      tkSet:
      begin
        SetLength(ASet, SizeOfSet(aValue.TypeInfo));
        StringToSet(aValue.TypeInfo, StrSurround(aData, '[', ']'), ASet);
        TValue.Make(ASet, aValue.TypeInfo, aValue);
        exit(true);
      end;

      tkRecord:
        // Only TGUID is supported
        if AValue.TypeInfo.Name = PTypeInfo(TypeInfo(TGUID)).Name then
        begin
          AGuid := StringToGUID(AData);
          TValue.Make(@AGuid, AValue.TypeInfo, AValue);
          exit(true);
        end;
    end;

  except
    on E: Exception do
    begin
      MessageDlg(Format('%s: %s', [ClassName, E.Message]), mtError, [mbOK], 0);
      exit(false);
    end;
  end;

  raise EConvertError.CreateFmt('%s: Data type [%s] not supported', [ClassName, aValue.TypeInfo.Name]);
end;


// Called when a certain value has to be converted to string in order to write
// it to the INI file
class function TIniPersistence.GetValue(const aValue: TValue; const Options: TIniStorageOptions): string;
begin
  case aValue.Kind of
    tkWChar,
    tkLString,
    tkWString,
    tkString,
    tkChar,
    tkUString,
    tkEnumeration,
    tkSet:
    begin
      Result := aValue.ToString;
      exit;
    end;

    tkInteger:
    begin
      // Special treatment for TColor values
      if aValue.TypeInfo.Name = PTypeInfo(TypeInfo(TColor)).Name then
        Result := To8DigitHexString(ColorToRGB(aValue.AsType<TColor>))

      // Check if value should be stored as hex string
      else if isoAsHex in Options then
        Result := To8DigitHexString(aValue.AsType<cardinal>)
      else
        Result := aValue.ToString;

      exit;
    end;

    tkInt64:
    begin
      // Check if value should be stored as hex string
      if isoAsHex in Options then
        Result := To16DigitHexString(aValue.AsUInt64)
      else
        Result := aValue.ToString;

      exit;
    end;

    tkFloat:
    begin
      // Special treatment for TDateTime values
      if aValue.TypeInfo.Name = PTypeInfo(TypeInfo(TDateTime)).Name then
        Result := DateToISO8601(ToUniversalTime(aValue.AsExtended))
      else
        Result := FloatToStr(aValue.AsExtended, IPFormatSettings);

      exit;
    end;

    tkRecord:
      // Only TGUID is supported
      if AValue.TypeInfo.Name = PTypeInfo(TypeInfo(TGUID)).Name then
      begin
        Result := AValue.AsType<TGUID>.ToString();
        exit;
      end
  end;

  raise EConvertError.CreateFmt('%s: Data type [%s] not supported', [ClassName, aValue.TypeInfo.Name]);
end;



initialization

IPFormatSettings := TFormatSettings.Invariant;


end.

