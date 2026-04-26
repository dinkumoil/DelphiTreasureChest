// MIT License
//
// Based on ideas of Robert Love, 2009
// Extended: Andreas Heim, 2015 - 2026
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
  System.SysUtils, System.DateUtils, System.Math, Soap.XSBuiltIns, System.UITypes,
  System.Classes, System.TypInfo, System.Rtti, System.IniFiles, Vcl.Graphics, Vcl.Dialogs;


type
  // ***************************************************************************
  // Forward declarations
  // ***************************************************************************

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


  // Attribute class to specify how much elements a property or field of a dynamic array type should have
  IniArrayLengthAttribute = class(TCustomAttribute)
  strict protected
    FLength: NativeInt;

  public
    constructor Create(const aLength: NativeInt = 0);

    property    Length: NativeInt read FLength write FLength;

  end;


  // Base class for all attribute classes supporting interfaces, should NOT be used directly
  InterfacedIniValueAttribute = class(IniValueAttribute, IInterface)
  strict protected
    // Methods of IInterface, reference counting is deactivated
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;

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


  // This attribute class can also be used as IniStrValue
  IniStrValueAttribute = class(IniValueAttribute)
  end;


  // This attribute class can also be used as IniBoolValue
  IniBoolValueAttribute = class(IniValueAttribute)
  public
    constructor Create(const aSection, aName: string; const aDefaultValue: boolean = false); overload;
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
  strict private const
    cErrMsg_ClassNoRTTI           = '%s: Class type [%s] has no runtime type information for class member [%s].';
    cErrMsg_ArrayNoRTTI           = '%s: Class member [%s] is of array type [%s] that has no runtime type information.';
    cErrMsg_ArrayElementNoRTTI    = '%s: Class member [%s] is of array type [%s] whose element type has no runtime type information.';
    cErrMsg_ArrayNoDims           = '%s: Class member [%s] is of array type [%s] which has no dimensions.';
    cErrMsg_ArrayMultipleDims     = '%s: Class member [%s] is of array type [%s] which has more than one dimension.';
    cErrMsg_ArrayInplaceIndexType = '%s: Class member [%s] is of array type [%s] which must not have inplace type declaration for indices.';
    cErrMsg_ArrayTypeUnsupported  = '%s: Class member [%s] is of data type [%s] which is an unsupported array type.';
    cErrMsg_UnsupportedDataType   = '%s: Data type [%s] not supported.';

  strict private
    class procedure Deserialize<T: TRttiDataMember>(aIni: TIniFile; aRttiCtx: TRttiContext; aObj: TObject; aObjMember: T; aAttr: IniValueAttribute; aArrayLength: NativeInt);
    class procedure Serialize<T: TRttiDataMember>(aIni: TIniFile; aRttiCtx: TRttiContext; aObj: TObject; aObjMember: T; aAttr: IniValueAttribute);

    class function  GetIniAttribute(aObj: TRttiObject; out aArrayLength: NativeInt): IniValueAttribute;
    class function  SetValue(var aValue: TValue; const aData: string): boolean;
    class function  GetValue(const aValue: TValue; const aOptions: TIniStorageOptions): string;

  public
    class procedure Load(const aFilePath: string; aObj: TObject);
    class procedure Save(const aFilePath: string; aObj: TObject);

  end;


var
  IPFormatSettings: TFormatSettings;



implementation

// *****************************************************************************
// Helper and converter functions
// *****************************************************************************

function StrSurround(const aStr, aStartStr, aEndStr: string): string;
begin
  Result := aStr;

  if not Result.StartsWith(aStartStr) then Result := aStartStr + Result;
  if not Result.EndsWith  (aEndStr)   then Result := Result   + aEndStr;
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


function ToUniversalTime(const aDateTime: TDateTime; const aForceDaylight: boolean = false): TDateTime;
begin
  Result := TTimeZone.Local.ToUniversalTime(aDateTime, aForceDaylight);
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
// IniArrayLengthAttribute
// *****************************************************************************

constructor IniArrayLengthAttribute.Create(const aLength: NativeInt = 0);
begin
  inherited Create;

  FLength := aLength;
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
// IniBoolValueAttribute
// *****************************************************************************

constructor IniBoolValueAttribute.Create(const aSection, aName: string; const aDefaultValue: boolean = false);
begin
  inherited Create(aSection, aName, BoolToStr(aDefaultValue, true));
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
    Create(aSection, aName, XMLTimeToDateTime(aDefaultValue, true), true)
  else
    Create(aSection, aName, 0, false);
end;


constructor IniDateTimeValueAttribute.Create(const aSection, aName: string; const aDefaultValue: TDateTime = 0; const aDefaultValueIsUTC: boolean = false);
begin
  if not aDefaultValueIsUTC then
    inherited Create(aSection, aName, DateTimeToXMLTime(ToUniversalTime(aDefaultValue), false))
  else
    inherited Create(aSection, aName, DateTimeToXMLTime(aDefaultValue, false));
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

class procedure TIniPersistence.Load(const aFilePath: string; aObj: TObject);
var
  RttiCtx:     TRttiContext;
  ObjType:     TRttiType;
  Field:       TRttiField;
  Prop:        TRttiProperty;
  Attr:        IniValueAttribute;
  Ini:         TIniFile;
  ArrayLength: NativeInt;

begin
  RttiCtx := TRttiContext.Create;

  try
    Ini := TIniFile.Create(aFilePath);

    try
      ObjType := RttiCtx.GetType(aObj.ClassInfo);

      // Load values of properties
      for Prop in ObjType.GetProperties do
      begin
        Attr := GetIniAttribute(Prop, ArrayLength);

        if Assigned(Attr) then
          Deserialize<TRttiProperty>(Ini, RttiCtx, aObj, Prop, Attr, ArrayLength);
      end;

      // Load values of field variables
      for Field in ObjType.GetFields do
      begin
        Attr := GetIniAttribute(Field, ArrayLength);

        if Assigned(Attr) then
          Deserialize<TRttiField>(Ini, RttiCtx, aObj, Field, Attr, ArrayLength);
      end;

    finally
      Ini.Free;
    end;

  finally
    RttiCtx.Free;
  end;
end;


class procedure TIniPersistence.Save(const aFilePath: string; aObj: TObject);
var
  RttiCtx:     TRttiContext;
  ObjType:     TRttiType;
  Field:       TRttiField;
  Prop:        TRttiProperty;
  Attr:        IniValueAttribute;
  Ini:         TIniFile;
  IniSections: TStringList;
  IniSection:  string;
  ArrayLength: NativeInt;

begin
  RttiCtx := TRttiContext.Create;

  try
    Ini         := TIniFile.Create(aFilePath);
    IniSections := TStringList.Create;

    try
      // Clear content of INI file
      Ini.ReadSections(IniSections);

      for IniSection in IniSections do
        Ini.EraseSection(IniSection);

      ObjType := RttiCtx.GetType(aObj.ClassInfo);

      // Save values of properties
      for Prop in ObjType.GetProperties do
      begin
        Attr := GetIniAttribute(Prop, ArrayLength);

        if Assigned(Attr) then
          Serialize<TRttiProperty>(Ini, RttiCtx, aObj, Prop, Attr);
      end;

      // Save values of field variables
      for Field in ObjType.GetFields do
      begin
        Attr := GetIniAttribute(Field, ArrayLength);

        if Assigned(Attr) then
          Serialize<TRttiField>(Ini, RttiCtx, aObj, Field, Attr);
      end;

    finally
      IniSections.Free;
      Ini.Free;
    end;

  finally
    RttiCtx.Free;
  end;
end;


class function TIniPersistence.GetIniAttribute(aObj: TRttiObject; out aArrayLength: NativeInt): IniValueAttribute;
var
  Attr: TCustomAttribute;

begin
  Result       := nil;
  aArrayLength := 0;

  for Attr in aObj.GetAttributes do
  begin
    if Attr is IniValueAttribute then
      Result := IniValueAttribute(Attr)

    else if Attr is IniArrayLengthAttribute then
      aArrayLength := IniArrayLengthAttribute(Attr).Length;
  end;
end;


// Called to read a value from INI file and store it into variable or property of class instance
class procedure TIniPersistence.Deserialize<T>(aIni: TIniFile; aRttiCtx: TRttiContext; aObj: TObject; aObjMember: T; aAttr: IniValueAttribute; aArrayLength: NativeInt);
var
  Options:    TIniStorageOptions;
  ArrType:    TRttiArrayType;
  DynArrType: TRttiDynamicArrayType;
  IndexType:  TRttiOrdinalType;
  ElemType:   TRttiType;
  Value:      TValue;
  ElemValue:  TValue;
  ElemCnt:    NativeInt;
  ValueList:  TStringList;
  Idx:        integer;
  MinIdx:     integer;
  MaxIdx:     integer;
  Data:       string;

begin
  Options := [];

  // Get TValue holding variable/property type info
  // and buffer to hold the actual data later on
  Value := aObjMember.GetValue(aObj);

  if not Assigned(Value.TypeInfo) then
    raise ENotSupportedException.CreateFmt(cErrMsg_ClassNoRTTI, [ClassName, aObj.ClassName, aObjMember.Name]);

  if not Value.IsArray then
  begin
    // Process variable/property marked as scalar type
    Data := aIni.ReadString(aAttr.Section, aAttr.Name, aAttr.DefaultValue(Value.TypeInfo));

    // Transform string value read from INI file into TValue and
    // write TValue into variable/property of class instance
    if SetValue(Value, Data) then      // Data conversion from string to actual type
      aObjMember.SetValue(aObj, Value);  // Transfer converted data into class instance
  end
  else
  begin
    // Process variable/property marked as array type
    MinIdx := 0;
    MaxIdx := 0;

    // Get type info of array element type in order to determine its default value
    // For static array additionally get its min and max indices
    case Value.TypeInfo.Kind of
      tkArray:
      begin
        ArrType := aRttiCtx.GetType(Value.TypeInfo) as TRttiArrayType;

        if not Assigned(ArrType) then
          raise ENotSupportedException.CreateFmt(cErrMsg_ArrayNoRTTI, [ClassName, aObjMember.Name, Value.TypeInfo.Name]);

        ElemType := ArrType.ElementType;

        if not Assigned(ElemType) then
          raise ENotSupportedException.CreateFmt(cErrMsg_ArrayElementNoRTTI, [ClassName, aObjMember.Name, Value.TypeInfo.Name]);

        if ArrType.DimensionCount <= 0 then
          raise ENotSupportedException.CreateFmt(cErrMsg_ArrayNoDims, [ClassName, aObjMember.Name, Value.TypeInfo.Name]);

        if ArrType.DimensionCount > 1 then
          raise ENotSupportedException.CreateFmt(cErrMsg_ArrayMultipleDims, [ClassName, aObjMember.Name, Value.TypeInfo.Name]);

        IndexType := ArrType.Dimensions[0] as TRttiOrdinalType;

        if not Assigned(IndexType) then
          raise ENotSupportedException.CreateFmt(cErrMsg_ArrayInplaceIndexType, [ClassName, aObjMember.Name, Value.TypeInfo.Name]);

        MinIdx := IndexType.MinValue;
        MaxIdx := IndexType.MaxValue;
      end;

      tkDynArray:
      begin
        DynArrType := aRttiCtx.GetType(Value.TypeInfo) as TRttiDynamicArrayType;

        if not Assigned(DynArrType) then
          raise ENotSupportedException.CreateFmt(cErrMsg_ArrayNoRTTI, [ClassName, aObjMember.Name, Value.TypeInfo.Name]);

        ElemType := DynArrType.ElementType;
      end;

      else
        raise ENotSupportedException.CreateFmt(cErrMsg_ArrayTypeUnsupported, [ClassName, aObjMember.Name, Value.TypeInfo.Name]);
    end;

    ValueList := TStringList.Create;

    try
      ValueList.Duplicates      := dupAccept;
      ValueList.QuoteChar       := #0;
      ValueList.Delimiter       := #0;
      ValueList.StrictDelimiter := true;
      ValueList.CaseSensitive   := false;
      ValueList.Sorted          := false;

      // Read indexed values from INI file into string list
      Idx := 1;  // Indices of INI file indexed values have to be 1-based and contiguously numbered

      while aIni.ValueExists(aAttr.Section, aAttr.Name + IntToStr(Idx)) do
      begin
        Data := aIni.ReadString(aAttr.Section, aAttr.Name + IntToStr(Idx), aAttr.DefaultValue(ElemType.Handle));
        ValueList.Add(Data);

        Inc(Idx);
      end;

      // Transfer read values into data buffer of TValue
      case Value.TypeInfo.Kind of
        tkArray:
        begin
          // Fill all elements of static array using default values if necessary
          // Note: Memory of static array has been already allocated by compiler
          for Idx := MinIdx to MaxIdx do
          begin
            ElemValue := Value.GetArrayElement(Idx);

            if Idx - MinIdx < ValueList.Count then
              Data := ValueList[Idx - MinIdx]
            else
              Data := aAttr.DefaultValue(ElemType.Handle);

            if SetValue(ElemValue, Data) then         // Data conversion from string to actual type
              Value.SetArrayElement(Idx, ElemValue);  // Transfer type info and converted data into TValue
          end;
        end;

        tkDynArray:
        begin
          // Fill up dynamic array to requested number of elements using default values if necessary
          // Note: Memory of dynamic array has to be allocated manually right now
          ElemCnt := Max(aArrayLength, ValueList.Count);
          DynArraySetLength(PPointer(Value.GetReferenceToRawData)^, Value.TypeInfo, 1, @ElemCnt);

          for Idx := 0 to Pred(ElemCnt) do
          begin
            ElemValue := Value.GetArrayElement(Idx);

            if Idx < ValueList.Count then
              Data := ValueList[Idx]
            else
              Data := aAttr.DefaultValue(ElemType.Handle);

            if SetValue(ElemValue, Data) then         // Data conversion from string to actual type
              Value.SetArrayElement(Idx, ElemValue);  // Transfer type info and converted data into TValue
          end;
        end;
      end;

    finally
      ValueList.Free;
    end;

    // Transfer content of TValue's data buffer into class instance
    aObjMember.SetValue(aObj, Value);
  end;
end;


// Called to get the value of a variable or property of a class instance and write it to INI file
class procedure TIniPersistence.Serialize<T>(aIni: TIniFile; aRttiCtx: TRttiContext; aObj: TObject; aObjMember: T; aAttr: IniValueAttribute);
var
  Options:   TIniStorageOptions;
  ArrType:   TRttiArrayType;
  IndexType: TRttiOrdinalType;
  AttrHexIf: IAsHex;
  Value:     TValue;
  ElemValue: TValue;
  ElemCnt:   NativeInt;
  Idx:       integer;
  MinIdx:    integer;
  MaxIdx:    integer;
  Data:      string;

begin
  Options := [];

  // Read processing options of variable/property
  if aAttr is InterfacedIniValueAttribute then
  begin
    if Supports(aAttr, IAsHex, AttrHexIf) and
       AttrHexIf.AsHex                    then
      Include(Options, isoAsHex);
  end;

  // Get TValue holding variable/property type info
  // and buffer that holds the actual data
  Value := aObjMember.GetValue(aObj);

  if not Assigned(Value.TypeInfo) then
    raise ENotSupportedException.CreateFmt(cErrMsg_ClassNoRTTI, [ClassName, aObj.ClassName, aObjMember.Name]);

  if not Value.IsArray then
  begin
    // Process variable/property marked as scalar type
    // Data conversion from actual type to string
    Data := GetValue(Value, Options);

    // Write value to INI file
    aIni.WriteString(aAttr.Section, aAttr.Name, Data);
  end
  else
  begin
    // Process variable/property marked as array type
    MinIdx := 0;
    MaxIdx := 0;

    // Get type info of static array in order to get its min and max indices
    case Value.TypeInfo.Kind of
      tkArray:
      begin
        ArrType := aRttiCtx.GetType(Value.TypeInfo) as TRttiArrayType;

        if not Assigned(ArrType) then
          raise ENotSupportedException.CreateFmt(cErrMsg_ArrayNoRTTI, [ClassName, aObjMember.Name, Value.TypeInfo.Name]);

        if not Assigned(ArrType.ElementType) then
          raise ENotSupportedException.CreateFmt(cErrMsg_ArrayElementNoRTTI, [ClassName, aObjMember.Name, Value.TypeInfo.Name]);

        if ArrType.DimensionCount <= 0 then
          raise ENotSupportedException.CreateFmt(cErrMsg_ArrayNoDims, [ClassName, aObjMember.Name, Value.TypeInfo.Name]);

        if ArrType.DimensionCount > 1 then
          raise ENotSupportedException.CreateFmt(cErrMsg_ArrayMultipleDims, [ClassName, aObjMember.Name, Value.TypeInfo.Name]);

        IndexType := ArrType.Dimensions[0] as TRttiOrdinalType;

        if not Assigned(IndexType) then
          raise ENotSupportedException.CreateFmt(cErrMsg_ArrayInplaceIndexType, [ClassName, aObjMember.Name, Value.TypeInfo.Name]);

        MinIdx := IndexType.MinValue;
        MaxIdx := IndexType.MaxValue;
      end;

      tkDynArray:
      begin
        // Nothing to do
      end;

      else
        raise ENotSupportedException.CreateFmt(cErrMsg_ArrayTypeUnsupported, [ClassName, aObjMember.Name, Value.TypeInfo.Name]);
    end;

    // Write value from class instance to INI file
    case Value.TypeInfo.Kind of
      tkArray:
      begin
        for Idx := MinIdx to MaxIdx do
        begin
          ElemValue := Value.GetArrayElement(Idx);    // Transfer type info and actual data into TValue
          Data      := GetValue(ElemValue, Options);  // Data conversion from actual type to string

          aIni.WriteString(aAttr.Section, aAttr.Name + IntToStr(Succ(Idx - MinIdx)), Data);
        end;
      end;

      tkDynArray:
      begin
        ElemCnt := Value.GetArrayLength;

        for Idx := 0 to Pred(ElemCnt) do
        begin
          ElemValue := Value.GetArrayElement(Idx);    // Transfer type info and actual data into TValue
          Data      := GetValue(ElemValue, Options);  // Data conversion from actual type to string

          aIni.WriteString(aAttr.Section, aAttr.Name + IntToStr(Succ(Idx)), Data);
        end;
      end;
    end;
  end;
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
          aValue := ToLocalTime(XMLTimeToDateTime(aData, true))
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
          if aData.IsEmpty then
            AGuid := TGUID.Empty
          else
            AGuid := StringToGUID(aData);

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

  raise ENotSupportedException.CreateFmt(cErrMsg_UnsupportedDataType, [ClassName, aValue.TypeInfo.Name]);
end;


// Called when a certain value has to be converted to string in order to write
// it to the INI file
class function TIniPersistence.GetValue(const aValue: TValue; const aOptions: TIniStorageOptions): string;
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
      else if isoAsHex in aOptions then
        Result := To8DigitHexString(aValue.AsType<cardinal>)
      else
        Result := aValue.ToString;

      exit;
    end;

    tkInt64:
    begin
      // Check if value should be stored as hex string
      if isoAsHex in aOptions then
        Result := To16DigitHexString(aValue.AsUInt64)
      else
        Result := aValue.ToString;

      exit;
    end;

    tkFloat:
    begin
      // Special treatment for TDateTime values
      if aValue.TypeInfo.Name = PTypeInfo(TypeInfo(TDateTime)).Name then
        Result := DateTimeToXMLTime(ToUniversalTime(aValue.AsExtended), false)
      else
        Result := FloatToStr(aValue.AsExtended, IPFormatSettings);

      exit;
    end;

    tkRecord:
      // Only TGUID is supported
      if aValue.TypeInfo.Name = PTypeInfo(TypeInfo(TGUID)).Name then
      begin
        Result := aValue.AsType<TGUID>.ToString();
        exit;
      end
  end;

  raise ENotSupportedException.CreateFmt(cErrMsg_UnsupportedDataType, [ClassName, aValue.TypeInfo.Name]);
end;



initialization

IPFormatSettings := TFormatSettings.Invariant;


end.

