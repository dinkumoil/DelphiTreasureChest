unit Generics.Utils;


interface

uses
  System.SysUtils, System.Math, System.TypInfo, System.Classes, System.Generics.Defaults,
  System.Generics.Collections;


type
// =============================================================================
// TArray
// =============================================================================

   TArrayHelper = class helper for System.Generics.Collections.TArray
     class function Construct<T>(Count: Integer; const InitialValue: T): TArray<T>;
     class function Contains<T>(const Values: array of T; const Item: T): boolean;
     class function IndexOf<T>(const Values: array of T; const Item: T): integer;
   end;


// =============================================================================
// TGeneric
// =============================================================================

  TGeneric = record
  public
    class function IfThen<T>(Condition: boolean; const Op1, Op2: T): T; static;
  end;


// =============================================================================
// TEnum
// =============================================================================

  TEnum = record
  strict private const
    ERR_MSG_NO_TYPE_INFO          = 'No runtime type information available for this data type.';
    ERR_MSG_NOT_AN_ENUM_TYPE      = 'Data type %s is not an enum type.';
    ERR_MSG_NOT_AN_ENUM_VALUE     = '%d is not a valid value of enum %s.';
    ERR_MSG_UNSUPPORTED_ENUM_SIZE = 'Enum %s has an unsupported size.';

  private
    class procedure CheckIfEnum<T: record>; static;
    class procedure CheckRange<T: record>(const Value: integer); static;

  public
    class function AsString<T: record>(const AEnum: T): string; static;
    class function AsInteger<T: record>(const AEnum: T): Integer; static;

    class function FromString<T: record>(const Value: string): T; static;
    class function FromInteger<T: record>(const Value: integer): T; static;

    class function FromStringDef<T: record>(const Value: string; const DefaultValue: T): T; static;
    class function FromIntegerDef<T: record>(const Value: integer; const DefaultValue: T): T; static;
  end;


// =============================================================================
// TSet
// =============================================================================

  TSet = record
  strict private const
    ERR_MSG_NO_TYPE_INFO   = 'No runtime type information available for this data type.';
    ERR_MSG_NOT_A_SET_TYPE = 'Data type %s is not a set type.';

    FSetBitsOfValue: packed array[0..15] of byte = (0,1,1,2,1,2,2,3,1,2,2,3,2,3,3,4);

  private
    class procedure CheckIfSet<T>; static;

  public
    class function CountItems<T>(const ASet: T): integer; static;

    class function AsString<T; TI: record>(const ASet: T): string; static;
    class function AsArray<T; TI: record>(const ASet: T): TArray<TI>; static;
    class function AsStringArray<T; TI: record>(const ASet: T): TArray<string>; static;

    class function FromString<T; TI: record>(const ASetAsStr: string): T; static;
    class function FromArray<T; TI: record>(const AArr: TArray<TI>): T; static;
    class function FromStringArray<T; TI: record>(const AStrArr: TArray<string>): T; static;
  end;



implementation

// =============================================================================
// TArray
// =============================================================================

class function TArrayHelper.Construct<T>(Count: Integer; const InitialValue: T): TArray<T>;
var
  I: integer;

begin
  SetLength(Result, Count);

  for I := 0 to High(Result) do
    Result[I] := InitialValue;
end;


class function TArrayHelper.Contains<T>(const Values: array of T; const Item: T): boolean;
begin
  Result := (IndexOf(Values, Item) >= 0);
end;


class function TArrayHelper.IndexOf<T>(const Values: array of T; const Item: T): integer;
var
  I:   integer;
  Cmp: IEqualityComparer<T>;

begin
  Result := -1;
  Cmp    := TEqualityComparer<T>.Default;

  if not Assigned(Cmp) then
    exit;

  for I := Low(Values) to High(Values) do
    if Cmp.Equals(Item, Values[I]) then
      exit(I);
end;



// =============================================================================
// TGeneric
// =============================================================================

class function TGeneric.IfThen<T>(Condition: boolean; const Op1, Op2: T): T;
begin
  if Condition then exit(Op1)
  else              exit(Op2);
end;



// =============================================================================
// TEnum
// =============================================================================

class procedure TEnum.CheckIfEnum<T>;
begin
  try
    if PTypeInfo(TypeInfo(T)).Kind <> tkEnumeration then
      raise EConvertError.CreateFmt(ERR_MSG_NOT_AN_ENUM_TYPE, [PTypeInfo(TypeInfo(T)).Name]);
  except
    on EConvertError do
      raise;
    else
      raise EConvertError.Create(ERR_MSG_NO_TYPE_INFO);
  end;
end;


class procedure TEnum.CheckRange<T>(const Value: integer);
var
  TypData: PTypeData;

begin
  TypData := GetTypeData(TypeInfo(T));

  if not InRange(Value, TypData.MinValue, TypData.MaxValue) then
    raise EConvertError.CreateFmt(ERR_MSG_NOT_AN_ENUM_VALUE, [Value, PTypeInfo(TypeInfo(T)).Name]);
end;


class function TEnum.AsString<T>(const AEnum: T): string;
begin
  Result := GetEnumName(TypeInfo(T), AsInteger(AEnum));
end;


class function TEnum.AsInteger<T>(const AEnum: T): Integer;
begin
  CheckIfEnum<T>();

  case SizeOf(T) of
    1:   Result := PByte(@AEnum)^;
    2:   Result := PWord(@AEnum)^;
    4:   Result := PCardinal(@AEnum)^;
    else raise EConvertError.CreateFmt(ERR_MSG_UNSUPPORTED_ENUM_SIZE, [PTypeInfo(TypeInfo(T)).Name]);
  end;

  CheckRange<T>(Result);
end;


class function TEnum.FromString<T>(const Value: string): T;
begin
  CheckIfEnum<T>();

  Result := FromInteger<T>(GetEnumValue(TypeInfo(T), Value));
end;


class function TEnum.FromInteger<T>(const Value: integer): T;
begin
  CheckIfEnum<T>();
  CheckRange<T>(Value);

  case SizeOf(T) of
    1:   PByte(@Result)^     := Byte(Value);
    2:   PWord(@Result)^     := Word(Value);
    4:   PCardinal(@Result)^ := Cardinal(Value);
    else raise EConvertError.CreateFmt(ERR_MSG_UNSUPPORTED_ENUM_SIZE, [PTypeInfo(TypeInfo(T)).Name]);
  end;
end;


class function TEnum.FromStringDef<T>(const Value: string; const DefaultValue: T): T;
var
  Dummy: integer;

begin
  Dummy := AsInteger<T>(DefaultValue);

  try
    Result := FromString<T>(Value);
  except
    Result := DefaultValue;
  end;
end;


class function TEnum.FromIntegerDef<T>(const Value: integer; const DefaultValue: T): T;
var
  Dummy: integer;

begin
  Dummy := AsInteger<T>(DefaultValue);

  try
    Result := FromInteger<T>(Value);
  except
    Result := DefaultValue;
  end;
end;



// =============================================================================
// TSet
// =============================================================================

class procedure TSet.CheckIfSet<T>;
begin
  try
    if PTypeInfo(TypeInfo(T)).Kind <> tkSet then
      raise EConvertError.CreateFmt(ERR_MSG_NOT_A_SET_TYPE, [PTypeInfo(TypeInfo(T)).Name]);
  except
    on EConvertError do
      raise;
    else
      raise EConvertError.Create(ERR_MSG_NO_TYPE_INFO);
  end;
end;


class function TSet.CountItems<T>(const ASet: T): integer;
var
  I: integer;
  B: byte;

begin
  Result := 0;

  CheckIfSet<T>();

  for I := 0 to Pred(SizeOf(T)) do
  begin
    B := PByte(@ASet)[I];

    Inc(Result, FSetBitsOfValue[B and $0F]);
    Inc(Result, FSetBitsOfValue[B shr 4]);
  end;
end;


class function TSet.AsString<T, TI>(const ASet: T): string;
var
  Items: TArray<TI>;
  I:     integer;

begin
  Result := '[';

  try
    CheckIfSet<T>();
    TEnum.CheckIfEnum<TI>();

    Items := AsArray<T, TI>(ASet);

    for I := 0 to Pred(Length(Items)) do
      if I = 0 then
        Result := Result + TEnum.AsString<TI>(Items[I])
      else
        Result := Result + ', ' + TEnum.AsString<TI>(Items[I]);

  finally
    Result := Result + ']'
  end;
end;


class function TSet.AsArray<T, TI>(const ASet: T): TArray<TI>;
var
  I: integer;
  J: integer;
  K: integer;
  V: integer;
  B: byte;

begin
  Result := TArray<TI>.Create();

  CheckIfSet<T>();
  TEnum.CheckIfEnum<TI>();

  SetLength(Result, CountItems(ASet));

  K := 0;
  V := 0;

  for I := 0 to Pred(SizeOf(T)) do
  begin
    B := PByte(@ASet)[I];

    for J := 1 to 8 do
    begin
      if B and $01 <> 0 then
      begin
        Result[K] := TEnum.FromInteger<TI>(V);
        Inc(K);
      end;

      B := B shr 1;
      Inc(V);
    end;
  end;
end;


class function TSet.AsStringArray<T, TI>(const ASet: T): TArray<string>;
var
  Items: TArray<TI>;
  I:     integer;

begin
  Result := TArray<string>.Create();

  CheckIfSet<T>();
  TEnum.CheckIfEnum<TI>();

  Items := AsArray<T, TI>(ASet);
  SetLength(Result, Length(Items));

  for I := 0 to Pred(Length(Items)) do
    Result[I] := TEnum.AsString<TI>(Items[I]);
end;


class function TSet.FromString<T, TI>(const ASetAsStr: string): T;
var
  S: TStringList;

begin
  Result := Default(T);

  CheckIfSet<T>();
  TEnum.CheckIfEnum<TI>();

  S := TStringList.Create;

  try
    S.Duplicates        := dupIgnore;
    S.QuoteChar         := #0;
    S.Delimiter         := ',';
    S.StrictDelimiter   := true;
    S.CaseSensitive     := false;
    S.TrailingLineBreak := false;

    S.DelimitedText := ASetAsStr.Replace('[', '', [rfReplaceAll]).Replace(']', '', [rfReplaceAll]).Replace(' ', '', [rfReplaceAll]);

    Result := FromStringArray<T, TI>(S.ToStringArray);

  finally
    S.Free;
  end;
end;


class function TSet.FromArray<T, TI>(const AArr: TArray<TI>): T;
var
  I: integer;
  J: integer;
  V: integer;
  B: integer;

begin
  Result := Default(T);

  CheckIfSet<T>();
  TEnum.CheckIfEnum<TI>();

  for I := 0 to Pred(Length(AArr)) do
  begin
    V := TEnum.AsInteger<TI>(AArr[I]);
    J := V div 8;
    B := V mod 8;

    PByte(@Result)[J] := PByte(@Result)[J] or (1 shl B);
  end;
end;


class function TSet.FromStringArray<T, TI>(const AStrArr: TArray<string>): T;
var
  Items: TArray<TI>;
  I:     integer;

begin
  Result := Default(T);

  CheckIfSet<T>();
  TEnum.CheckIfEnum<TI>();

  Items := TArray<TI>.Create();
  SetLength(Items, Length(AStrArr));

  for I := 0 to Pred(Length(AStrArr)) do
    Items[I] := TEnum.FromString<TI>(AStrArr[I]);

  Result := FromArray<T, TI>(Items);
end;


end.

