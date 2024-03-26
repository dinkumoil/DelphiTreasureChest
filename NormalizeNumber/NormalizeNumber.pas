uses
  System.Math, System.SysUtils;


function NormalizeNumber(ANumber: extended; NumberSystemBase: cardinal; out Prefix: string; BinaryPrefix: boolean = false): extended;
type
  TPrefix = (pfxYocto, pfxZepto, pfxAtto, pfxFemto, pfxPico, pfxNano, pfxMicro, pfxMilli,
             pfxNone,
             pfxKilo,  pfxMega,  pfxGiga, pfxTera,  pfxPeta, pfxExa,  pfxZetta, pfxYotta);

const
  Prefixes: array[TPrefix] of string = ('y', 'z', 'a', 'f', 'p', 'n', 'µ', 'm',
                                        '',
                                        'k', 'M', 'G', 'T', 'P', 'E', 'Z', 'Y');

var
  IntPartSign:    integer;
  PrefixFactor:   double;
  Factor:         double;
  StepDir:        integer;
  PrefixIdxLimit: TPrefix;

  TmpNumber:      double;
  TmpIntPartSgin: integer;
  PrefixIdx:      TPrefix;

label
  Quit;

begin
  Result    := ANumber;
  PrefixIdx := pfxNone;

  if ANumber = 0                       then goto Quit;
  if not (NumberSystemBase in [2, 10]) then goto Quit;

  PrefixFactor := IfThen(NumberSystemBase = 2, 1024, 1000);

  if Int(ANumber) = 0 then
  begin
    IntPartSign    := 1;
    StepDir        := -1;
    Factor         := PrefixFactor;
    PrefixIdxLimit := Low(TPrefix);
  end
  else
  begin
    IntPartSign    := 0;
    StepDir        := 1;
    Factor         := 1 / PrefixFactor;
    PrefixIdxLimit := High(TPrefix);
  end;

  repeat
    TmpNumber      := Result * Factor;
    TmpIntPartSgin := Sign(Abs(Int(TmpNumber)));

    if (TmpIntPartSgin = IntPartSign) and (IntPartSign = 0) then break;
    if PrefixIdx       = PrefixIdxLimit                     then break;

    Result := TmpNumber;
    Inc(PrefixIdx, StepDir);

    if (TmpIntPartSgin = IntPartSign) and (IntPartSign = 1) then break;
  until false;

Quit:
  Prefix := Prefixes[PrefixIdx];

  if (NumberSystemBase = 2) and BinaryPrefix and (PrefixIdx <> pfxNone) then
    Prefix := Prefix + 'i';
end;
