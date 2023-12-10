(*******************************************************************************)
(*                                                                             *)
(* Smart pointer implementation                                                *)
(* See: https://adugmembers.wordpress.com/2011/12/05/smart-pointers/           *)
(*                                                                             *)
(* Example usage                                                               *)
(* -------------                                                               *)
(*                                                                             *)
(* type                                                                        *)
(*   TPerson = class                                                           *)
(*   public                                                                    *)
(*     constructor Create(const AName: string; const AAge: Integer);           *)
(*     procedure Birthday; // Increment Age                                    *)
(*     property Name: string ...                                               *)
(*     property Age: integer ...                                               *)
(*   end;                                                                      *)
(*                                                                             *)
(* var                                                                         *)
(*   Person1: ISmartPointer<TPerson>;                                          *)
(*   Person2: ISmartPointer<TPerson>;                                          *)
(*   Person3: ISmartPointer<TPerson>;                                          *)
(*   PersonObj: TPerson;                                                       *)
(*                                                                             *)
(*   // Smart pointer param                                                    *)
(*   procedure ShowName(APerson: ISmartPointer<TPerson>);                      *)
(*                                                                             *)
(*   // TPerson param                                                          *)
(*   procedure ShowAge(APerson: TPerson);                                      *)
(*                                                                             *)
(* begin                                                                       *)
(*   // Typical usage when creating a new object to manage                     *)
(*   Person1 := TSmartPointer<TPerson>.Create(TPerson.Create('Fred', 100));    *)
(*   Person1.Birthday; // Direct member access!                                *)
(*   ShowName(Person1); // Pass as smart pointer                               *)
(*   ShowAge(Person1); // Pass as the managed object!                          *)
(*   //Person1 := nil; // Release early                                        *)
(*                                                                             *)
(*   // Same as above but hand over to smart pointer later                     *)
(*   PersonObj := TPerson.Create('Wilma', 90);                                 *)
(*   Person2 := TSmartPointer<TPerson>.Create(PersonObj);                      *)
(*   ShowName(Person2);                                                        *)
(*   // Note: PersonObj is freed by the smart pointer                          *)
(*                                                                             *)
(*   // Smart pointer constructs the TPerson instance                          *)
(*   Person3 := TSmartPointer<TPerson>.Create(); // or Create(nil)             *)
(*                                                                             *)
(*   // The smart pointer references are released in reverse declaration order *)
(*   // (Person3, Person2, Person1)                                            *)
(* end;                                                                        *)
(*                                                                             *)
(*                                                                             *)
(*                                                                             *)
(* Adding a record and you can get rid of the extra Create. Only disadvantage  *)
(* you cannot combine the power of records operator overloads and the implicit *)
(* Invoke call of the interface. But you can combine them:                     *)
(*                                                                             *)
(* var                                                                         *)
(*   s: SmartPointer<TPerson>;                                                 *)
(*   i: ISmartPointer<TPerson>;                                                *)
(* begin                                                                       *)
(*   s := TPerson.Create;                                                      *)
(*   i := s;                                                                   *)
(*   i.Age := 32;                                                              *)
(* end;                                                                        *)
(*                                                                             *)
(*******************************************************************************)


unit SmartPointer;


interface

uses
  System.SysUtils;


type
  ISmartPointer<T> = reference to function: T;


  TSmartPointer<T: class, constructor> = class(TInterfacedObject, ISmartPointer<T>)
  private
    FValue: T;

  public
    constructor Create; overload;
    constructor Create(AValue: T); overload;
    destructor  Destroy; override;

    function    Invoke: T;

  end;


  SmartPointer<T: class, constructor> = record
  strict private
    FPointer: ISmartPointer<T>;

  public
    class operator Implicit(Value: T): SmartPointer<T>;
    class operator Implicit(const Value: SmartPointer<T>): ISmartPointer<T>;

  end;



implementation


{ TSmartPointer<T> }

constructor TSmartPointer<T>.Create;
begin
  inherited;

  FValue := T.Create;
end;


constructor TSmartPointer<T>.Create(AValue: T);
begin
  inherited Create;

  if AValue = nil then
    FValue := T.Create
  else
    FValue := AValue;
end;


destructor TSmartPointer<T>.Destroy;
begin
  FValue.Free;

  inherited;
end;


function TSmartPointer<T>.Invoke: T;
begin
  Result := FValue;
end;



{ SmartPointer<T> }

class operator SmartPointer<T>.Implicit(Value: T): SmartPointer<T>;
begin
  Result.FPointer := TSmartPointer<T>.Create(Value);
end;


class operator SmartPointer<T>.Implicit(const Value: SmartPointer<T>): ISmartPointer<T>;
begin
  Result := Value.FPointer;
end;


end.
