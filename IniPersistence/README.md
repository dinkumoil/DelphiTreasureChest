# IniPersistence

The unit _System.IniFiles.Persistence.pas_ provides generic INI file handling based on _Delphi_ attributes. An example settings class using it is included in file _Settings.pas_.

The code is based on ideas of Robert Love (as part of his [blog post here](http://robstechcorner.blogspot.com/2009/09/so-what-is-rtti-rtti-is-acronym-for-run.html)) and was extensively improved by me.


## What is it for?

Generic INI file handling means that you can decorate the properties or field variables of a _Delphi_ class with attributes in order to determine the INI file's section, the key name and its default value where you want to store the property's/field's value. When you need a new property/field to be stored in the INI file you simply decorate it with an appropriate attribute (depending on its data type) and you are done. Thus, you don't have to write any code, adding new properties/fields to your INI file is reduced to declarative programming.


## How does it work (the easy way)?

The class _TSettings_ in unit _Settings.pas_ is not only an example of how to use generic INI file handling. The unit is ready-to-use for handling INI files in your own projects. Simply include it in your project and adapt its properties (and their attributes) to your needs and you are done. Thanks to its auto-load feature it will search for an INI file in your user profile under `%AppData%\<your-exe-name>\<your-exe-name>.ini`. If it isn't there (for example when your program runs for the very first time) it creates a _TSettings_ object with the default values you provided when declaring the attributes. Saving the INI file is also done automatically. When the program terminates, the destructor of _TSettings_ is executed due to the call of its `Free` method from the `finalization` section of unit _Settings.pas_. The destructor in turn calls the `Save` method. Since _TSettings_ is a singleton class, to access properties of the settings object write code like `if TSettings.Instance.Active then ...`.


## Doing things by yourself

If you want more control over where the INI file is stored and which name it has, you can disable auto-loading by setting in unit _Settings.pas_ the variable `SettingsAutoLoad` to false. Then you will have the following options to load your INI file:

1. Pass the path to your INI file as an argument to _TSettings_' constructor. The INI file will still be loaded and saved automatically. The life-time management of the settings object is also done automatically.

```Pascal
procedure TMainForm.FormCreate(Sender: TObject);
begin
  TSettings.Create('.\MyIniFile.ini');
end;
```

2. Set an event handler for retrieving the INI files's path and call _TSettings_' `Load` method manually. The event handler is used every time the settings class needs to know the path of the INI file. This way you are even able to change the INI file's path at runtime. But you have to be careful what the code of the event handler does. Since it is executed immediately before program termination it should not access resources that are already freed at that time. To avoid this you can also free the settings object manually. To do so, comment out the call to `TSettings.Instance.Free` in the `finalization` section of unit _TSettings.pas_ and add that call to the code of your main form's destructor or `FormDestroy` event handler.

```Pascal
procedure TMainForm.FormCreate(Sender: TObject);
begin
  TSettings.OnGetFilePath := GetSettingsFilePath;
  TSettings.Load;
end;

function TMainForm.GetSettingsFilePath: string;
begin
  Result := '.\MyIniFile.ini';
end;
```

3. Set an anonymous procedure (or closure) for retrieving the INI files's path and call _TSettings_' `Load` method manually. The anonymous procedure is used every time the settings class needs to know the path of the INI file, so you are able to change the INI file's path at runtime by using this variant as well. Like with the event handler variant mentioned above you must be careful about what resources the anonymous procedure's code accesses. See my advice there.

```Pascal
procedure TMainForm.FormCreate(Sender: TObject);
begin
  TSettings.FnGetFilePath :=
    function: string
    begin
      Result := '.\MyIniFile.ini';
    end;

  TSettings.Load;
end;
```


## Advanced Topics

### About its internal working

The code heavily relies on attributes and Delphi's RTTI (runtime type information) system. For each Delphi data type that is to be supported, there must be a corresponding attribute class derived from the `IniValueAttribute` class.

Attribute classes are not only used to provide information about the INI section and the key within that section where the values of the properties and field variables they decorate are to be stored. They are also used to specify the default values for those properties and fields, so type-specific attributes are required.

To add special options to certain data types, you can also define an attribute class derived from `InterfacedIniValueAttribute`. This is used with class `IniNumericValueAttribute` and its descendants to simplify support for decimal and hex notation of integer values in the INI file.

For properties and field variables the following data types are supported out-of-the-box:

- string
- boolean
- 32-bit signed integer (in decimal and hex notation)
- 32-bit unsigned integer (in decimal and hex notation)
- 64-bit signed integer (in decimal and hex notation)
- 64-bit unsigned integer (in decimal and hex notation)
- Floating point values (extended, double and single precision)
- TDateTime (with UTC and local time support)
- TColor (always in hex notation)
- arbitrary enumeration types (for which RTTI is available)
- sets of arbitrary types (which are supported by Delphi's `StringToSet`)
- TGUID

All of these data types are supported both as scalar values and as dynamic and static arrays. If the property/field is of an array type whose elements are of a supported scalar type, the elements are mapped to indexed values in the INI file, see chapter [Indexed INI values mapped to arrays](#indexed-ini-values-mapped-to-arrays).

If you want to add support for additional data types, you have to write your own corresponding attribute classes and to extend the class methods `TIniPersistence.GetValue` and `TIniPersistence.SetValue`. List types may require you to extend the class methods `TIniPersistence.Serialize<T>` and `TIniPersistence.Deserialize<T>` as well.


### Indexed INI values mapped to arrays

In order to use indexed values in an INI file (e.g. `File1=xxx File2=yyy`) it is possible to declare properties or field variables as an array. The most common use case is to declare them as a modern dynamic array, for example `TArray<string>` (the array's element type can be of any supported scalar data type, have a look at the `IniXxxValueAttribute` classes). But it is also possible to declare them as a traditional dynamic array (e.g. `array of string`) or as a static array (e.g. `array[0..9] of string`).

However, these last two cases are special because of Delphi's type identity rules (see chapter [Type Compatibility and Identity](https://docwiki.embarcadero.com/RADStudio/Florence/en/Type_Compatibility_and_Identity_(Delphi)) of the _Delphi_ documentation). Thus you have to declare intermediate types (e.g. `TStrDynArr = array of string;`). The declaration `array[0..9] of string` requires even more effort - additionally you have to declare a subrange type for the array indices. So, the final declaration would be `TArrIndex = 0..9; TStrStatArr = array[TArrIndex] of string;`. If you don't do it that way, your code won't compile or the _Delphi_ compiler will create no or incomplete runtime type information for the type and you will get runtime errors.

Indexed INI values must always start with index 1 (e.g. `File1=xxx`). The index must increase contiguously, indexed values beyond a gap are ignored. The first indexed value is mapped to the first array element (i.e. for `TArray<string>` the value of `File1` would be in index `0`).

For properties or field variables of a **dynamic array type** the following rules are applied:

- When the INI file is read and contains no indexed values for a certain property or field, the related array remains empty.
- If the array is empty when the containing object is written to an INI file, no INI values are written at all.
- If you want the dynamic array of a property or field variable to contain a certain number of indexed values (e.g. `File1` to `File10`), it is possible to additionally decorate it with the `IniArrayLength` attribute and provide this number as a parameter (e.g. `IniArrayLength(4)` to have 4 array elements at least). When the INI file is read, the array elements of missing indexed values will be set to the default value passed as parameter to the `IniXxxValue` attribute.

For properties or field variables of a **static array type** the following rules are applied:

- The array always has the number of elements that has been specified in its type declaration.
- When the INI file is read and contains less indexed values for a certain property or field variable than the related array is able to hold, the remaining elements of the array will be initialized with the default value passed as parameter to the `IniXxxValue` attribute. Indexed values in the INI file that exceed the maximum number of elements of the array will be ignored.
- When the object is written to an INI file, as many indexed values are written as there are elements in the array, according to its type declaration.


## History

v2.0 - April 2026

- Added support for indexed INI values mapped to arrays

v1.0 - May 2024

- Initial version
