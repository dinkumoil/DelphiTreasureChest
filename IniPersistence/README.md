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

2. Set an event handler for retrieving the INI files's path and call _TSettings_' `Load` method manually. The event handler is used every time the settings class needs to know the path of the INI file. This way you are even able to change the INI file's path at runtime. But you have to be carful what the code of the event handler does. Since it is executed immediately before program termination it should not access resources that are already freed at that time. To avoid this you can also free the settings object manually. To do so, comment out the call to `TSettings.Instance.Free` in the `finalization` section of unit _TSettings.pas_ and add that call to the code of your main form's destructor or `FormDestroy` event handler.

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
