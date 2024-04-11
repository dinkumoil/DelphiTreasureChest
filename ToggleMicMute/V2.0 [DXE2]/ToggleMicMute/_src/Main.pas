{ **************************************************************************** }
//
// ToggleMicMute, a Windows system tray application to mute and unmute your
// microphone by clicking its icon or using a configurable keyboard shortcut.
//
// Written in Delphi XE2.
//
// Target OS: MS Windows (Vista and later)
// Author   : Andreas Heim, 2011
//
//
// This program is free software; you can redistribute it and/or modify it
// under the terms of the GNU General Public License version 3 as published
// by the Free Software Foundation.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License along
// with this program; if not, write to the Free Software Foundation, Inc.,
// 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
//
{ **************************************************************************** }

unit Main;


interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ImgList, IOUtils, Contnrs, StrUtils, IniFiles, ShellAPI,
  ShlObj, ActiveX, CoreAudio, PropSys2, ConfigHotKey, KeyMapper;


resourcestring
  cCopyRight                       = '©2011 by Andreas Heim';

  cMicSection                      = 'Microphone';
  cMicKey                          = 'SelectMicrophone';
  cHotKeySection                   = 'HotKey';
  cHotKeyKeys                      = 'Keys';

  cNoMicAvailable                  = 'No microphone available';

  cErrHotKeyReg                    = 'Failed to register keyboard shortcut.';

  cInfMsgCaption                   = ' - Info';
  cErrMsgCaption                   = ' - Error!';

  cNewMicConnectedAndSelected      = 'A microphone has been connected or activated' + #13+#10 +
                                     'an is selected now:' + #13+#10 +
                                     '%s';

  cNewMicConnectedStillSelectedMic = 'A new microphone has been connected:' + #13+#10 +
                                     '%s' + #13+#10 + #13+#10 +
                                     'Selected microphone still is:' + #13+#10 +
                                     '%s';

  cMicRemovedNewSelectedMic        = 'A microphone has been disconnected or deactivated:' + #13+#10 +
                                     '%s' + #13+#10 + #13+#10 +
                                     'Newly selected microphone:' + #13+#10 +
                                     '%s';

  cMicRemovedStillSelectedMic      = 'A microphone has been disconnected or deactivated' + #13+#10 +
                                     '%s' + #13+#10 + #13+#10 +
                                     'Selected microphone still is:' + #13+#10 +
                                     '%s';

  cMicRemovedNoMicSelected         = 'A microphone has been disconnected or deactivated:' + #13+#10 +
                                     '%s' + #13+#10 + #13+#10 +
                                     'No microphone selected.';


const
  WM_ICONTRAY             = WM_USER + 1;
  WM_EndpointMute         = WM_User + 9;
  WM_MMDeviceStateChanged = WM_User + 1927;
  MSG_MAGIC               = 918273;


type
  TMMNotificationClient = class(TInterfacedObject, IMMNotificationClient)
  public
    function  OnDeviceStateChanged(DeviceID: LPCWSTR; NewState: DEVICE_STATE): HRESULT; stdcall;
    function  OnDeviceAdded(DeviceID: LPCWSTR): HRESULT; stdcall;
    function  OnDeviceRemoved(DeviceID: LPCWSTR): HRESULT; stdcall;
    function  OnDefaultDeviceChanged(Flow: EDataFlow; Role: ERole; NewDefaultDeviceID: LPCWSTR): HRESULT; stdcall;
    function  OnPropertyValueChanged(DeviceID: LPCWSTR; const Key: PROPERTYKEY): HRESULT; stdcall;
  end;


  TAudioEndpointVolumeCallback = class(TInterfacedObject, IAudioEndpointVolumeCallback)
  public
    function OnNotify(pNotify: PAUDIO_VOLUME_NOTIFICATION_DATA): HRESULT; stdcall;
  end;


  TMicDevice = class(TObject)
  public
    Name     : String;
    DeviceID : String;
    DeviceIdx: Cardinal;
  end;


  TfrmMain = class(TForm)
  private
    FIsInit                     : Boolean;

    FAppName                    : String;
    FAppGUID                    : TGUID;

    FIniFileDir                 : String;
    FIniFileName                : String;
    FIniFilePath                : String;

    FAtomStr                    : String;
    FAtom                       : Word;
    FMuteState                  : LongBool;
    FHotKey                     : String;
    FHotKeyRegistered           : Boolean;

    FActMic                     : TMicDevice;
    FMicList                    : TObjectList;

    FTrayIconData               : TNotifyIconData;

    FDeviceEnumerator           : IMMDeviceEnumerator;
    FDeviceCollection           : IMMDeviceCollection;
    FDeviceNotificationClient   : TMMNotificationClient;

    FAudioEndpointVolume        : IAudioEndpointVolume;
    FAudioEndpointVolumeCallback: TAudioEndpointVolumeCallback;

    procedure psmiMicDeviceClick    (Sender: TObject);
    procedure TrayIconClick         (var Msg: TMessage);  message WM_ICONTRAY;
    procedure WMHotkey              (var Msg: TWMHotkey); message WM_HOTKEY;
    procedure WMEndpointMute        (var Msg: TMessage);  message WM_EndpointMute;
    procedure WMMMDeviceStateChanged(var Msg: TMessage);  message WM_MMDeviceStateChanged;

    procedure CreateAppGUID;

    procedure InitApp;
    procedure DeInitApp;

    procedure InitCoreAudio;
    procedure DeInitCoreAudio;

    procedure InitTrayMenu;
    procedure FinalizeTrayMenu;

    procedure GetDefaultMic;
    procedure SetActiveMic;

    procedure SetTrayIcon;
    procedure UnSetTrayIcon;

    procedure SetHotKey(AHotKey: String);
    procedure InstallHotKey;
    procedure DeInstallHotKey;

    procedure DoMute;
    procedure QueryMuteState;
    procedure ToggleMuteState;

    procedure EnumMicrophones;
    function  GetDeviceProperties(DeviceIndex: Cardinal; var DeviceID: String): String;

    procedure ReadIniFile;
    procedure WriteIniFile;
    function  GetAppDataFolder: String;


  public
    property  HotKey : String read FHotKey write SetHotKey;
    property  AppName: String read FAppName;
    property  AppGUID: TGUID  read FAppGUID;


  published
    imlImageList    : TImageList;

    pmTrayPopupMenu : TPopupMenu;

    pmiInfo         : TMenuItem;
    pmiConfigHotKey : TMenuItem;
    pmiAvailableMics: TMenuItem;
    pmiQuit         : TMenuItem;

    psmiDefault     : TMenuItem;

    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);

    procedure pmiInfoClick(Sender: TObject);
    procedure pmiConfigHotKeyClick(Sender: TObject);
    procedure pmiQuitClick(Sender: TObject);

  end;


var
  frmMain     : TfrmMain;
  objKeyMapper: TKeyMapper;



implementation

{$R *.dfm}


//********************************* TfrmMain ***********************************

procedure TfrmMain.FormCreate(Sender: TObject);
var
  ExePath,
  ExeName: String;

begin
  FIsInit := True;

  ExePath             := ReverseString(Application.ExeName);
  ExeName             := ReverseString(AnsiLeftStr(ExePath, PosEx('\', ExePath) - 1));
  FAppName            := AnsiLeftStr(ExeName, PosEx('.', ExeName) - 1);

  Self.Caption        := FAppName;
  Application.Title   := FAppName;
  psmiDefault.Caption := cNoMicAvailable;

  FMicList            := TObjectList.Create(True);
  objKeyMapper        := TKeyMapper.Create;

  FHotKey             := '';
  FHotKeyRegistered   := False;

  FIniFileDir         := GetAppDataFolder;
  FIniFileName        := FAppName + '.ini';
  FIniFilePath        := FIniFileDir + '\' + FIniFileName;

  FAtomStr            := FAppName + 'HotKey';
  FAtom               := GlobalAddAtom(PWideChar(FAtomStr));

  CreateAppGUID;
  InitCoreAudio;
  InitApp;
  SetActiveMic;
  QueryMuteState;
  SetTrayIcon;
  InstallHotKey;

  FIsInit := False;
end;


procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  WriteIniFile;

  DeInstallHotKey;
  GlobalDeleteAtom(FAtom);

  DeInitApp;
  DeInitCoreAudio;

  UnSetTrayIcon;

  objKeyMapper.Free;
  FMicList.Free;
end;


procedure TfrmMain.pmiInfoClick(Sender: TObject);
var
  MsgTxt  : String;
  MsgTitle: String;

begin
  MsgTxt   := FAppName + ', ' + cCopyRight;
  MsgTitle := FAppName + cInfMsgCaption;
  MessageBox(0, PWideChar(MsgTxt), PWideChar(MsgTitle), MB_ICONINFORMATION or MB_OK or MB_SETFOREGROUND);
end;


procedure TfrmMain.pmiConfigHotKeyClick(Sender: TObject);
begin
  frmConfigHotKey.Show;
end;


procedure TfrmMain.pmiQuitClick(Sender: TObject);
begin
  Close;
end;


procedure TfrmMain.psmiMicDeviceClick(Sender: TObject);
var
  I   : Integer;
  AMic: TMicDevice;

begin
  AMic := nil;

  for I := 0 to FMicList.Count - 1 do
    if (FMicList.Items[I] as TMicDevice).Name = (Sender as TMenuItem).Caption then
    begin
      AMic := FMicList.Items[I] as TMicDevice;
      break;
    end;

  if Assigned(AMic) then
  begin
    FActMic := AMic;
    SetActiveMic;
    QueryMuteState;
    SetTrayIcon;
  end;
end;


procedure TfrmMain.TrayIconClick(var Msg: TMessage);
var
  p: TPoint;

begin
  if (Msg.Msg = WM_ICONTRAY) then
  begin
    case Msg.LParam of
      WM_LBUTTONDOWN:
        DoMute;

      WM_RBUTTONDOWN:
      begin
        SetForegroundWindow(Handle);
        GetCursorPos(p);
        pmTrayPopUpMenu.Popup(p.x, p.y);
        PostMessage(Self.Handle, WM_NULL, 0, 0);
      end;
    end;
  end;
end;


procedure TfrmMain.WMHotkey(var Msg: TWMHotkey );
begin
  if (Msg.Msg = WM_HOTKEY) and (Msg.Hotkey = FAtom) then DoMute;
end;


procedure TfrmMain.DoMute;
begin
  QueryMuteState;
  ToggleMuteState;
  SetTrayIcon;
end;


procedure TfrmMain.WMEndpointMute(var Msg: TMessage);
begin
  if (Msg.Msg = WM_EndpointMute) and (Msg.LParam = MSG_MAGIC) then
  begin
    FMuteState := LongBool(Msg.WParam);
    SetTrayIcon;
  end;
end;


procedure TfrmMain.WMMMDeviceStateChanged(var Msg: TMessage);
var
  I             : Integer;
  J             : UInt;
  DeviceID      : String;
  OldMicName    : String;
  ChangedMicName: String;
  ADeviceID     : String;
  MsgTxt        : String;
  MsgTitle      : String;
  DeviceCount   : UInt;
  HR            : HRESULT;

begin
  if (Msg.Msg = WM_MMDeviceStateChanged) and Assigned(PWideChar(Msg.WParam)) then
  begin
    SetString(DeviceID, PWideChar(Msg.WParam), StrLen(PWideChar(Msg.WParam)));
    CoTaskMemFree(PWideChar(Msg.WParam));

    if Assigned(FActMic) then OldMicName := FActMic.Name;

    for I := 0 to FMicList.Count - 1 do
      if ((FMicList.Items[I] as TMicDevice).DeviceID = DeviceID) then
      begin
        ChangedMicName := (FMicList.Items[I] as TMicDevice).Name;
        break;
      end;

    DeInitApp;
    DeInitCoreAudio;
    InitCoreAudio;
    InitApp;

    if Assigned(FActMic) then
    begin
      if (FActMic.Name <> OldMicName) then
        for I := 0 to FMicList.Count - 1 do
          if ((FMicList.Items[I] as TMicDevice).Name = OldMicName) then
          begin
            FActMic := FMicList.Items[I] as TMicDevice;
            FinalizeTrayMenu;
            break;
          end;

      SetActiveMic;
      QueryMuteState;
    end;

    SetTrayIcon;

    if Assigned(FActMic) then
    begin
      // Vor dem Ereignis war kein Mikro ausgewählt.
      // Das neu angeschlossene Mikro wurde ausgewählt.
      if (ChangedMicName = '') and (OldMicName = '') and (FActMic.Name <> '') then
      begin
        MsgTxt   := Format(cNewMicConnectedAndSelected, [FActMic.Name]);
        MsgTitle := FAppName + cInfMsgCaption;

        MessageBox(0, PWideChar(MsgTxt), PWideChar(MsgTitle), MB_ICONINFORMATION or MB_OK or MB_TOPMOST);
      end

      // Vor dem Ereignis war ein Mikro ausgewählt, das jetzt entfernt wurde.
      // Ein anderes Mikro wurde ausgewählt.
      else if (ChangedMicName = OldMicName) and (FActMic.Name <> OldMicName) then
      begin
        MsgTxt   := Format(cMicRemovedNewSelectedMic, [ChangedMicName, FActMic.Name]);
        MsgTitle := FAppName + cInfMsgCaption;

        MessageBox(0, PWideChar(MsgTxt), PWideChar(MsgTitle), MB_ICONINFORMATION or MB_OK or MB_TOPMOST);
      end

      // Vor dem Ereignis war ein Mikro ausgewählt. Ein anderes Mikro wurde entfernt.
      // Das bisher ausgewählte Mikro bleibt ausgewählt.
      else if (ChangedMicName <> '') and (FActMic.Name = OldMicName) then
      begin
        MsgTxt   := Format(cMicRemovedStillSelectedMic, [ChangedMicName, FActMic.Name]);
        MsgTitle := FAppName + cInfMsgCaption;

        MessageBox(0, PWideChar(MsgTxt), PWideChar(MsgTitle), MB_ICONINFORMATION or MB_OK or MB_TOPMOST);
      end

      // Vor dem Ereignis war ein Mikro ausgewählt. Ein neues Mikro wurde angeschlossen.
      // Das bisher ausgewählte Mikro bleibt ausgewählt.
      else if (ChangedMicName = '') and (FActMic.Name = OldMicName) then
      begin
        HR := FDeviceCollection.GetCount(DeviceCount);
        if HR <> S_OK then raise Exception.Create('Error : Unable to get device collection length');

        J := 0;

        while (J < DeviceCount) do
        begin
          ChangedMicName := GetDeviceProperties(J, ADeviceID);
          if (ADeviceID = DeviceID) then break;
          ChangedMicName := '';
          Inc(J);
        end;

        if (ChangedMicName <> '') then
        begin
          MsgTxt   := Format(cNewMicConnectedStillSelectedMic, [ChangedMicName, FActMic.Name]);
          MsgTitle := FAppName + cInfMsgCaption;

          MessageBox(0, PWideChar(MsgTxt), PWideChar(MsgTitle), MB_ICONINFORMATION or MB_OK or MB_TOPMOST);
        end;
      end;
    end
    else
    begin
      // Vor dem Ereignis war ein Mikro ausgewählt, das jetzt entfernt wurde.
      // Es steht kein anderes Mikro zur Verfügung.
      if (ChangedMicName <> '') and (ChangedMicName = OldMicName) then
      begin
        MsgTxt   := Format(cMicRemovedNoMicSelected, [OldMicName]);
        MsgTitle := FAppName + cInfMsgCaption;

        MessageBox(0, PWideChar(MsgTxt), PWideChar(MsgTitle), MB_ICONINFORMATION or MB_OK or MB_TOPMOST);
      end;
    end;
  end;
end;


procedure TfrmMain.CreateAppGUID;
var
  HR: HRESULT;

begin
  try
    HR := CoCreateGUID(FAppGUID);
    if HR <> S_OK then raise Exception.Create('Error : Unable to create GUID');
  except
    PostMessage(Self.Handle, WM_Close, 0, 0);
  end;
end;


procedure TfrmMain.InitCoreAudio;
var
  HR: HResult;

begin
  HR := CoCreateInstance(CLSID_MMDeviceEnumerator, nil, CLSCTX_All,
                         IID_IMMDeviceEnumerator, FDeviceEnumerator);
  if HR <> S_OK then raise Exception.Create('Error : Unable to instantiate device enumerator');

  FDeviceNotificationClient := TMMNotificationClient.Create;

  HR := FDeviceEnumerator.RegisterEndpointNotificationCallback(FDeviceNotificationClient);
  if HR <> S_OK then raise Exception.Create('Error : Unable to register IMMNotificationClient interface');

  HR := FDeviceEnumerator.EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE, FDeviceCollection);
  if HR <> S_OK then raise Exception.Create('Error : Unable to get collection of capture devices');

  FAudioEndpointVolume := nil;
end;


procedure TfrmMain.DeInitCoreAudio;
var
  HR: HRESULT;

begin
  try
    if Assigned(FAudioEndpointVolume) then
    begin
      HR := FAudioEndpointVolume.UnRegisterControlChangeNotify(FAudioEndpointVolumeCallBack);
      if HR <> S_OK then raise Exception.Create('Error : Unable to unregister IAudioEndpointVolumeCallback interface');
    end;

  finally
    try
      if Assigned(FDeviceNotificationClient) then
      begin
        HR := FDeviceEnumerator.UnRegisterEndpointNotificationCallback(FDeviceNotificationClient);
        if HR <> S_OK then raise Exception.Create('Error : Unable to unregister IMMNotificationClient interface');
      end;

    finally
      FAudioEndpointVolume := nil;
      FDeviceCollection    := nil;
      FDeviceEnumerator    := nil;
    end;
  end;
end;


procedure TfrmMain.InitApp;
begin
  FActMic := nil;
  InitTrayMenu;
  GetDefaultMic;
  ReadIniFile;
  FinalizeTrayMenu;
end;


procedure TfrmMain.DeInitApp;
begin
  if Assigned(FActMic) then
  begin
    FActMic := nil;
    pmiAvailableMics.Clear;
    pmiAvailableMics.Add(psmiDefault);
  end;
end;


procedure TfrmMain.InitTrayMenu;
var
  I      : Integer;
  ADevice: TMenuItem;

begin
  EnumMicrophones;

  if pmiAvailableMics.IndexOf(psmiDefault) >= 0 then
  begin
    if FMicList.Count > 0 then
      pmiAvailableMics.Remove(psmiDefault)
  end
  else
  begin
    pmiAvailableMics.Clear;

    if FMicList.Count = 0 then
      pmiAvailableMics.Add(psmiDefault);
  end;

  for I := 0 to FMicList.Count - 1 do
  begin
    ADevice            := TMenuItem.Create(Self);
    ADevice.Caption    := (FMicList.Items[I] as TMicDevice).Name;
    ADevice.GroupIndex := 1;
    ADevice.AutoCheck  := True;
    ADevice.RadioItem  := True;
    ADevice.OnClick    := psmiMicDeviceClick;

    pmiAvailableMics.Add(ADevice);
  end;

  if (FMicList.Count > 0) then
    FActMic := FMicList.Items[0] as TMicDevice;
end;


procedure TfrmMain.FinalizeTrayMenu;
var
  J: Integer;

begin
  if Assigned(FActMic) then
    for J := 0 to pmiAvailableMics.Count - 1 do
      if (pmiAvailableMics.Items[J].Caption = FActMic.Name) then
      begin
        pmiAvailableMics.Items[J].Checked := True;
        break;
      end;
end;


procedure TfrmMain.SetTrayIcon;
var
  TrayIcon : TIcon;

begin
  TrayIcon := TIcon.Create;

  if Assigned(FActMic) then
  begin
    if FMuteState then imlImageList.GetIcon(1, TrayIcon)
    else               imlImageList.GetIcon(0, TrayIcon);
  end
  else
    imlImageList.GetIcon(2, TrayIcon);

  with FTrayIconData do
  begin
    cbSize           := System.SizeOf(FTrayIconData);
    Wnd              := Self.Handle;
    uID              := 0;
    uFlags           := NIF_MESSAGE + NIF_ICON + NIF_TIP;
    uCallbackMessage := WM_ICONTRAY;
    hIcon            := TrayIcon.Handle;

    if Assigned(FActMic) then
      StrPCopy(szTip, FActMic.Name)
    else
      StrPCopy(szTip, cNoMicAvailable);
  end;

  if FIsInit then
    Shell_NotifyIcon(NIM_ADD, @FTrayIconData)
  else
    Shell_NotifyIcon(NIM_MODIFY, @FTrayIconData);

  TrayIcon.Free;
end;


procedure TfrmMain.UnSetTrayIcon;
begin
  Shell_NotifyIcon(NIM_DELETE, @FTrayIconData);
end;


procedure TfrmMain.SetHotKey(AHotKey: String);
begin
  FHotKey := AHotKey;

  if FHotKeyRegistered then
    DeInstallHotKey;

  if (FHotKey <> '') then
    InstallHotKey;
end;


procedure TfrmMain.InstallHotKey;
var
  Modifiers: Cardinal;
  Key      : Cardinal;
  WndHandle: HWND;
  MsgTitle : String;

begin
  if (FHotKey <> '') then
  begin
    objKeyMapper.ParseHotKeyStr(FHotKey, Modifiers, Key);

    if (Modifiers <> 0) or (Key <> 0) then
      FHotKeyRegistered := RegisterHotkey(Handle, FAtom, Modifiers, Key);

    if not FHotKeyRegistered then
    begin
      if FIsInit then WndHandle := 0
      else            WndHandle := frmConfigHotKey.Handle;

      MsgTitle := FAppName + cErrMsgCaption;
      MessageBox(WndHandle, PWideChar(cErrHotKeyReg), PWideChar(MsgTitle), MB_ICONERROR or MB_OK or MB_TOPMOST or MB_APPLMODAL);
      FHotKey := '';
    end;
  end;
end;


procedure TfrmMain.DeInstallHotKey;
begin
  if FHotKeyRegistered then
  begin
    UnRegisterHotkey(Handle, FAtom);
    FHotKeyRegistered := False;
  end;
end;


procedure TfrmMain.GetDefaultMic;
var
  I       : Integer;
  HR      : HRESULT;
  Device  : IMMDevice;
  ID      : PWideChar;
  DeviceID: String;

begin
  if Assigned(FActMic) then
  begin
    HR := FDeviceEnumerator.GetDefaultAudioEndpoint(eCapture, eCommunications, Device);
    if HR <> S_OK then raise Exception.Create('Error : Unable to get standard communications capture device.');

    try
      HR := Device.GetId(ID);
      if HR <> S_OK then raise Exception.Create('Error : Unable to get device id of standard communications capture device.');

      SetString(DeviceID, ID, StrLen(ID));

      for I := 0 to FMicList.Count - 1 do
        if ((FMicList.Items[I] as TMicDevice).DeviceID = DeviceID) then
        begin
          FActMic := FMicList.Items[I] as TMicDevice;
          break;
        end;

    finally
      if ID <> nil then CoTaskMemFree(ID);

    end;
  end;
end;


procedure TfrmMain.SetActiveMic;
var
  HR    : HRESULT;
  Device: IMMDevice;

begin
  if Assigned(FAudioEndpointVolume) then
  begin
    HR := FAudioEndpointVolume.UnRegisterControlChangeNotify(FAudioEndpointVolumeCallBack);
    if HR <> S_OK then raise Exception.Create('Error : Unable to unregister IAudioEndpointVolumeCallback interface');
    FAudioEndpointVolume := nil;
  end;

  if Assigned(FActMic) then
  begin
    HR := FDeviceCollection.Item(FActMic.DeviceIdx, Device);
    if HR <> S_OK then raise Exception.Create('Error : Unable to get IMMDevice interface for device ' + FActMic.Name);

    HR := Device.Activate(IID_IAudioEndpointVolume, CLSCTX_ALL, nil, IUnknown(FAudioEndpointVolume));
    if HR <> S_OK then raise Exception.Create('Error : Unable to get IAudioEndpointVolume interface for device ' + FActMic.Name);

    FAudioEndpointVolumeCallback := TAudioEndpointVolumeCallback.Create;

    HR := FAudioEndpointVolume.RegisterControlChangeNotify(FAudioEndpointVolumeCallback);
    if HR <> S_OK then raise Exception.Create('Error : Unable to register IAudioEndpointVolumeCallback interface for device ' + FActMic.Name);
  end;
end;


procedure TfrmMain.QueryMuteState;
var
  HR: HRESULT;

begin
  if Assigned(FActMic) then
  begin
    HR := FAudioEndpointVolume.GetMute(FMuteState);

    if HR <> S_OK then begin
      FMuteState := False;
      raise Exception.Create('Error : Unable to get mute state of device ' + FActMic.Name);
    end;
  end;
end;


procedure TfrmMain.ToggleMuteState;
var
  B : Integer;
  HR: HRESULT;

begin
  if Assigned(FActMic) then
  begin
    if FMuteState then B:= 0 else B:= 1;
    HR := FAudioEndpointVolume.SetMute(LongBool(B), @FAppGUID);

    if HR <> S_OK then raise Exception.Create('Error : Unable to set mute state of device ' + FActMic.Name)
    else               FMuteState := not FMuteState;
  end;
end;


procedure TfrmMain.EnumMicrophones;
var
  AMic            : TMicDevice;
  DeviceCount     : Cardinal;
  I               : Cardinal;
  HR              : HRESULT;

begin
  FMicList.Clear;

  if Assigned(FDeviceCollection) then
  begin
    HR := FDeviceCollection.GetCount(DeviceCount);
    if HR <> S_OK then raise Exception.Create('Error : Unable to get device collection length');

    I := 0;

    while (I < DeviceCount) do
    begin
      AMic           := TMicDevice.Create;
      AMic.Name      := Trim(GetDeviceProperties(I, AMic.DeviceID));
      AMic.DeviceIdx := I;
      FMicList.Add(AMic);
      Inc(I);
    end;
  end;
end;


function TfrmMain.GetDeviceProperties(DeviceIndex: Cardinal; var DeviceID: String): String;
var
  PropertyStore: IPropertyStore;
  FriendlyName : PROPVARIANT;
  Device       : IMMDevice;
  ID           : PWideChar;
  HR           : HRESULT;

begin
  Result   := '';
  Device   := nil;
  DeviceId := '';

  try
    HR := FDeviceCollection.Item(DeviceIndex, Device);
    if HR <> S_OK then raise Exception.Create('Error : Unable to get IMMDevice interface for device #' + IntToStr(DeviceIndex));

    HR := Device.GetId(ID);
    if HR <> S_OK then raise Exception.Create('Error : Unable to get device id for device #' + IntToStr(DeviceIndex));

    SetString(DeviceID, ID, StrLen(ID));

    HR := Device.OpenPropertyStore(STGM_READ, PropertyStore);
    if HR <> S_OK then raise Exception.Create('Error : Unable to open device property store for device #' + IntToStr(DeviceIndex));

    PropVariantInit(FriendlyName);

    HR := PropertyStore.GetValue(PKEY_Device_FriendlyName, FriendlyName);
    if HR <> S_OK then raise Exception.Create('Error : Unable to get friendly name for device #' + IntToStr(DeviceIndex));

    if (FriendlyName.vt = VT_LPWSTR) then
      SetString(Result, FriendlyName.pwszVal, StrLen(FriendlyName.pwszVal));

  finally
    if ID <> nil then CoTaskMemFree(ID);
    PropVariantClear(@FriendlyName);

  end;
end;


procedure TfrmMain.ReadIniFile;
var
  I      : Integer;
  MicName: String;
  IniFile: TIniFile;

begin
  IniFile := TIniFile.Create(FIniFilePath);

  MicName := IniFile.ReadString(cMicSection, cMicKey, '');
  if FIsInit then FHotKey := IniFile.ReadString(cHotKeySection, cHotKeyKeys, '');

  if (MicName <> '') then
    for I := 0 to FMicList.Count - 1 do
      if ((FMicList.Items[I] as TMicDevice).Name = MicName) then
      begin
        FActMic := FMicList.Items[I] as TMicDevice;
        break;
      end;

  IniFile.Free;
end;


procedure TfrmMain.WriteIniFile;
var
  IniFile: TIniFile;

begin
  if Assigned(FActMic) then
  begin
    if not TDirectory.Exists(FIniFileDir) then
      TDirectory.CreateDirectory(FIniFileDir);

    IniFile := TIniFile.Create(FIniFilePath);
    IniFile.WriteString(cMicSection, cMicKey, FActMic.Name);
    IniFile.WriteString(cHotKeySection, cHotKeyKeys, FHotKey);
    IniFile.Free;
  end;
end;


function TfrmMain.GetAppDataFolder: String;
begin
  SetLength(Result, MAX_PATH);

  If SHGetSpecialFolderPath(Handle, PWideChar(Result), CSIDL_APPDATA, False) then
  begin
    SetLength(Result, StrLen(PWideChar(Result)));
    Result := Result + '\' + FAppName;
  end
  else
    Result := '.';
end;
//------------------------------------------------------------------------------


//**************************** TMMNotificationClient ***************************

function TMMNotificationClient.OnDeviceStateChanged(DeviceID: LPCWSTR; NewState: DEVICE_STATE): HRESULT; stdcall;
var
  _DeviceID: PWideChar;

begin
  if not Assigned(DeviceID) then
  begin
    Result := E_INVALIDARG;
    exit;
  end;

  _DeviceID := CoTaskMemAlloc((StrLen(DeviceID) + 1) * SizeOf(WideChar));

  if not Assigned(_DeviceID) then
  begin
    Result := E_OUTOFMEMORY;
    exit;
  end;

  StrCopy(_DeviceID, DeviceID);

  PostMessage(frmMain.Handle, WM_MMDeviceStateChanged, Integer(_DeviceID), Integer(NewState));

  Result := S_OK;
end;


function TMMNotificationClient.OnDeviceAdded(DeviceID: LPCWSTR): HRESULT; stdcall;
begin
  Result := S_OK;
end;


function TMMNotificationClient.OnDeviceRemoved(DeviceID: LPCWSTR): HRESULT; stdcall;
begin
  Result := S_OK;
end;


function TMMNotificationClient.OnDefaultDeviceChanged(Flow: EDataFlow; Role: ERole; NewDefaultDeviceId: LPCWSTR): HRESULT; stdcall;
begin
  Result := S_OK;
end;


function TMMNotificationClient.OnPropertyValueChanged(DeviceID: LPCWSTR; const Key: PROPERTYKEY): HRESULT; stdcall;
begin
  Result := S_OK;
end;
//------------------------------------------------------------------------------


//*********************** TAudioEndpointVolumeCallback *************************

function TAudioEndpointVolumeCallback.OnNotify(pNotify: PAUDIO_VOLUME_NOTIFICATION_DATA):HRESULT; stdcall;
begin
  if not Assigned(pNotify) then
  begin
    Result := E_INVALIDARG;
    exit;
  end;

  if (pNotify^.guidEventContext <> frmMain.AppGUID) then
    PostMessage(frmMain.Handle, WM_EndpointMute, Integer(pNotify^.bMuted), MSG_MAGIC);

  Result := S_OK;
end;
//------------------------------------------------------------------------------


end.
