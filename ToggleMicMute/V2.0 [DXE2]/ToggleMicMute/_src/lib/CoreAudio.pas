{ **************************************************************************** }
//
// Partial port of the following Windows header files:
//   - Mmdeviceapi.h (Multi Media Device API)
//   - Audioclient.h and Audiopolicy.h (WASAPI, Windows Audio Session API)
//   - Endpointvolume.h (Endpoint Volume API)
//
// URL: https://learn.microsoft.com/en-us/windows/win32/coreaudio
//
// This code is part of ToggleMicMute, a Windows system tray application to mute
// and unmute your microphone by clicking its icon or using a keyboard shortcut.
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

unit CoreAudio;


{$MINENUMSIZE 4}
{$WEAKPACKAGEUNIT}


interface

uses
  Windows, ActiveX, ComObj, PropSys2, MMSystem;


const
  // COM-GUIDs of the CoreAudio-Interfaces
  CLSID_MMDeviceEnumerator        : TGUID = '{BCDE0395-E52F-467C-8E3D-C4579291692E}';

  IID_IMMDevice                   : TGUID = '{D666063F-1587-4E43-81F1-B948E807363F}';
  IID_IMMEndpoint                 : TGUID = '{1BE09788-6894-4089-8586-9A2A6C265AC5}';
  IID_IMMDeviceEnumerator         : TGUID = '{A95664D2-9614-4F35-A746-DE8DB63617E6}';
  IID_IMMDeviceCollection         : TGUID = '{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}';
  IID_IMMNotificationClient       : TGUID = '{7991EEC9-7E89-4D85-8390-6C703CEC60C0}';

  IID_IAudioSessionManager        : TGUID = '{BFA971F1-4D5E-40BB-935E-967039BFBEE4}';
  IID_IAudioSessionControl        : TGUID = '{F4B1A599-7266-4319-A8CA-E70ACB11E8CD}';
  IID_IAudioSessionEvents         : TGUID = '{24918ACC-64B3-37C1-8CA9-74A66E9957A8}';
  IID_IAudioClient                : TGUID = '{1CB9AD4C-DBFA-4c32-B178-C2F568A703B2}';
  IID_IAudioCaptureClient         : TGUID = '{C8ADBD64-E71E-48a0-A4DE-185C395CD317}';
  IID_IAudioRenderClient          : TGUID = '{F294ACFC-3146-4483-A7BF-ADDCA7C260E2}';
  IID_ISimpleAudioVolume          : TGUID = '{87CE5498-68D6-44E5-9215-6DA47EF883D8}';
  IID_IAudioStreamVolume          : TGUID = '{93014887-242D-4068-8A15-CF5E93B90FE3}';
  IID_IChannelAudioVolume         : TGUID = '{1C158861-B533-4B30-B1CF-E853E51C59B8}';
  IID_IAudioClock                 : TGUID = '{CD63314F-3FBA-4A1B-812C-EF96358728E7}';
  IID_IAudioClock2                : TGUID = '{6F49FF73-6727-49AC-A008-D98CF5E70048}';
  IID_IAudioClockAdjustment       : TGUID = '{F6E4C0A0-46D9-4FB9-BE21-57A3EF2B626C}';

  // ******* undocumented *******
  IID_IAudioSessionQuerier        : TGUID = '{94BE9D30-53AC-4802-829C-F13E5AD34776}';
  IID_IAudioSessionQuery          : TGUID = '{94BE9D30-53AC-4802-829C-F13E5AD34775}';
  IID_IRemoteAudioSession         : TGUID = '{33969B1D-D06F-4281-B837-7EAAFD21A9C0}';
  // ****************************

  IID_IAudioEndpointVolume        : TGUID = '{5CDF2C82-841E-4546-9722-0CF74078229A}';
  IID_IAudioEndpointVolumeEx      : TGUID = '{66E11784-F695-4F28-A505-A7080081A78F}';
  IID_IAudioEndpointVolumeCallback: TGUID = '{657804FA-D6AD-4496-8A60-352752AF4F89}';
  IID_IAudioMeterInformation      : TGUID = '{C02216F6-8C67-4B5B-9D00-D008E73E0064}';



/////// Constans, Datatypes and COM-Interfaces of CoreAudio ///////

(* ------------------- MultiMediaDevice API -------------------- *)
const
  eRender                 = $00000000;
  eCapture                = $00000001;
  eAll                    = $00000002;
  EDataFlow_enum_count    = $00000003;

  eConsole                = $00000000;
  eMultimedia             = $00000001;
  eCommunications         = $00000002;
  ERole_enum_count        = $00000003;

  DEVICE_STATE_ACTIVE     = $00000001;
  DEVICE_STATE_DISABLED   = $00000002;
  DEVICE_STATE_NOTPRESENT = $00000004;
  DEVICE_STATE_UNPLUGGED  = $00000008;
  DEVICE_STATEMASK_ALL    = $0000000F;

  STGM_READ               = $00000000;
  STGM_WRITE              = $00000001;
  STGM_READWRITE          = $00000002;


type
  EDataFlow    = TOleEnum;
  ERole        = TOleEnum;
  DEVICE_STATE = DWORD;


type
  // Forward Deklarationen
  IMMDevice             = interface;
  IMMEndpoint           = interface;
  IMMDeviceEnumerator   = interface;
  IMMDeviceCollection   = interface;
  IMMNotificationClient = interface;


  IMMDevice = interface(IUnknown)
    ['{D666063F-1587-4E43-81F1-B948E807363F}']
    function Activate(const iid: TGUID; dwClsCtx: DWORD; pActivationParams: PPropVariant; out ppInterface: IUnknown): HRESULT; stdcall;
    function OpenPropertyStore(stgmAccess: DWORD; out ppProperties: IPropertyStore): HRESULT; stdcall;
    function GetId(out ppstrId: PWChar): HRESULT; stdcall;
    function GetState(out pdwState: DEVICE_STATE): HRESULT; stdcall;
  end;


  IMMEndpoint = interface(IUnknown)
    ['{1BE09788-6894-4089-8586-9A2A6C265AC5}']
    function GetDataFlow(out pDataFlow: EDataFlow): HRESULT; stdcall;
  end;


  IMMDeviceEnumerator = interface(IUnknown)
    ['{A95664D2-9614-4F35-A746-DE8DB63617E6}']
    function EnumAudioEndpoints(dataFlow: EDataFlow; dwStateMask: DEVICE_STATE; out ppDevices: IMMDeviceCollection): HRESULT; stdcall;
    function GetDefaultAudioEndpoint(dataFlow: EDataFlow; role: ERole; out ppEndpoint: IMMDevice): HRESULT; stdcall;
    function GetDevice(pwstrId: PWChar; out ppDevice: IMMDevice): HRESULT; stdcall;
    function RegisterEndpointNotificationCallback(pClient: IMMNotificationClient): HRESULT; stdcall;
    function UnregisterEndpointNotificationCallback(pClient: IMMNotificationClient): HRESULT; stdcall;
  end;


  IMMDeviceCollection = interface(IUnknown)
    ['{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}']
    function GetCount(out pcDevices: DWORD):HRESULT; stdcall;
    function Item(nDevice: DWORD; out ppDevice: IMMDevice): HRESULT; stdcall;
  end;


  IMMNotificationClient = interface(IUnknown)
    ['{7991EEC9-7E89-4D85-8390-6C703CEC60C0}']
    function OnDeviceStateChanged(DeviceId: LPCWSTR; NewState: DEVICE_STATE): HRESULT; stdcall;
    function OnDeviceAdded(DeviceId: LPCWSTR): HRESULT; stdcall;
    function OnDeviceRemoved(DeviceId: LPCWSTR): HRESULT; stdcall;
    function OnDefaultDeviceChanged(Flow: EDataFlow; Role: ERole; NewDefaultDeviceId: LPCWSTR): HRESULT; stdcall;
    function OnPropertyValueChanged(DeviceId: LPCWSTR; const Key: PROPERTYKEY): HRESULT; stdcall;
  end;
(* ----------------- MultiMediaDevice API End ------------------ *)



(* -------------------- WindowsAudioSession API ---------------- *)
const
  CyclePerSec                                  = 25;
  REFTIMES_PER_SEC                             = 10000000;
  REFTIMES_PER_MILLISEC                        = 10000;

  AUDCLNT_SHAREMODE_SHARED                     = 0;
  AUDCLNT_SHAREMODE_EXCLUSIVE                  = 1;

  AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY       = 1;
  AUDCLNT_BUFFERFLAGS_SILENT                   = 2;
  AUDCLNT_BUFFERFLAGS_TIMESTAMP_ERROR          = 4;

  AUDCLNT_STREAMFLAGS_CROSSPROCESS             = $00010000;
  AUDCLNT_STREAMFLAGS_LOOPBACK                 = $00020000;
  AUDCLNT_STREAMFLAGS_EVENTCALLBACK            = $00040000;
  AUDCLNT_STREAMFLAGS_NOPERSIST                = $00080000;
  AUDCLNT_STREAMFLAGS_RATEADJUST               = $00100000;

  AUDCLNT_SESSIONFLAGS_EXPIREWHENUNOWNED       = $10000000;
  AUDCLNT_SESSIONFLAGS_DISPLAY_HIDE            = $20000000;
  AUDCLNT_SESSIONFLAGS_DISPLAY_HIDEWHENEXPIRED = $40000000;

  // Constants for AudioSessionState enumeration
  AudioSessionStateInactive                    = 0; // The last running stream in the session stops.
  AudioSessionStateActive                      = 1; // The session has one or more streams that are running.
  AudioSessionStateExpired                     = 2; // The client destroys the last stream in the session by
                                                    // releasing all references to the stream object.

  // Constants for AudioSessionDisconnectReason enumeration
  DisconnectReasonDeviceRemoval                = 0; // The user removed the audio endpoint device.
  DisconnectReasonServerShutdown               = 1; // The Windows audio service has stopped.
  DisconnectReasonFormatChanged                = 2; // The stream format changed for the device that the audio session is connected to.
  DisconnectReasonSessionLogoff                = 3; // The user logged off the Windows Terminal Services (WTS) session that the audio session was running in.
  DisconnectReasonSessionDisconnected          = 4; // The WTS session that the audio session was running in was disconnected.
  DisconnectReasonExclusiveModeOverride        = 5; // The (shared-mode) audio session was disconnected to make the audio endpoint device available for an exclusive-mode connection.


type
  REFERENCE_TIME                = Int64;
  AUDCLNT_SHAREMODE             = TOleEnum;
  AudioSessionState             = TOleEnum;
  AudioSessionDisconnectReason  = TOleEnum;
  

type
  // Forward Deklarationen
  IAudioSessionManager  = interface;
  IAudioSessionControl  = interface;
  ISimpleAudioVolume    = interface;
  IAudioStreamVolume    = interface;
  IChannelAudioVolume   = interface;
  IAudioClient          = interface;
  IAudioRenderClient    = interface;
  IAudioCaptureClient   = interface;
  IAudioSessionEvents   = interface;
  IAudioClock           = interface;
  IAudioClock2          = interface;
  IAudioClockAdjustment = interface;


  IAudioSessionManager = interface(IUnknown)
    ['{BFA971F1-4D5E-40BB-935E-967039BFBEE4}']
    function GetAudioSessionControl(AudioSessionGuid: pGUID; CrossProcessSession: LongBool;
      out SessionControl: IAudioSessionControl): HResult; stdcall;
    function GetSimpleAudioVolume(AudioSessionGuid: pGuid; StreamFlag: DWORD;
      out AudioVolume: ISimpleAudioVolume): HResult; stdcall;
  end;


  IAudioSessionControl = interface(IUnknown)
    ['{F4B1A599-7266-4319-A8CA-E70ACB11E8CD}']
    function GetState(out pRetVal: AudioSessionState): HResult; stdcall;
    function GetDisplayName(out pRetVal: LPWSTR): HResult; stdcall; // pRetVal must be freed by CoTaskMemFree
    function SetDisplayName(Value: LPCWSTR; EventContext: pGuid): HResult; stdcall;
    function GetIconPath(out pRetVal: LPWSTR): HResult; stdcall; // pRetVal must be freed by CoTaskMemFree
    function SetIconPath(Value: LPCWSTR; EventContext: pGuid): HResult; stdcall;
    function GetGroupingParam(pRetVal: pGuid): HResult; stdcall;
    function SetGroupingParam(OverrideValue, EventContext: pGuid): HResult; stdcall;
    function RegisterAudioSessionNotification(NewNotifications: IAudioSessionEvents): HResult; stdcall;
    function UnregisterAudioSessionNotification(NewNotifications: IAudioSessionEvents): HResult; stdcall;
  end;


  ISimpleAudioVolume = interface(IUnknown)
    ['{87CE5498-68D6-44E5-9215-6DA47EF883D8}']
    function SetMasterVolume(fLevel: Single; EventContext: pGuid): HResult; stdcall;
    function GetMasterVolume(out fLevel: Single): HResult; stdcall;
    function SetMute(bMute: LongBool; EventContext: pGuid): HResult; stdcall;
    function GetMute(out bMute: LongBool): HResult; stdcall;
  end;


  IAudioStreamVolume = interface(IUnknown)
    ['{93014887-242D-4068-8A15-CF5E93B90FE3}']
    function GetChannelCount(out pdwCount: UINT32): HRESULT; stdcall;
    function SetChannelVolume(dwIndex: UINT32; const fLevel: Single): HRESULT; stdcall;
    function GetChannelVolume(dwIndex: UINT32; out pfLevel: Single): HRESULT; stdcall;
    function SetAllVolumes(dwCount: UINT32; const pfVolumes: array of Single): HRESULT; stdcall;
    function GetAllVolumes(dwCount: UINT32; out pfVolumes: array of Single): HRESULT; stdcall;
  end;


  IChannelAudioVolume = interface(IUnknown)
    ['{1C158861-B533-4B30-B1CF-E853E51C59B8}']
    function GetChannelCount(out pdwCount: UINT32): HRESULT; stdcall;
    function SetChannelVolume(dwIndex: UINT32; const fLevel: Single; EventContext: pGuid): HRESULT; stdcall;
    function GetChannelVolume(dwIndex: UINT32; out pfLevel: Single): HRESULT; stdcall;
    function SetAllVolumes(dwCount: UINT32; const pfVolumes: array of Single; EventContext: pGuid): HRESULT; stdcall;
    function GetAllVolumes(dwCount: UINT32; out pfVolumes: array of Single): HRESULT; stdcall;
  end;


  IAudioClient = interface(IUnknown)
    ['{1CB9AD4C-DBFA-4c32-B178-C2F568A703B2}']
    function Initialize(ShareMode: AUDCLNT_SHAREMODE; StreamFlags: DWORD; hmsBufferDuration: REFERENCE_TIME;
      hmsPeriodicity: REFERENCE_TIME; pFormat: PWAVEFORMATEX; AudioSessionGuid: pGuid): HRESULT; stdcall;
    function GetBufferSize(out NumBufferFrames: UINT32): HRESULT; stdcall;
    function GetStreamLatency(out hmsLatency: REFERENCE_TIME): HRESULT; stdcall;
    function GetCurrentPadding(out NumPaddingFrames: UINT32): HRESULT; stdcall;
    function IsFormatSupported(ShareMode: DWORD; pFormat: PWAVEFORMATEX; out pClosestMatch: PWAVEFORMATEX): HRESULT; stdcall;
    function GetMixFormat(out pFormat: PWAVEFORMATEX): HRESULT; stdcall;
    function GetDevicePeriod(out hmsDefaultDevicePeriod, hmsMinimumDevicePeriod: REFERENCE_TIME): HRESULT; stdcall;
    function Start: HRESULT; stdcall;
    function Stop: HRESULT; stdcall;
    function Reset: HRESULT; stdcall;
    function SetEventHandle(eventHandle: HWND): HRESULT; stdcall;
    function GetService(const iid: TGUID; out ppInterface: IUnknown): HRESULT; stdcall;
  end;


  IAudioCaptureClient = interface(IUnknown)
    ['{C8ADBD64-E71E-48a0-A4DE-185C395CD317}']
    function GetBuffer(out pData: PBYTE; out NumFramesToRead: UINT32; out dwFlags: DWORD;
      pu64DevicePosition: PUINT64; pu64QPCPosition: PUINT64): HResult; stdcall;
    function ReleaseBuffer(NumFramesRead: UINT32): HResult; stdcall;
    function GetNextPacketSize(out NumFramesInNextPacket: UINT32): HResult; stdcall;
  end;


  IAudioRenderClient = interface(IUnknown)
    ['{F294ACFC-3146-4483-A7BF-ADDCA7C260E2}']
    function GetBuffer(NumFramesRequested: UINT32; out pData: pByte): HResult; stdcall;
    function ReleaseBuffer(NumFramesWritten: UINT32; dwFlags: DWORD): HResult; stdcall;
  end;


  IAudioSessionEvents = interface(IUnknown)
    ['{24918ACC-64B3-37C1-8CA9-74A66E9957A8}']
    function OnDisplayNameChanged(NewDisplayName: LPCWSTR; EventContext: pGuid): HResult; stdcall;
    function OnIconPathChanged(NewIconPath: LPCWSTR; EventContext: pGuid): HResult; stdcall;
    function OnSimpleVolumeChanged(NewVolume: Single; NewMute: LongBool; EventContext: pGuid): HResult; stdcall;
    function OnChannelVolumeChanged(ChannelCount: DWORD; const NewChannelArray: array of Single; ChangedChannel: DWORD;
      EventContext: pGuid): HResult; stdcall;
    function OnGroupingParamChanged(NewGroupingParam, EventContext: pGuid): HResult; stdcall;
    function OnStateChanged(NewState: AudioSessionState): HResult; stdcall;
    function OnSessionDisconnected(DisconnectReason: AudioSessionDisconnectReason): HResult; stdcall;
  end;


  IAudioClock = interface(IUnknown)
    ['{CD63314F-3FBA-4A1B-812C-EF96358728E7}']
    function GetFrequency(out pu64Frequency: UINT64): HRESULT; stdcall;
    function GetPosition(out pu64Position: UINT64; out pu64QPCPosition: UINT64): HRESULT; stdcall;
    function GetCharacteristics(out pdwCharacteristics: DWORD): HRESULT; stdcall;
  end;


  IAudioClock2 = interface(IUnknown)
    ['{6F49FF73-6727-49AC-A008-D98CF5E70048}']
    function GetDevicePosition(out DevicePosition: UINT64; out QPCPosition: UINT64): HRESULT; stdcall;
  end;


  IAudioClockAdjustment = interface(IUnknown)
    ['{F6E4C0A0-46D9-4FB9-BE21-57A3EF2B626C}']
    function SetSampleRate(flSampleRate: single): HRESULT; stdcall;
  end;


  { ******* undocumented ******** }
  IRemoteAudioSession = interface(IUnknown)
    ['{33969B1D-D06F-4281-B837-7EAAFD21A9C0}']
    function func_a: HResult; stdcall;
    function func_b: HResult; stdcall;
    function func_c: HResult; stdcall;
    function func_d: HResult; stdcall;
    function func_e: HResult; stdcall;
    function func_f: HResult; stdcall;
    function func_g: HResult; stdcall;
    function func_h: HResult; stdcall;
    function func_i: HResult; stdcall;
    function func_j: HResult; stdcall;
    function func_k: HResult; stdcall;
    function GetProcessID(out pid: DWORD): HResult; stdcall;
  end;


  { ******* undocumented ******** }
  IAudioSessionQuerier = interface(IUnknown)
    ['{94BE9D30-53AC-4802-829C-F13E5AD34776}']
    function GetNumSessions(out NumSessions: DWORD): HResult; stdcall;
    function QuerySession(Num: DWORD; out Session: IUnknown): HResult; stdcall;
  end;


  { ******* undocumented ******** }
  IAudioSessionQuery = interface(IUnknown)
    ['{94BE9D30-53AC-4802-829C-F13E5AD34775}']
    function GetQueryInterface(out AudioQuerier: IAudioSessionQuerier): HResult; stdcall;
  end;
(* ---------------- WindowsAudioSession API End ---------------- *)



(* ------------------- EndpointVolume API ---------------------- *)
const
  ENDPOINT_HARDWARE_SUPPORT_VOLUME = $00000001;
  ENDPOINT_HARDWARE_SUPPORT_MUTE   = $00000002;
  ENDPOINT_HARDWARE_SUPPORT_METER  = $00000004;


type
  PAUDIO_VOLUME_NOTIFICATION_DATA = ^AUDIO_VOLUME_NOTIFICATION_DATA;

  AUDIO_VOLUME_NOTIFICATION_DATA = packed record
    guidEventContext: TGUID;
    bMuted          : BOOL;
    fMasterVolume   : Single;
    nChannels       : UINT;
    afChannelVolumes: array [0..0] of Single;
  end;


type
  // Forward Deklarationen
  IAudioEndpointVolumeCallback = interface;
  IAudioEndpointVolume         = interface;
  IAudioEndpointVolumeEx       = interface;
  IAudioMeterInformation       = interface;


  IAudioEndpointVolumeCallback = interface(IUnknown)
    ['{657804FA-D6AD-4496-8A60-352752AF4F89}']
    function OnNotify(pNotify:PAUDIO_VOLUME_NOTIFICATION_DATA):HRESULT; stdcall;
  end;


  IAudioEndpointVolume = interface(IUnknown)
    ['{5CDF2C82-841E-4546-9722-0CF74078229A}']
    function RegisterControlChangeNotify(pNotify: IAudioEndpointVolumeCallback): HRESULT; stdcall;
    function UnregisterControlChangeNotify(pNotify: IAudioEndpointVolumeCallback): HRESULT; stdcall;
    function GetChannelCount(out pnChannelCount: UINT): HRESULT; stdcall;
    function SetMasterVolumeLevel(fLevelDB: Single; pguidEventContext: PGuid): HRESULT; stdcall;
    function SetMasterVolumeLevelScalar(fLevel: Single; pguidEventContext: PGuid): HRESULT; stdcall;
    function GetMasterVolumeLevel(out pfLevelDB: Single): HRESULT; stdcall;
    function GetMasterVolumeLevelScalar(out pfLevel: Single): HRESULT; stdcall;
    function SetChannelVolumeLevel(nChannel: UINT; fLevelDB: Single; pguidEventContext: PGuid): HRESULT; stdcall;
    function SetChannelVolumeLevelScalar(nChannel: UINT; fLevel:Single; pguidEventContext: PGuid): HRESULT; stdcall;
    function GetChannelVolumeLevel(nChannel: UINT; fLevelDB: Single): HRESULT; stdcall;
    function GetChannelVolumeLevelScalar(nChannel: UINT; fLevel: Single): HRESULT; stdcall;
    function SetMute(bMute: BOOL; pguidEventContext: PGuid): HRESULT; stdcall;
    function GetMute(out pbMute: BOOL): HRESULT; stdcall;
    function GetVolumeStepInfo(out pnStep: UINT; out pnStepCount: UINT): HRESULT; stdcall;
    function VolumeStepUp(pguidEventContext: PGuid): HRESULT; stdcall;
    function VolumeStepDown(pguidEventContext: PGuid): HRESULT; stdcall;
    function QueryHardwareSupport(out pdwHardwareSupportMask: UINT): HRESULT; stdcall;
    function GetVolumeRange(out pflVolumeMindB: Single; out pflVolumeMaxdB: Single;
                            out pflVolumeIncrementdB: Single): HRESULT; stdcall;
  end;


  IAudioEndpointVolumeEx = interface(IUnknown)
    ['{66E11784-F695-4F28-A505-A7080081A78F}']
    function GetVolumeRangeChannel(iChannel: UINT; out pflVolumeMindB: Single; out pflVolumeMaxdB: Single; out pflVolumeIncrementdB: Single): HRESULT; stdcall;
  end;


  IAudioMeterInformation = interface(IUnknown)
    ['{C02216F6-8C67-4B5B-9D00-D008E73E0064}']
    function GetPeakValue(out fPeak: Single): HRESULT; stdcall;
    function GetMeteringChannelCount(out nChannelCount: UINT): HRESULT; stdcall;
    function GetChannelsPeakValues(u32ChannelCount: UINT; pfPeakValues: pSingle): HRESULT; stdcall;
    function QueryHardwareSupport(out dwHardwareSupportMask: UINT): HRESULT; stdcall;
  end;
(* ------------------ EndpointVolume API End ------------------- *)



implementation

end.
