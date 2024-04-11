{ **************************************************************************** }
//
// Partial port of propsys.h
//
// URL: https://learn.microsoft.com/en-us/windows/win32/api/propsys
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

unit PropSys2;


interface

uses
  Windows, ActiveX, ComObj;


type
  IInitializeWithFile = interface(IUnknown)
    ['{b7d14566-0509-4cce-a71f-0a554233bd9b}']
    function Initialize(pszFilePath: PAnsiChar; grfMode: DWORD): HRESULT; stdcall;
  end;


  IInitializeWithStream = interface(IUnknown)
    ['{b824b49d-22ac-4161-ac8a-9916e8fa3f7f}']
    function Initialize(var pIStream: IStream; grfMode: DWORD): HRESULT; stdcall;
  end;


  _tagpropertykey = packed record
    fmtid: TGUID;
    pid  : DWORD;
  end;

  PROPERTYKEY  = _tagpropertykey;
  PPropertyKey = ^TPropertyKey;
  TPropertyKey = _tagpropertykey;

  IPropertyStore = interface(IUnknown)
    ['{886d8eeb-8cf2-4446-8d02-cdba1dbdcf99}']
    function GetCount(out cProps: DWORD): HResult; stdcall;
    function GetAt(iProp: DWORD; out pKey: TPropertyKey): HResult; stdcall;
    function GetValue(const Key: TPropertyKey; out pPropVar: TPropVariant): HResult; stdcall;
    function SetValue(const Key: TPropertyKey; const pPropVar: TPropVariant): HResult; stdcall;
    function Commit: HResult; stdcall;
  end;


  IPropertyStoreCapabilities = interface(IUnknown)
    ['{c8e2d566-186e-4d49-bf41-6909ead56acc}']
    function IsPropertyWritable(pPropKey: PPropertyKey): HRESULT; stdcall;
  end;


const
  PKEY_DeviceInterface_FriendlyName: TPropertyKey = (fmtid: '{026E516E-B814-414B-83CD-856D6FEF4822}';
  {TPropVariant.vt = VT_LPWSTR}                      pid  : 2
                                                    );

  PKEY_Device_DeviceDesc           : TPropertyKey = (fmtid: '{A45C254E-DF1C-4EFD-8020-67D146A850E0}';
  {TPropVariant.vt = VT_LPWSTR}                      pid  : 2
                                                    );

  PKEY_Device_FriendlyName         : TPropertyKey = (fmtid: '{A45C254E-DF1C-4EFD-8020-67D146A850E0}';
  {TPropVariant.vt = VT_LPWSTR}                      pid  : 14
                                                    );

procedure PropVariantInit(var PV: TPropVariant);
function  PropVariantClear(PVar: PPropVariant): HRESULT; stdcall; external 'ole32.dll';



implementation


procedure PropVariantInit(var PV: TPropVariant);
begin
  PV.vt := VT_EMPTY;
end;


end.
