{ **************************************************************************** }
//
// MSXML bindings
//
// Written in Delphi 10 Seattle.
//
// Target OS: MS Windows (Windows 7 and later)
// Author   : Andreas Heim, 2017-01
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

unit MSXMLIntfSupport;


interface

uses
  Winapi.ActiveX, WinApi.msxml, Winapi.MSXMLIntf, System.Win.ComObj, Xml.xmldom,
  Xml.XMLConst;


function CreateDOMDocument(AFreeThreaded: boolean = false): IXMLDOMDocument;
function CreateSchemaCache: IXMLDOMSchemaCollection;
function CreateXSLTemplate: IXSLTemplate;
function CreateMXXMLWriter: IMXWriter;
function CreateSAXXMLReader: IVBSAXXMLReader;



implementation

function TryObjectCreate(const GUIDList: array of TGUID): IUnknown; forward;


function CreateDOMDocument(AFreeThreaded: boolean = false): IXMLDOMDocument;
begin
  if AFreeThreaded then
    Result := TryObjectCreate([CLASS_FreeThreadedDOMDocument60,
                               CLASS_FreeThreadedDOMDocument40,
                               CLASS_FreeThreadedDOMDocument30,
                               CLASS_FreeThreadedDOMDocument26,
                               Winapi.msxml.CLASS_FreeThreadedDOMDocument]
                             ) as IXMLDOMDocument
  else
    Result := TryObjectCreate([CLASS_DOMDocument60,
                               CLASS_DOMDocument40,
                               CLASS_DOMDocument30,
                               CLASS_DOMDocument26,
                               Winapi.msxml.CLASS_DOMDocument]
                             ) as IXMLDOMDocument;

  if not Assigned(Result) then
    raise DOMException.Create(SMSDOMNotInstalled);
end;


function CreateSchemaCache: IXMLDOMSchemaCollection;
begin
  Result := TryObjectCreate([CLASS_XMLSchemaCache60,
                             CLASS_XMLSchemaCache40,
                             CLASS_XMLSchemaCache30,
                             CLASS_XMLSchemaCache26,
                             Winapi.msxml.CLASS_XMLSchemaCache]
                           ) as IXMLDOMSchemaCollection;

  if not Assigned(Result) then
    raise DOMException.Create(SMSDOMNotInstalled);
end;


function CreateXSLTemplate: IXSLTemplate;
begin
  Result := TryObjectCreate([CLASS_XSLTemplate60,
                             CLASS_XSLTemplate40,
                             CLASS_XSLTemplate30,
                             CLASS_XSLTemplate26,
                             Winapi.msxml.CLASS_XSLTemplate]
                            ) as IXSLTemplate;

  if not Assigned(Result) then
    raise DOMException.Create(SMSDOMNotInstalled);
end;


function CreateMXXMLWriter: IMXWriter;
begin
  Result := TryObjectCreate([CLASS_MXXMLWriter60,
                             CLASS_MXXMLWriter40,
                             CLASS_MXXMLWriter30,
                             Winapi.msxml.CLASS_MXXMLWriter]
                            ) as IMXWriter;

  if not Assigned(Result) then
    raise DOMException.Create(SMSDOMNotInstalled);
end;


function CreateSAXXMLReader: IVBSAXXMLReader;
begin
  Result := TryObjectCreate([CLASS_SAXXMLReader60,
                             CLASS_SAXXMLReader40,
                             CLASS_SAXXMLReader30,
                             Winapi.msxml.CLASS_SAXXMLReader]
                            ) as IVBSAXXMLReader;

  if not Assigned(Result) then
    raise DOMException.Create(SMSDOMNotInstalled);
end;


function TryObjectCreate(const GUIDList: array of TGUID): IUnknown;
var
  I:      integer;
  Status: HResult;

begin
  Status := S_OK;

  for I := Low(GUIDList) to High(GUIDList) do
  begin
    Status := CoCreateInstance(GUIDList[I],
                               nil,
                               CLSCTX_INPROC_SERVER
                                 or CLSCTX_LOCAL_SERVER,
                               IDispatch,
                               Result
                              );
    if Status = S_OK then
      Exit;
  end;

  OleCheck(Status);
end;


end.

