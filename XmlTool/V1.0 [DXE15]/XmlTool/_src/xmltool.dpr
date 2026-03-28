{ **************************************************************************** }
//
// XmlTool, a Windows command-line utility for XML processing
//
// Written in Delphi 12.3 Athens.
// Supported from Delphi 11.1 Alexandria onwards.
//
// Target OS: MS Windows (Windows 7 and later)
// Author   : Andreas Heim, 2026-03
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

program xmltool;

{$APPTYPE CONSOLE}

{$R *.res}


uses
  Winapi.Windows, Winapi.ActiveX, Winapi.msxml,Winapi.MSXMLIntf,
  System.SysUtils, System.StrUtils, System.IOUtils, System.Classes,
  System.Generics.Collections,

  MSXMLIntfSupport in 'lib\MSXMLIntfSupport.pas',
  GpCommandLineParser in 'lib\GpCommandLineParser.pas';


// -----------------------------------------------------------------------------
// Messages
// -----------------------------------------------------------------------------

const
  csErrMsgUnhandledTask     = 'Unhandled task found';
  csErrMsgUnhandledRetCode  = 'Unhandled return code found';

  csErrMsgNoXmlFile         = 'Please provide an XML file.';
  csErrMsgXmlFileNotFound   = 'The specified XML file was not found.';
  csErrMsgXsdFileNotFound   = 'The specified schema file was not found.';
  csErrMsgXsltFileNotFound  = 'The specified XSLT file was not found.';
  csErrMsgNoXPathQuery      = 'Please provide an XPath expression.';
  csErrMsgLoadingXmlFailed  = 'Loading XML file failed.';
  csErrMsgValidationFailed  = 'XML validation failed.';
  csErrMsgXPathFailed       = 'The XPath query returned no matches.';
  csErrMsgTransformFailed   = 'The XSL transformation failed.';
  csErrMsgNoNamespacesFound = 'No namespaces were found.';

  csMsgValidationSuccess    = 'XML validation succeeded.';


// -----------------------------------------------------------------------------
// Basic enums
// -----------------------------------------------------------------------------

type
  TTaskToDo = (
    ttdNone,
    ttdValidate,
    ttdListNamespaces,
    ttdXPathQuery,
    ttdTransform,
    ttdPrettyPrint
  );

  TOutputTarget = (
    otStdOut,
    otStdErr
  );

  TResultCode = (
    rcOK,
    rcShowHelp,
    rcCmdLineError,
    rcNoXmlFile,
    rcXmlFileNotFound,
    rcXsdFileNotFound,
    rcXsltFileNotFound,
    rcNoXPathQuery,
    rcLoadingXmlFailed,
    rcValidationFailed,
    rcXPathFailed,
    rcTransformFailed,
    rcNoNamespacesFound
  );


// -----------------------------------------------------------------------------
// TSettings, declarative definition of command line interface
// -----------------------------------------------------------------------------

type
  TSettings = class
  strict private
    FTaskToDo:           TTaskToDo;

    FExecXsdValidation:  boolean;
    FListNamespaces:     boolean;
    FExecXPathQuery:     boolean;
    FExecXslTransform:   boolean;
    FPrettyPrint:        boolean;

    FSaveToFile:         boolean;
    FQuiet:              boolean;
    FShowHelp:           boolean;

    FXmlFilePath:        string;
    FXsdFilePath:        string;
    FXsltFilePath:       string;
    FOutFilePath:        string;
    FXPathQuery:         string;
    FTransformParams:    string;
    FFormatPretty:       boolean;
    FFormatLinearized:   boolean;
    FFormat:             boolean;

    FHasTransformParams: boolean;

    FReadFromStdIn:      boolean;
    FInputData:          IStream;

    FStdInIsTTY:         boolean;
    FStdOutIsTTY:        boolean;
    FStdErrIsTTY:        boolean;
    FConsoleOutputCP:    cardinal;

  public
    constructor Create;

    // Do not use CLPDescription because output of Usage function looks weird

    [CLPPosition(1), CLPDefault('-'), CLPRequired]
    property XmlFilePath: string read FXmlFilePath write FXmlFilePath;

    [CLPName('s'), CLPLongName('schema')]
    property XsdFilePath: string read FXsdFilePath write FXsdFilePath;

    [CLPName('l'), CLPLongName('list-namespaces', 'list')]
    property ListNamespaces: boolean read FListNamespaces write FListNamespaces;

    [CLPName('x'),  CLPLongName('xpath-query', 'xpath')]
    property XPathQuery: string read FXPathQuery write FXPathQuery;

    [CLPName('t'), CLPLongName('transform')]
    property XsltFilePath: string read FXsltFilePath write FXsltFilePath;

    [CLPName('p'), CLPLongName('transform-params', 'transform-p')]
    property TransformParams: string read FTransformParams write FTransformParams;

    [CLPName('f'), CLPLongName('format'), CLPExtendable]
    property Format: boolean read FFormat write FFormat;

    [CLPName('o'), CLPLongName('output', 'out')]
    property OutFilePath: string read FOutFilePath write FOutFilePath;

    [CLPName('q'), CLPLongName('quiet')]
    property Quiet: boolean read FQuiet write FQuiet;

    [CLPName('h'), CLPLongName('help')]
    property ShowHelp: boolean read FShowHelp write FShowHelp;

    property TaskToDo:           TTaskToDo read FTaskToDo           write FTaskToDo;

    property ExecXsdValidation:  boolean   read FExecXsdValidation  write FExecXsdValidation;
    property ExecXPathQuery:     boolean   read FExecXPathQuery     write FExecXPathQuery;
    property ExecXslTransform:   boolean   read FExecXslTransform   write FExecXslTransform;

    property PrettyPrint:        boolean   read FPrettyPrint        write FPrettyPrint;
    property FormatPretty:       boolean   read FFormatPretty       write FFormatPretty;
    property FormatLinearized:   boolean   read FFormatLinearized   write FFormatLinearized;

    property SaveToFile:         boolean   read FSaveToFile         write FSaveToFile;

    property HasTransformParams: boolean   read FHasTransformParams write FHasTransformParams;

    property ReadFromStdIn:      boolean   read FReadFromStdIn      write FReadFromStdIn;
    property InputData:          IStream   read FInputData          write FInputData;

    property StdInIsTTY:         boolean   read FStdInIsTTY         write FStdInIsTTY;
    property StdOutIsTTY:        boolean   read FStdOutIsTTY        write FStdOutIsTTY;
    property StdErrIsTTY:        boolean   read FStdErrIsTTY        write FStdErrIsTTY;
    property ConsoleOutputCP:    cardinal  read FConsoleOutputCP    write FConsoleOutputCP;

  end;


// -----------------------------------------------------------------------------
// TNamespaceInfo and TNamespaceMap
// -----------------------------------------------------------------------------

  {$UNDEF HAS_ORDERED_DICT}

  {$IFDEF CONDITIONALEXPRESSIONS}
    {$IF Declared(CompilerVersion)}
      {$IF CompilerVersion >= 36.0}
        {$IF Declared(RTLVersion122)}
           {$DEFINE HAS_ORDERED_DICT}
        {$IFEND}
      {$IFEND}
    {$IFEND}
  {$ENDIF}

  TNamespaceInfo = {$IFDEF HAS_ORDERED_DICT} TOrderedDictionary<string, integer> {$ELSE} TDictionary<string, integer> {$ENDIF};
  TNamespaceMap  = {$IFDEF HAS_ORDERED_DICT} TObjectOrderedDictionary<string, TNamespaceInfo> {$ELSE} TObjectDictionary<string, TNamespaceInfo> {$ENDIF};


var
  Settings: TSettings;


function  ParseCommandLine: TResultCode; forward;
function  CheckResult(const AValue: TResultCode; out ARetCode: integer): boolean; forward;
procedure ShowHelp; forward;

function  ValidateXml: TResultCode; forward;
function  ListNamespaces: TResultCode; forward;
function  ExecuteXPathQuery: TResultCode; forward;
function  TransformXml: TResultCode; forward;
function  PrettyPrint: TResultCode; forward;

function  LoadXmlDocument(const AFilePath: string; APreserveWhiteSpace: boolean; AResolveExternals: boolean; AFreeThreaded: boolean = false): IXMLDOMDocument2; overload; forward;
function  LoadXmlDocument(const AStream: IStream; APreserveWhiteSpace: boolean; AResolveExternals: boolean; AFreeThreaded: boolean = false): IXMLDOMDocument2; overload; forward;
function  LoadXmlString(const AXml: string; APreserveWhiteSpace: boolean; AResolveExternals: boolean; AFreeThreaded: boolean = false): IXMLDOMDocument2; overload; forward;
function  CheckXmlDocument(const ADoc: IXMLDOMDocument): boolean; forward;

function  ReadStdInToStream: IStream; forward;
procedure EmitText(const AText: string; const ATarget: TOutputTarget = otStdOut); forward;
procedure EmitStream(const AStream: TStringStream); forward;

function  GetDocEncoding(const ADoc: IXMLDOMDocument; const ADefaultEncoding: string): TEncoding; forward;
function  GetDocNamespaces(const ADoc: IXMLDOMDocument): TNamespaceMap; forward;



// *****************************************************************************
// Implementation
// *****************************************************************************

// -----------------------------------------------------------------------------
// TSettings
// -----------------------------------------------------------------------------

constructor TSettings.Create;
begin
  inherited;

  FTaskToDo           := ttdNone;

  FExecXsdValidation  := false;
  FListNamespaces     := false;
  FExecXPathQuery     := false;
  FExecXslTransform   := false;
  FPrettyPrint        := false;
  FFormatPretty       := false;
  FFormatLinearized   := false;

  FSaveToFile         := false;
  FQuiet              := false;
  FShowHelp           := false;

  FXmlFilePath        := '';
  FXsdFilePath        := '';
  FXsltFilePath       := '';
  FXPathQuery         := '';
  FTransformParams    := '';
  FOutFilePath        := '';

  FHasTransformParams := false;
  FReadFromStdIn      := false;
  FInputData          := nil;

  FStdInIsTTY         := GetFileType(GetStdHandle(STD_INPUT_HANDLE))  = FILE_TYPE_CHAR;
  FStdOutIsTTY        := GetFileType(GetStdHandle(STD_OUTPUT_HANDLE)) = FILE_TYPE_CHAR;
  FStdErrIsTTY        := GetFileType(GetStdHandle(STD_ERROR_HANDLE))  = FILE_TYPE_CHAR;
  FConsoleOutputCP    := GetConsoleOutputCP;
end;


// -----------------------------------------------------------------------------
// Command Line & Help
// -----------------------------------------------------------------------------

function ParseCommandLine: TResultCode;
begin
  Result := rcOK;

  // Command line parsing
  if not CommandLineParser.Parse(Settings) then
  begin
    EmitText(CommandLineParser.ErrorInfo.SwitchName + ': ' + CommandlineParser.ErrorInfo.Text, otStdErr);
    exit(rcCmdLineError);
  end;

  // Command line evaluation
  Settings.ExecXPathQuery     := not Settings.XPathQuery.IsEmpty;
  Settings.ExecXslTransform   := not Settings.XsltFilePath.IsEmpty;
  Settings.HasTransformParams := not Settings.TransformParams.IsEmpty;
  Settings.ReadFromStdIn      := Settings.XmlFilePath = '-';
  Settings.SaveToFile         := not Settings.OutFilePath.IsEmpty;

  // Validating an XML file is the default use case
  Settings.ExecXsdValidation  := not Settings.XsdFilePath.IsEmpty
                                 or (not Settings.ListNamespaces
                                     and not Settings.ExecXPathQuery
                                     and not Settings.ExecXslTransform
                                     and not Settings.Format
                                    );

  // Extendable parameter f/format evaluation
  if Settings.Format then
  begin
    Settings.PrettyPrint      := CommandLineParser.GetExtension('f') = '';
    Settings.FormatPretty     := IndexStr(CommandLineParser.GetExtension('f'), ['tp', '-transform-p', '-transform-pretty']) >= 0;
    Settings.FormatLinearized := IndexStr(CommandLineParser.GetExtension('f'), ['tl', '-transform-l', '-transform-linearized']) >= 0;

    if not Settings.PrettyPrint
       and not Settings.FormatPretty
       and not Settings.FormatLinearized then
    begin
      EmitText('-f' + CommandLineParser.GetExtension('f') + ': Unknown switch.', otStdErr);
      exit(rcCmdLineError);
    end;
  end;

  // Help
  if (ParamCount = 0)
     or Settings.ShowHelp then
    exit(rcShowHelp);

  // Task evaluation, in reverse priority order
  if Settings.ExecXslTransform then  // lowest priority
    Settings.TaskToDo := ttdTransform;

  if Settings.ExecXPathQuery then
     Settings.TaskToDo := ttdXPathQuery;

  if Settings.ListNamespaces then
   Settings.TaskToDo := ttdListNamespaces;

  if Settings.PrettyPrint then
    Settings.TaskToDo := ttdPrettyPrint;

  if Settings.FormatPretty
     or Settings.FormatLinearized then
    Settings.TaskToDo := ttdTransform;

  if Settings.ExecXsdValidation then  // highest priority
    Settings.TaskToDo := ttdValidate;

  // Validation checks
  if Settings.XmlFilePath = '' then  // XML file path is mandatory
    exit(rcNoXmlFile);

  if not Settings.ReadFromStdIn and not FileExists(Settings.XmlFilePath) then
    exit(rcXmlFileNotFound);

  // Task specific checks
  case Settings.TaskToDo of
    ttdValidate:
    begin
      if (Settings.XsdFilePath <> '') and not FileExists(Settings.XsdFilePath) then
        exit(rcXsdFileNotFound);
    end;

    ttdXPathQuery:
    begin
      if Settings.XPathQuery = '' then
        exit(rcNoXPathQuery);
    end;

    ttdTransform:
    begin
      if not Settings.FormatPretty and not Settings.FormatLinearized then
      begin
        if Settings.XsltFilePath = '' then
          exit(rcXsltFileNotFound);

        if not FileExists(Settings.XsltFilePath) then
          exit(rcXsltFileNotFound);
      end;
    end;
  end;
end;


function CheckResult(const AValue: TResultCode; out ARetCode: integer): boolean;
begin
  Result := false;

  case AValue of
    rcOK:
    begin
      ARetCode := 0;
      Result := true;
    end;

    rcShowHelp:
    begin
      ARetCode := 0;
      ShowHelp;
    end;

    rcCmdLineError:
    begin
      ARetCode := 1;
    end;

    rcNoXmlFile:
    begin
      ARetCode := 2;
      EmitText(csErrMsgNoXmlFile, otStdErr);
    end;

    rcXmlFileNotFound:
    begin
      ARetCode := 3;
      EmitText(csErrMsgXmlFileNotFound, otStdErr);
    end;

    rcXsdFileNotFound:
    begin
      ARetCode := 4;
      EmitText(csErrMsgXsdFileNotFound, otStdErr);
    end;

    rcXsltFileNotFound:
    begin
      ARetCode := 5;
      EmitText(csErrMsgXsltFileNotFound, otStdErr);
    end;

    rcNoXPathQuery:
    begin
      ARetCode := 6;
      EmitText(csErrMsgNoXPathQuery, otStdErr);
    end;

    rcLoadingXmlFailed:
    begin
      ARetCode := 7;
      EmitText(csErrMsgLoadingXmlFailed, otStdErr);
    end;

    rcValidationFailed:
    begin
      ARetCode := 8;
      if not Settings.Quiet then
        EmitText(csErrMsgValidationFailed, otStdErr);
    end;

    rcXPathFailed:
    begin
      ARetCode := 9;
      if not Settings.Quiet then
        EmitText(csErrMsgXPathFailed, otStdErr);
    end;

    rcTransformFailed:
    begin
      ARetCode := 10;
      if not Settings.Quiet then
        EmitText(csErrMsgTransformFailed, otStdErr);
    end;

    rcNoNamespacesFound:
    begin
      ARetCode := 11;
      if not Settings.Quiet then
        EmitText(csErrMsgNoNamespacesFound, otStdErr);
    end;

    else
    begin
      ARetCode := 20;
      Assert(false, csErrMsgUnhandledRetCode);
    end
  end;
end;


procedure ShowHelp;
var
  ExeName: string;
  PadStr:  string;
  OutStr:  string;

begin
  ExeName := TPath.GetFileNameWithoutExtension(ParamStr(0));
  PadStr  := StringOfChar(' ', Length(ExeName));

  OutStr  := ExeName   + ' [-q] [-s:<XsdFile>] <XmlFile>' + sLineBreak
             + PadStr  + ' [-q] -l <XmlFile>' + sLineBreak
             + PadStr  + ' [-q] -x:<XPathQuery> <XmlFile> [-o:<OutputFile>]' + sLineBreak
             + PadStr  + ' [-q] -t:<XslFile> [-p:<Parameters>] <XmlFile> [-o:<OutputFile>]' + sLineBreak
             + PadStr  + ' [-q] -f <XmlFile> [-o:<OutputFile>]' + sLineBreak
             + PadStr  + ' [-q] -ftp <XmlFile> [-o:<OutputFile>]' + sLineBreak
             + PadStr  + ' [-q] -ftl <XmlFile> [-o:<OutputFile>]' + sLineBreak
             + PadStr  + ' -h' + sLineBreak
             + sLineBreak
             + '<XmlFile>                           XML file to process (or "-" for StdIn, default when omitted)' + sLineBreak
             + sLineBreak
             + '-s:<XsdFile>  --schema:<XsdFile>    Validate the XML file against the XSD schema file (schemaLocation and' + sLineBreak
             + '                                    noNamespaceSchemaLocation attributes in the XML file are ignored),' + sLineBreak
             + '                                    exit code 0 indicates successful validation' + sLineBreak
             + sLineBreak
             + '-l  --list  --list-namespaces       List namespaces used in the XML file, exit code 0 indicates namespaces' + sLineBreak
             + '                                    were found. Similar output as when using -x:"/*/namespace::*"' + sLineBreak
             + sLineBreak
             + '-x:<XPathQuery>                     Execute the XPath query on the XML file, exit code 0 indicates' + sLineBreak
             + '  --xpath:<XPathQuery>              successful execution' + sLineBreak
             + '  --xpath-query:<XPathQuery>' + sLineBreak
             + sLineBreak
             + '-t:<XslFile>                        Execute an XSL transformation on the XML file using the stylesheet' + sLineBreak
             + '  --transform:<XslFile>             from the XSL file' + sLineBreak
             + sLineBreak
             + '-p:<Parameters>                     Transformation parameters in name=''value'' format (space-separated,' + sLineBreak
             + '  --transform-p:<Parameters>        only with -t)' + sLineBreak
             + '  --transform-params:<Parameters>' + sLineBreak
             + sLineBreak
             + '-f  --format                        Pretty-print XML file using SAX reader/writer' + sLineBreak
             + sLineBreak
             + '-ftp  --format-transform-p          Pretty-print XML file using an internal XSLT style sheet' + sLinebreak
             + '  --format-transform-pretty' + sLinebreak
             + sLineBreak
             + '-ftl  --format-transform-l          Linearize XML file using an internal XSLT style sheet' + sLinebreak
             + '  --format-transform-linearized' + sLinebreak
             + sLineBreak
             + '-o:<OutputFile>                     Write output to file (only with -x, -t, -f, -ftp and -ftl)' + sLineBreak
             + '  --out:<OutputFile>' + sLineBreak
             + '  --output:<OutputFile>' + sLineBreak
             + sLineBreak
             + '-q  --quiet                         Suppress all output, only set exit code' + sLineBreak
             + sLineBreak
             + '-h  --help                          Show this help';

  EmitText(OutStr);
end;


// -----------------------------------------------------------------------------
// Schema Validation
// -----------------------------------------------------------------------------

function ValidateXml: TResultCode;
var
  XmlDoc:      IXMLDOMDocument2;
  XsdDoc:      IXMLDOMDocument2;
  SchemaCache: IXMLDOMSchemaCollection;
  ParseError:  IXMLDOMParseError;
  Attr:        IXMLDOMNode;
  TargetNS:    string;
  Msg:         string;

begin
  // Load XML
  // "preserveWhiteSpace" must be TRUE so that position information in error messages is correct
  // "resolveExternals" must be TRUE so that schema files referenced in the "schemaLocation" attribute are processed
  if Settings.ReadFromStdIn then
    XmlDoc := LoadXmlDocument(Settings.InputData, true, true)
  else
    XmlDoc := LoadXmlDocument(Settings.XmlFilePath, true, true);

  if not CheckXmlDocument(XmlDoc) then
    exit(rcLoadingXmlFailed);

  // Create schema cache
  SchemaCache := CreateSchemaCache;

  if Settings.XsdFilePath <> '' then
  begin
    // Remove xsi:schemaLocation and xsi:noNamespaceSchemaLocation from the DOM
    // so that only the explicitly specified schema file is used
    Attr := XmlDoc.documentElement.attributes.getNamedItem('xsi:schemaLocation');
    if Assigned(Attr) then
      XmlDoc.documentElement.attributes.removeNamedItem('xsi:schemaLocation');

    Attr := XmlDoc.documentElement.attributes.getNamedItem('xsi:noNamespaceSchemaLocation');
    if Assigned(Attr) then
      XmlDoc.documentElement.attributes.removeNamedItem('xsi:noNamespaceSchemaLocation');

    // Load XSD
    // "resolveExternals" must be TRUE so that additional schema files included via "xs:include" tags are processed
    XsdDoc := LoadXmlDocument(Settings.XsdFilePath, false, true);

    if not CheckXmlDocument(XsdDoc) then
      exit(rcLoadingXmlFailed);

    // Determine targetNamespace from the XSD
    Attr := XsdDoc.documentElement.attributes.getNamedItem('targetNamespace');

    if Assigned(Attr) then
      TargetNS := Attr.nodeValue
    else
      TargetNS := '';

    // Add schema to cache
    SchemaCache.add(TargetNS, XsdDoc);

    // Set schema cache
    XmlDoc.schemas := SchemaCache;
  end;

  // Execute validation
  ParseError := XmlDoc.validate;

  if ParseError.errorCode = 0 then
  begin
    Result := rcOK;

    if not Settings.Quiet then
      EmitText(csMsgValidationSuccess, otStdOut);
  end
  else
  begin
    Result := rcValidationFailed;

    // Output validation result
    if not Settings.Quiet then
    begin
      Msg := 'Error';

      if ParseError.url <> '' then
        Msg := Msg + Format(' in file %s,', [ParseError.url]);

      if ParseError.line > 0 then
        Msg := Msg + Format(' at line %d', [ParseError.line]);

      if ParseError.linePos > 0 then
        if ParseError.line > 0 then
          Msg := Msg + Format(', position %d', [ParseError.linePos])
        else
          Msg := Msg + Format(' at position %d', [ParseError.linePos]);

      Msg := Msg + Format(', code 0x%.8x, reason: %s', [ParseError.errorCode, ParseError.reason]);

      if ParseError.srcText <> '' then
        Msg := Msg + Copy(ParseError.srcText, 1, 40);

      EmitText(Msg, otStdErr);
    end;
  end;
end;


// -----------------------------------------------------------------------------
// Namespace Listing
// -----------------------------------------------------------------------------

function ListNamespaces: TResultCode;
var
  XMLDoc:       IXMLDOMDocument2;
  Namespaces:   TNamespaceMap;
  NamespaceIDs: TNamespaceInfo;
  Msg:          TStringBuilder;
  NSN:          string;
  NSUri:        string;

begin
  // Load XML
  if Settings.ReadFromStdIn then
    XMLDoc := LoadXmlDocument(Settings.InputData, false, false)
  else
    XMLDoc := LoadXmlDocument(Settings.XmlFilePath, false, false);

  if not CheckXmlDocument(XmlDoc) then
    exit(rcLoadingXmlFailed);

  Namespaces := GetDocNamespaces(XMLDoc);

  try
    if Namespaces.Count = 0 then
      exit(rcNoNamespacesFound);

    Result := rcOK;

    // Output namespaces
    if not Settings.Quiet then
    begin
      Msg := TStringBuilder.Create;

      try
        for NSN in Namespaces.Keys do
        begin
          NamespaceIDs := Namespaces[NSN];

          for NSUri in NamespaceIDs.Keys do
          begin
            if Msg.Length > 0 then
              Msg.AppendLine;

            Msg.AppendFormat('%s = %s', [NSN, NSUri]);
          end;
        end;

        EmitText(Msg.ToString);

      finally
        Msg.Free;
      end;
    end;

  finally
    Namespaces.Free;
  end;
end;


// -----------------------------------------------------------------------------
// XPath-Query
// -----------------------------------------------------------------------------

function ExecuteXPathQuery: TResultCode;
var
  XMLDoc:       IXMLDOMDocument2;
  NodeList:     IXMLDOMNodeList;
  Namespaces:   TNamespaceMap;
  NamespaceIDs: TNamespaceInfo;
  OutStream:    TStringStream;
  Msg:          TStringBuilder;
  DocEnc:       TEncoding;
  NSN:          string;
  NSUri:        string;
  Idx:          integer;

begin
  // Load XML
  // "preserveWhiteSpace" must be TRUE so that the output of matching nodes is well-formatted
  // "resolveExternals" must be TRUE so that schema files referenced in the "schemaLocation" attribute are processed
  if Settings.ReadFromStdIn then
    XMLDoc := LoadXmlDocument(Settings.InputData, true, true)
  else
    XMLDoc := LoadXmlDocument(Settings.XmlFilePath, true, true);

  if not CheckXmlDocument(XmlDoc) then
    exit(rcLoadingXmlFailed);

  // Determine XML file encoding
  DocEnc := GetDocEncoding(XMLDoc, 'UTF-8');

  OutStream := TStringStream.Create('', DocEnc, true);  // Manages lifetime of TEncoding

  try
    // Detect namespaces and build SelectionNamespaces string
    Namespaces := GetDocNamespaces(XMLDoc);

    Msg := TStringBuilder.Create;

    try
      try
        for NSN in Namespaces.Keys do
        begin
          NamespaceIDs := Namespaces[NSN];

          if NamespaceIDs.Count = 1 then
            Msg.AppendFormat(' %s="%s"', [NSN, NamespaceIDs.Keys.ToArray[0]])
          else
            for NSUri in NamespaceIDs.Keys do
              Msg.AppendFormat(' %s%d="%s"', [NSN, NamespaceIDs[NSUri], NSUri]);
        end;

      finally
        Namespaces.Free;
      end;

      // Set SelectionNamespaces
      XMLDoc.setProperty('SelectionNamespaces', Trim(Msg.ToString));

    finally
      Msg.Free;
    end;

    // Execute XPath query
    NodeList := XMLDoc.selectNodes(Settings.XPathQuery);

    if NodeList.length = 0 then
      exit(rcXPathFailed);

    Result := rcOK;

    // Output query result
    if not Settings.Quiet then
    begin
      Msg := TStringBuilder.Create;

      try
        for Idx := 0 to Pred(NodeList.length) do
        begin
          if Msg.Length > 0 then
            Msg.AppendLine;

          Msg.Append(NodeList[Idx].xml);
        end;

        OutStream.WriteString(Msg.ToString);

      finally
        Msg.Free;
      end;

      EmitStream(OutStream);
    end;

  finally
    OutStream.Free;
  end;
end;


// -----------------------------------------------------------------------------
// XSL Transformation
// -----------------------------------------------------------------------------

function TransformXml: TResultCode;
const
  cXSLT_PrettyPrint = '<?xml version="1.0" encoding="UTF-8"?>'
                      + '<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">'
                      + '<xsl:output method="xml" omit-xml-declaration="no" /><xsl:param name="indent-increment" select="''&#x9;''" />'
                      + '<xsl:template match="*"><xsl:param name="indent" select="''&#xA;''" /><xsl:value-of select="$indent" />'
                      + '<xsl:copy><xsl:copy-of select="@*" /><xsl:apply-templates><xsl:with-param name="indent" select="concat($indent, $indent-increment)" />'
                      + '</xsl:apply-templates><xsl:if test="*"><xsl:value-of select="$indent" /></xsl:if></xsl:copy></xsl:template>'
                      + '<xsl:template match="comment()|processing-instruction()"><xsl:copy /></xsl:template><xsl:template match="text()[normalize-space(.)='''']" />'
                      + '</xsl:stylesheet>';

  cXSLT_Linearize   = '<?xml version="1.0" encoding="UTF-8"?>'
                      + '<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"><xsl:output method="xml" omit-xml-declaration="no" indent="no" />'
                      + '<xsl:template match="@*|node()"><xsl:copy><xsl:apply-templates select="@*|node()" /></xsl:copy>'
                      + '</xsl:template><xsl:template match="text()"><xsl:value-of select="normalize-space()" /></xsl:template>'
                      + '</xsl:stylesheet>';

var
  XMLDoc:        IXMLDOMDocument2;
  XSLDoc:        IXMLDOMDocument2;
  XSLTemplate:   IXSLTemplate;
  XSLProc:       IXSLProcessor;
  Node:          IXMLDOMNode;
  Attr:          IXMLDOMNode;
  OutStream:     TStringStream;
  OleStream:     IStream;
  DocEnc:        TEncoding;
  TargetEncName: string;
  Params:        TDictionary<string, string>;
  ParamDefs:     TArray<string>;
  ParamDef:      TArray<string>;
  ParamDefStr:   string;
  Param:         string;

begin
  // Load XML
  if Settings.ReadFromStdIn then
    XMLDoc := LoadXmlDocument(Settings.InputData, false, true)
  else
    XMLDoc := LoadXmlDocument(Settings.XmlFilePath, false, true);

  if not CheckXmlDocument(XmlDoc) then
    exit(rcLoadingXmlFailed);

  if Settings.FormatPretty then
    XSLDoc := LoadXmlString(cXSLT_PrettyPrint, false, true, true)
  else if Settings.FormatLinearized then
    XSLDoc := LoadXmlString(cXSLT_Linearize, false, true, true)
  else
    XSLDoc := LoadXmlDocument(Settings.XsltFilePath, false, true, true);

  if not CheckXmlDocument(XSLDoc) then
    exit(rcLoadingXmlFailed);

  // Determine XML file encoding
  DocEnc := GetDocEncoding(XMLDoc, 'UTF-8');

  try
    // Determine target encoding from XSL file
    Node := XSLDoc.documentElement.selectSingleNode('./*[local-name()=''output'']');

    if not Assigned(Node) then
    begin
      Node := XSLDoc.createNode(NODE_ELEMENT, 'output', 'http://www.w3.org/1999/XSL/Transform');
      XSLDoc.documentElement.insertBefore(Node, XSLDoc.documentElement.firstChild);
    end;

    // If the XSL file does not specify a target encoding, the source file's encoding is used
    // Otherwise the target encoding is read and, if necessary, a TEncoding object with that encoding is created
    Attr := Node.attributes.getNamedItem('encoding');

    if not Assigned(Attr) then
    begin
      Attr := XSLDoc.createNode(NODE_ATTRIBUTE, 'encoding', '');
      Node.attributes.setNamedItem(Attr);

      Attr.nodeValue := DocEnc.MIMEName;
    end
    else
    begin
      TargetEncName := Attr.nodeValue;

      if not SameText(DocEnc.MIMEName, TargetEncName) then
      begin
        FreeAndNil(DocEnc);
        DocEnc := TEncoding.GetEncoding(TargetEncName);
      end;
    end;

    // Create all required objects and do the wiring
    OutStream := TStringStream.Create('', DocEnc, false);  // Lifetime management of TEncoding is done by try..finally
    OleStream := TStreamAdapter.Create(OutStream, soOwned) as IStream;  // Manages lifetime of TStringStream

    XSLTemplate            := CreateXSLTemplate;
    XSLTemplate.styleSheet := XSLDoc;

    XSLProc        := XSLTemplate.createProcessor;
    XSLProc.input  := XMLDoc;
    XSLProc.output := OleStream;

    // Process transformation parameters
    if Settings.HasTransformParams then
    begin
      ParamDefs := Settings.TransformParams.Split([' '], '''', '''', TStringSplitOptions.ExcludeEmpty);
      Params    := TDictionary<string, string>.Create(Length(ParamDefs));

      try
        for ParamDefStr in ParamDefs do
        begin
          ParamDef := ParamDefStr.Trim.Split(['='], '''', '''', TStringSplitOptions.ExcludeEmpty);

          if Length(ParamDef) = 2 then
            Params.AddOrSetValue(ParamDef[0].Trim, ParamDef[1].Trim.DeQuotedString(''''));
        end;

        // Add them to XSL processor
        for Param in Params.Keys do
          XSLProc.addParameter(Param, Params[Param], '');

      finally
        Params.Free;
      end;
    end;

    // Execute transformation and error check
    if not XSLProc.transform then
      exit(rcTransformFailed);

    Result := rcOK;

    // Output result
    if not Settings.Quiet then
      EmitStream(OutStream);

  finally
    DocEnc.Free;
  end;
end;


// -----------------------------------------------------------------------------
// Pretty-Printing
// -----------------------------------------------------------------------------

function PrettyPrint: TResultCode;
var
  XMLDoc:     IXMLDOMDocument2;
  OutStream:  TStringStream;
  OleStream:  IStream;
  XmlWriter:  IMXWriter;
  XmlReader:  IVBSAXXMLReader;
  DocEnc:     TEncoding;
  StandAlone: string;

begin
  Result := rcOK;

  // Load XML
  if Settings.ReadFromStdIn then
    XMLDoc := LoadXmlDocument(Settings.InputData, false, false)
  else
    XMLDoc := LoadXmlDocument(Settings.XmlFilePath, false, false);

  if not CheckXmlDocument(XmlDoc) then
    exit(rcLoadingXmlFailed);

  // Determine standalone attribute
  try
    StandAlone := XMLDoc.firstChild.attributes.getNamedItem('standalone').nodeValue;
  except
    StandAlone := 'yes';  // Emergency fallback
  end;

  // Determine XML file encoding
  DocEnc := GetDocEncoding(XMLDoc, 'UTF-8');

  OutStream := TStringStream.Create('', DocEnc, true);  // Manages lifetime of TEncoding
  OleStream := TStreamAdapter.Create(OutStream, soOwned) as IStream;  // Manages lifetime of TStringStream

  XmlWriter                       := CreateMXXMLWriter;
  XmlWriter.omitXMLDeclaration    := false;
  XmlWriter.standalone            := SameText(StandAlone, 'yes');
  XmlWriter.byteOrderMark         := false;
  XmlWriter.disableOutputEscaping := false;
  XmlWriter.encoding              := DocEnc.MIMEName;
  XmlWriter.indent                := true;
  XmlWriter.output                := OleStream;

  XmlReader                       := CreateSAXXMLReader;
  XmlReader.contentHandler        := XmlWriter as IVBSAXContentHandler;
  XmlReader.dtdHandler            := XmlWriter as IVBSAXDTDHandler;
  XmlReader.errorHandler          := XmlWriter as IVBSAXErrorHandler;
  XmlReader.putProperty('http://xml.org/sax/properties/lexical-handler', XmlWriter);
  XmlReader.putProperty('http://xml.org/sax/properties/declaration-handler', XmlWriter);

  // Execute pretty-print
  XmlReader.parse(XMLDoc.xml);

  // Output result
  if not Settings.Quiet then
    EmitStream(OutStream);
end;


// -----------------------------------------------------------------------------
// Helper Functions
// -----------------------------------------------------------------------------

function LoadXmlDocument(const AFilePath: string; APreserveWhiteSpace: boolean; AResolveExternals: boolean; AFreeThreaded: boolean = false): IXMLDOMDocument2;
begin
  Result                    := CreateDOMDocument(AFreeThreaded) as IXMLDOMDocument2;
  Result.async              := false;
  Result.preserveWhiteSpace := APreserveWhiteSpace;
  Result.validateOnParse    := false;  // Setting "validateOnParse" to TRUE only makes sense if the schema cache has already been set via the "schemas" property before calling "load"
  Result.resolveExternals   := AResolveExternals;
  Result.load(AFilePath);
end;


function LoadXmlDocument(const AStream: IStream; APreserveWhiteSpace: boolean; AResolveExternals: boolean; AFreeThreaded: boolean = false): IXMLDOMDocument2;
begin
  Result                    := CreateDOMDocument(AFreeThreaded) as IXMLDOMDocument2;
  Result.async              := false;
  Result.preserveWhiteSpace := APreserveWhiteSpace;
  Result.validateOnParse    := false;
  Result.resolveExternals   := AResolveExternals;
  Result.load(AStream);
end;


function LoadXmlString(const AXml: string; APreserveWhiteSpace: boolean; AResolveExternals: boolean; AFreeThreaded: boolean = false): IXMLDOMDocument2;
begin
  Result                    := CreateDOMDocument(AFreeThreaded) as IXMLDOMDocument2;
  Result.async              := false;
  Result.preserveWhiteSpace := APreserveWhiteSpace;
  Result.validateOnParse    := false;
  Result.resolveExternals   := AResolveExternals;
  Result.loadXml(AXml);
end;


function CheckXmlDocument(const ADoc: IXMLDOMDocument): boolean;
begin
  if ADoc.parseError.errorCode = 0 then
    exit(true)
  else
  begin
    EmitText(Trim(ADoc.parseError.reason), otStdErr);
    exit(false);
  end;
end;


function ReadStdInToStream: IStream;
var
  InStream:  THandleStream;
  Buffer:    TMemoryStream;
  Chunk:     TBytes;
  BytesRead: integer;

begin
  SetLength(Chunk, 4096);

  Buffer   := nil;
  InStream := THandleStream.Create(GetStdHandle(STD_INPUT_HANDLE));

  try
    Buffer := TMemoryStream.Create;  // Lifetime managed by TStreamAdapter
    Result := TStreamAdapter.Create(Buffer, soOwned) as IStream;

    // Read StdIn blockwise
    repeat
      BytesRead := InStream.Read(Chunk, Length(Chunk));

      if BytesRead > 0 then
        Buffer.WriteBuffer(Chunk, BytesRead);
    until BytesRead = 0;

  finally
    if Assigned(Buffer) then
      Buffer.Position := 0;

    InStream.Free;
  end;
end;


procedure EmitText(const AText: string; const ATarget: TOutputTarget = otStdOut);
var
  OutStream: THandleStream;
  OutBytes:  TBytes;
  OutEnc:    TEncoding;
  IsTTY:     boolean;
  OutHandle: cardinal;

begin
  if ATarget = otStdOut then
  begin
    IsTTY     := Settings.StdOutIsTTY;
    OutHandle := STD_OUTPUT_HANDLE;
  end
  else
  begin
    IsTTY     := Settings.StdErrIsTTY;
    OutHandle := STD_ERROR_HANDLE;
  end;

  if IsTTY then
  begin
    if ATarget = otStdOut then
      Writeln(Output, sLineBreak + AText)
    else
      Writeln(ErrOutput, sLineBreak + AText);
  end
  else
  begin
    OutEnc := TEncoding.GetEncoding(Settings.ConsoleOutputCP);

    try
      OutBytes := OutEnc.GetBytes(AText + sLineBreak);
    finally
      OutEnc.Free;
    end;

    OutStream := THandleStream.Create(GetStdHandle(OutHandle));

    try
      OutStream.WriteBuffer(OutBytes, Length(OutBytes));
    finally
      OutStream.Free;
    end;
  end;
end;


procedure EmitStream(const AStream: TStringStream);
var
  OutStream: TStream;

begin
  if Settings.SaveToFile then
    OutStream := TFileStream.Create(Settings.OutFilePath, fmCreate or fmShareExclusive)

  else if not Settings.StdOutIsTTY then
    OutStream := THandleStream.Create(GetStdHandle(STD_OUTPUT_HANDLE))

  else
    OutStream := nil;

  if not Assigned(OutStream) then
    Writeln(Output, sLineBreak + AStream.DataString)
  else
    try
      AStream.Position := 0;
      OutStream.CopyFrom(AStream, AStream.Size);
    finally
      OutStream.Free;
    end;
end;


function GetDocEncoding(const ADoc: IXMLDOMDocument; const ADefaultEncoding: string): TEncoding;
var
  DocEncName: string;

begin
  try
    DocEncName := ADoc.firstChild.attributes.getNamedItem('encoding').nodeValue;
    Result     := TEncoding.GetEncoding(DocEncName);
  except
    Result     := TEncoding.GetEncoding(ADefaultEncoding);  // Emergency fallback, i.e. no encoding attribute or unsupported encoding name
  end;
end;


function GetDocNamespaces(const ADoc: IXMLDOMDocument): TNamespaceMap;
var
  NodeList:    IXMLDOMNodeList;
  Node:        IXMLDOMNode;
  Idx:         integer;
  NSN:         string;
  NSUri:       string;
  IsRegularNS: boolean;

begin
  Result := TNamespaceMap.Create([doOwnsValues]);

  // Extract all namespaces from all XML nodes
  try
    NodeList := ADoc.selectNodes('//namespace::*');
  except
    NodeList := nil;
  end;

  if not Assigned(NodeList) then
    exit;

  for Idx := 0 to Pred(NodeList.length) do
  begin
    Node        := NodeList.item[Idx];
    NSN         := StringReplace(Node.nodeName, 'xmlns:', '', []);
    IsRegularNS := false;

    // The namespace alias "xml" must not be used, so set a custom alias
    if NSN = 'xml' then
    begin
      NSN   := 'xmlns:MyXml';
      NSUri := Node.nodeValue;
    end

    // Assign a custom alias to all default namespaces
    else if NSN = 'xmlns' then
    begin
      NSN   := 'xmlns:DefaultNS';
      NSUri := Node.nodeValue;
    end

    // Register all other namespaces for XPath as well
    else
    begin
      NSN         := Node.nodeName;
      NSUri       := Node.nodeValue;
      IsRegularNS := true;
    end;

    // Use namespace alias as key for a dictionary, storing a sub-dictionary
    // with the namespace URI as key
    if not Result.ContainsKey(NSN) then
      Result.Add(NSN, TNamespaceInfo.Create([TPair<string, integer>.Create(NSUri, 1)]))

    // The boolean IsRegularNS allows skipping the dictionary key lookup
    // for regular namespace declarations
    else if not IsRegularNS and not Result[NSN].ContainsKey(NSUri) then
      Result[NSN].Add(NSUri, Succ(Result[NSN].Count));
  end;
end;



// *****************************************************************************
// Main Program
// *****************************************************************************

begin
  {$IFDEF DEBUG}
  {$WARN SYMBOL_PLATFORM OFF}
  ReportMemoryLeaksOnShutdown := (DebugHook <> 0);
  {$WARN SYMBOL_PLATFORM ON}
  {$ENDIF}

  Settings := TSettings.Create;

  try
    try
      CoInitialize(nil);

      try
        if CheckResult(ParseCommandLine, ExitCode) then
        begin
          if Settings.ReadFromStdIn then
            Settings.InputData := ReadStdInToStream;

          case Settings.TaskToDo of
            ttdValidate:       CheckResult(ValidateXml, ExitCode);
            ttdListNamespaces: CheckResult(ListNamespaces, ExitCode);
            ttdXPathQuery:     CheckResult(ExecuteXPathQuery, ExitCode);
            ttdTransform:      CheckResult(TransformXml, ExitCode);
            ttdPrettyPrint:    CheckResult(PrettyPrint, ExitCode);
            else               Assert(false, csErrMsgUnhandledTask);
          end;
        end

        // Flush input pipe to prevent "Stream write error" exception
        else if not Settings.StdInIsTTY then
          ReadStdInToStream;

      finally
        CoUninitialize;
      end;

    except
      on E: Exception do
      begin
        EmitText(E.ClassName + ': ' + E.Message, otStdErr);
        ExitCode := 100;
      end;
    end;

    {$IFDEF DEBUG}
    {$WARN SYMBOL_PLATFORM OFF}
    if DebugHook <> 0 then
      Readln;
    {$WARN SYMBOL_PLATFORM ON}
    {$ENDIF}

  finally
    Settings.Free;
  end;
end.
