# XmlTool

A Windows command-line utility for XML processing, built with Delphi 12.3 Athens.


## Requirements

- Microsoft Windows
- MSXML 3.0 or later
- Delphi 11.1 Alexandria for compiling to executable
- At least Delphi 12.2 Athens to benefit from ordered dictionary types


## Third-party Components

This project uses `GpCommandLineParser` by Primoz Gabrijelcic, see [GpDelphiUnits](https://github.com/gabr42/GpDelphiUnits).


## Features

- **Schema Validation** — Validate an XML document against an XSD schema
- **Namespace Listing** — List all namespaces used in an XML document
- **XPath Queries** — Execute XPath expressions and output matching nodes
- **XSLT Transformation** — Transform XML using an XSL stylesheet, with support for parameters
- **Pretty-Print (SAX)** — Reformat an XML document with proper indentation using SAX Reader/Writer
- **Pretty-Print (XSLT)** — Reformat an XML document with proper indentation using an internal XSLT stylesheet
- **Linearize (XSLT)** — Remove whitespace from an XML document using an internal XSLT stylesheet


## Usage

```
xmltool [-q] [-s:<XsdFile>] <XmlFile>
xmltool [-q] -l <XmlFile>
xmltool [-q] -x:<XPathQuery> <XmlFile> [-o:<OutputFile>]
xmltool [-q] -t:<XslFile> [-p:<Parameters>] <XmlFile> [-o:<OutputFile>]
xmltool [-q] -f <XmlFile> [-o:<OutputFile>]
xmltool [-q] -ftp <XmlFile> [-o:<OutputFile>]
xmltool [-q] -ftl <XmlFile> [-o:<OutputFile>]
xmltool -h
```

### Switches

| Switch | Description |
|--------|-------------|
| `<XmlFile>` | XML input file (positional, required). Use `-` to read from StdIn (default when omitted). |
| `-s:<file>` · `--schema:<file>` | Validate against the given XSD schema file. When provided, `schemaLocation` and `noNamespaceSchemaLocation` attributes in the XML are ignored. |
| `-l` · `--list` · `--list-namespaces` | List namespaces found in the XML document. |
| `-x:<expr>` · `--xpath:<expr>` · `--xpath-query:<expr>` | Execute an XPath query and output the results. |
| `-t:<file>` · `--transform:<file>` | Transform the XML using the given XSLT stylesheet. |
| `-p:<params>` · `--transform-p:<params>` · `--transform-params:<params>` | Pass parameters to the XSLT transformation in `name='value'` format (space-separated). Only valid with `-t`. |
| `-f` · `--format` | Pretty-print the XML document using SAX Reader/Writer. |
| `-ftp` · `--format-transform-p` · `--format-transform-pretty` | Pretty-print the XML document using an internal XSLT stylesheet. |
| `-ftl` · `--format-transform-l` · `--format-transform-linearized` | Linearize the XML document using an internal XSLT stylesheet. |
| `-o:<file>` · `--out:<file>` · `--output:<file>` | Write output to a file instead of StdOut. Only valid with `-x`, `-t`, `-f`, `-ftp` and `-ftl`. |
| `-q` · `--quiet` | Quiet mode — suppress all output, only set the exit code. |
| `-h` · `--help` | Show help text. |

Switch values are attached using `:` or `=` as delimiter. For short-form switches, the delimiter may be omitted (e.g. `-s:file`, `-s=file`, or `-sfile`). Long-form switches always require the delimiter (e.g. `--schema:file` or `--schema=file`).

### Default Behavior

When only the XML file path is provided without any task switch (`-l`, `-x`, `-t`, `-f`, `-ftp`, `-ftl`), schema validation is performed as the default action. If no `-s` switch is given, the XML document is expected to contain a `schemaLocation` or `noNamespaceSchemaLocation` attribute.


## Namespace Handling

For XPath queries, namespaces are automatically detected by evaluating `//namespace::*` on the loaded document. The collected namespaces are then registered via MSXML's `SelectionNamespaces` property so that XPath expressions can reference them.

Special cases:
- The reserved `xml` namespace is re-aliased to `MyXml`
- Default namespaces (declared without a prefix) are re-aliased to `DefaultNS`
- When the same prefix maps to multiple URIs, numeric suffixes are appended to disambiguate (e.g. `DefaultNS1`, `DefaultNS2`)


## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success — validation passed, XPath query matched, transformation produced output, namespaces found, or help displayed |
| 1 | Command line parsing error |
| 2 | No XML input file specified |
| 3 | XML input file not found |
| 4 | XSD schema file not found |
| 5 | XSLT stylesheet file not found |
| 6 | No XPath expression provided |
| 7 | Loading XML file failed (parse error) |
| 8 | XML schema validation failed |
| 9 | XPath query returned no matches |
| 10 | XSLT transformation failed (empty result) |
| 11 | No namespaces found in the XML document |
| 100 | Fatal error (unhandled exception) |


## Examples

Validate using embedded schema references (default action):
```
xmltool data.xml
```

Validate an XML file against a schema:
```
xmltool -s:schema.xsd data.xml
```

List all namespaces:
```
xmltool -l data.xml
```

Run an XPath query:
```
xmltool -x:"//book[@year>2000]/title" library.xml
```

Run an XPath query and write results to a file:
```
xmltool -x:"//book[@year>2000]/title" library.xml -o:results.xml
```

Transform with XSLT and write to a file:
```
xmltool -t:transform.xsl -p:"author='Hesse' year='1915'" input.xml -o:output.html
```

Pretty-print an XML file (SAX):
```
xmltool -f data.xml
```

Pretty-print an XML file (XSLT):
```
xmltool -ftp data.xml
```

Linearize an XML file:
```
xmltool -ftl data.xml
```

Pipe XML through two transformations:
```
xmltool -t:first.xsl input.xml | xmltool -t:second.xsl -o:final.xml
```

Quiet validation (check exit code only):
```
xmltool -q -s:schema.xsd data.xml
echo %ERRORLEVEL%
```
