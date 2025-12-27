Option Explicit

' Multilanguage.xls BdR (c)2012 Free to use
'
' This VBA script writes content of Excel sheet to translation files
' which can be used in XCode, Eclipse and Visual Studio.
' Left most column contains the translations keys,
' each column contains the translation values.
'
' Manage translations words and phrases in spreadsheet
' and use this macro to export to different source files
' * GenerateLocalisationJson - JSON files for javascript
' * GenerateLocalisationXcode - XCode Localizable.strings files for iPhone
' * GenerateLocalisationEclipse - XML strings.xml files for Android
' * GenerateLocalisationVisualStudio - .resx xml files for Visual Studio

Sub SaveToFile(sFilename As String, sText As String)

  Dim fsT As Object
  Set fsT = CreateObject("ADODB.Stream")
  fsT.Type = 2 'Specify stream type - we want To save text/string data.
  fsT.Charset = "utf-8" 'Specify charset For the source text data.
  fsT.Open 'Open the stream And write binary data To the object
  fsT.WriteText sText
  fsT.SaveToFile sFilename, 2 'Save binary data To disk

End Sub

Function FolderCreate(ByVal path As String) As Boolean

    FolderCreate = True
    Dim fso As New FileSystemObject ' to use, select Tools -> References -> Microsoft Scripting Runtime
    
    If FolderExists(path) Then
        Exit Function
    Else
        On Error GoTo DeadInTheWater
        fso.CreateFolder path ' could there be any error with this, like if the path is really screwed up?
        Exit Function
    End If

DeadInTheWater:
    MsgBox "A folder could not be created for the following path: " & path & ". Check the path name and try again."
    FolderCreate = False
    Exit Function

End Function

Function FolderExists(ByVal path As String) As Boolean

    FolderExists = False
    Dim fso As New FileSystemObject

    If fso.FolderExists(path) Then FolderExists = True

End Function

Function FileOrDirExists(PathName As String) As Boolean
     'Macro Purpose: Function returns TRUE if the specified file
     '               or folder exists, false if not.
     'PathName     : Supports Windows mapped drives or UNC
     '             : Supports Macintosh paths
     'File usage   : Provide full file path and extension
     'Folder usage : Provide full folder path
     '               Accepts with/without trailing "\" (Windows)
     '               Accepts with/without trailing ":" (Macintosh)
 
    Dim iTemp As Integer
 
     'Ignore errors to allow for error evaluation
    On Error Resume Next
    iTemp = GetAttr(PathName)
 
     'Check if error exists and set response appropriately
    Select Case Err.Number
    Case Is = 0
        FileOrDirExists = True
    Case Else
        FileOrDirExists = False
    End Select
 
     'Resume error checking
    On Error GoTo 0
End Function

Sub GenerateLocalisationCSV()
 
    ' write worksheet translations to CSV files (for example Godot game engine)
    Dim ws As Worksheet
    Dim x As Integer
    Dim y As Integer
    Dim strFilename As String
    Dim strKey As String
    Dim strValue As String
    Dim strContents As String
 
    ' create csv directory if not exist
    Const OUTPUT_DIR = "csv\"
  
    Call FolderCreate(ActiveWorkbook.path & "\" & OUTPUT_DIR)
    
    ' output filename
    strFilename = ActiveWorkbook.path & "\" & OUTPUT_DIR & "\multilanguage.csv"
    
    ' check all worksheets
    For Each ws In Worksheets
 
        ' only the active worksheet
        If ws.Name = ActiveSheet.Name Then
        
            ' build column header
            strContents = FormatCsvValue("Keys") & ","
            For x = 2 To ws.Cells.SpecialCells(xlLastCell).Column
                strContents = strContents & """" & LCase(ws.Cells(2, x)) & ""","
            Next 'x
            ' remove last comma
            strContents = Left(strContents, Len(strContents) - 1)

            ' add all text string values
            For y = 6 To ws.Cells.SpecialCells(xlLastCell).Row
            
                ' get key and value
                strKey = ws.Cells(y, 1)
                    
                If (strKey = "") Then
                    ' empty line skip
                ElseIf (Left$(strKey, 2) = "//") Then
                    ' comment line skip
                Else
                    ' key value
                    strContents = strContents & vbCrLf & FormatCsvValue(strKey)

                    ' get all text string values
                    For x = 2 To ws.Cells.SpecialCells(xlLastCell).Column
                        ' format .csv cell values
                        strValue = FormatCsvValue(ws.Cells(y, x))

                        ' add to output
                        strContents = strContents & "," & strValue
                    Next x
                End If

            Next y


            ' overwrite and save file
            If FileOrDirExists(strFilename) Then
                Kill strFilename
            End If
            Call SaveToFile(strFilename, strContents)

            MsgBox ("Translation .csv file created in folder " & OUTPUT_DIR)
        End If

    Next ws
 
End Sub

Sub GenerateLocalisationJson()
 
    ' write worksheet translations to Javascript JSON files
 
    Dim ws As Worksheet
    Dim strLanguageFolder As String
    Dim x As Integer
    Dim y As Integer
    Dim strFilename As String
    Dim strLangDesc As String
    Dim strLangCode As String
    Dim strCode As String
    Dim strKey As String
    Dim strValue As String
    Dim strContents As String
    Dim strContentsAll As String
 
    ' create JSON directory if not exist
    Const OUTPUT_DIR = "json\"
    Call FolderCreate(ActiveWorkbook.path & "\" & OUTPUT_DIR)
    
    strContentsAll = ""

    ' check all worksheets
    For Each ws In Worksheets
 
        ' only the active worksheet
        If ws.Name = ActiveSheet.Name Then

            ' get all field definition
            For x = 2 To ws.Cells.SpecialCells(xlLastCell).Column
            
                ' language code
                strLangDesc = ws.Cells(1, x)
                strLangCode = LCase(ws.Cells(2, x))

                ' determine filename
                strFilename = ActiveWorkbook.path & "\" & OUTPUT_DIR & strLangDesc & ".json" ' example "English.json"
                
                ' initialise content
                strContents = vbTab & """" & strLangCode & """: {"

                For y = 6 To ws.Cells.SpecialCells(xlLastCell).Row
                    
                    ' get key and value
                    strKey = ws.Cells(y, 1)
                    strValue = ws.Cells(y, x)

                    ' ignore empty lines
                    If (strKey <> "") Then
                        If (Left$(strKey, 2) = "//") Then
                            ' JSON does not support comment lines
                        Else
                            ' language key value
                            strContents = strContents & vbCrLf & vbTab & vbTab & """" & Replace(strKey, """", "\""") & """: """ & Replace(strValue, """", "\""") & ""","
                        End If
                    End If
                Next y

                ' remove last comma
                strContents = Left(strContents, Len(strContents) - 1)
                
                ' close current language, and add to total
                strContents = strContents & vbCrLf & vbTab & "}"
                strContentsAll = strContentsAll & vbCrLf & strContents & ","

                ' close
                strContents = "{" & vbCrLf & strContents & vbCrLf & "}"

                ' write to single languages file
                If FileOrDirExists(strFilename) Then
                    Kill strFilename
                End If
                Call SaveToFile(strFilename, strContents)

            Next x
            
            ' remove last comma
            strContentsAll = Left(strContentsAll, Len(strContentsAll) - 1)
 
        End If
        
        ' open and closing brackets
        strContentsAll = "{" & strContentsAll & vbCrLf & "}"
    
        ' write to all languages combined in one file
        strFilename = ActiveWorkbook.path & "\" & OUTPUT_DIR & "all_translations.json" ' example "English.json"
        If FileOrDirExists(strFilename) Then
            Kill strFilename
        End If
        Call SaveToFile(strFilename, strContentsAll)

        MsgBox ("Translation files created in folder " & OUTPUT_DIR)

    Next ws
 
End Sub

Sub GenerateLocalisationXcode()
 
    ' write worksheet translations to XCode Localizable.strings files
 
    Dim ws As Worksheet
    Dim strLanguageFolder As String
    Dim x As Integer
    Dim y As Integer
    Dim strFilename As String
    Dim strLangCode As String
    Dim strCode As String
    Dim strKey As String
    Dim strValue As String
    Dim strContents As String
 
    ' create XCode directory if not exist
    Const OUTPUT_DIR = "xcode\"
    Call FolderCreate(ActiveWorkbook.path & "\" & OUTPUT_DIR)

    ' check all worksheets
    For Each ws In Worksheets
 
        ' only the active worksheet
        If ws.Name = ActiveSheet.Name Then

            ' get all field definition
            For x = 2 To ws.Cells.SpecialCells(xlLastCell).Column
            
                ' create language directories if not exist
                strLangCode = LCase(ws.Cells(2, x))
                strLanguageFolder = "\" & OUTPUT_DIR & strLangCode & ".lproj" ' example "en.lproj"

                Call FolderCreate(ActiveWorkbook.path & strLanguageFolder)

                ' initialise
                strFilename = ActiveWorkbook.path & strLanguageFolder & "\Localizable.strings"
                
                ' initialise comment
                strContents = "/*" & Chr$(10)
                strContents = strContents & vbTab & "Localizable.Strings" & Chr$(10)
                strContents = strContents & vbTab & ws.Name & " (" & ws.Cells(1, x) & ")" & Chr$(10)
                strContents = strContents & vbTab & "Translation by " & ws.Cells(4, x) & Chr$(10) & Chr$(10)

                strContents = strContents & vbTab & "Generated: " & Format(Now(), "dd-mmm-yyyy hh:nn") & Chr$(10)
                strContents = strContents & "*/" & Chr$(10) & Chr$(10)

                For y = 6 To ws.Cells.SpecialCells(xlLastCell).Row
                    
                    ' get key and value
                    strKey = ws.Cells(y, 1)
                    strValue = ws.Cells(y, x)

                    If (strKey = "") Then
                        ' empty line
                        strContents = strContents & Chr$(10)  ' char 10 = Unix linefeed
                    ElseIf (Left$(strKey, 2) = "//") Then
                        ' comment line
                        strContents = strContents & strKey & Chr$(10)  ' char 10 = Unix linefeed
                    Else
                        ' language key value
                        strContents = strContents & """" & Replace(strKey, """", "\""") & """ = """ & Replace(strValue, """", "\""") & """;" & Chr$(10) ' char 10 = Unix linefeed
                    End If
                    
                Next y

                ' overwrite and save file
                If FileOrDirExists(strFilename) Then
                    Kill strFilename
                End If
                Call SaveToFile(strFilename, strContents)

            Next x
 
        End If

        MsgBox ("Translation files created in folder " & OUTPUT_DIR)

    Next ws
 
End Sub

Sub GenerateLocalisationEclipse()
 
    ' write worksheet translations to Android xml files
 
    Dim ws As Worksheet
    Dim strLanguageFolder As String
    Dim x As Integer
    Dim y As Integer
    Dim strFilename As String
    Dim strLangCode As String
    Dim strCode As String
    Dim strKey As String
    Dim strXmlKey As String
    Dim strValue As String
    Dim strContents As String
 
    ' create eclipse directory if not exist
    Const OUTPUT_DIR = "eclipse\"
    Call FolderCreate(ActiveWorkbook.path & "\" & OUTPUT_DIR)
 
    ' check all worksheets
    For Each ws In Worksheets
 
        ' only the active worksheet
        If ws.Name = ActiveSheet.Name Then

            ' get all field definition
            For x = 2 To ws.Cells.SpecialCells(xlLastCell).Column
            
                ' create language directories if not exist
                strLangCode = LCase(ws.Cells(2, x))
                If (strLangCode = "en") Then
                    strLanguageFolder = "\" & OUTPUT_DIR & "values" ' english, default
                Else
                    strLanguageFolder = "\" & OUTPUT_DIR & "values-" & strLangCode ' other languages
                End If

                Call FolderCreate(ActiveWorkbook.path & strLanguageFolder)

                ' initialise
                strFilename = ActiveWorkbook.path & strLanguageFolder & "\strings.xml"
                strContents = "<?xml version=""1.0"" encoding=""utf-8""?>" & Chr$(10) & "<resources>" & Chr$(10)
                
                ' initialise comment
                strContents = strContents & vbTab & "<!--" & Chr$(10)
                strContents = strContents & vbTab & "Localizable.Strings" & Chr$(10)
                strContents = strContents & vbTab & ws.Name & " (" & ws.Cells(1, x) & ")" & Chr$(10)
                strContents = strContents & vbTab & "Translation by " & ws.Cells(4, x) & Chr$(10) & Chr$(10)

                strContents = strContents & vbTab & "Generated: " & Format(Now(), "dd-mmm-yyyy hh:nn") & Chr$(10)
                strContents = strContents & vbTab & "-->" & Chr$(10) & Chr$(10)

                For y = 6 To ws.Cells.SpecialCells(xlLastCell).Row
                    
                    ' get key and value
                    strKey = ws.Cells(y, 1)
                    strValue = ws.Cells(y, x)

                    If (strKey = "") Then
                        ' empty line
                        strContents = strContents & Chr$(10)  ' char 10 = Unix linefeed
                    ElseIf (Left$(strKey, 2) = "//") Then
                        ' comment line
                        strContents = strContents & vbTab & "<!--" & Mid$(strKey, 3) & " -->" & Chr$(10)  ' char 10 = Unix linefeed
                    Else
                        ' language key value
                        strContents = strContents & vbTab & "<string name=""" & ReplaceXmlKey(LCase(strKey)) & """>" & ReplaceXmlValue(strValue) & "</string>" & Chr$(10)  ' char 10 = Unix linefeed
                    End If
                Next y

                ' close XML
                strContents = strContents & "</resources>"

                ' overwrite and save file
                If FileOrDirExists(strFilename) Then
                    Kill strFilename
                End If
                Call SaveToFile(strFilename, strContents)

            Next x
 
        End If

        MsgBox ("Translation files created in folder " & OUTPUT_DIR)

    Next ws
 
End Sub

Sub GenerateLocalisationVisualStudio()
 
    ' write worksheet translations to Visual Studio .resx files
 
    Dim ws As Worksheet
    Dim x As Integer
    Dim y As Integer
    Dim strFilename As String
    Dim strLangCode As String
    Dim strLangExt As String
    Dim strCode As String
    Dim strKey As String
    Dim strXmlKey As String
    Dim strValue As String
    Dim strContents As String
    Dim strDesignerFileName As String
    Dim strDesignerCode As String
 
    ' create eclipse directory if not exist
    Const OUTPUT_DIR = "visualstudio\"
    Const OUTPUT_FILENAME = "Strings"
    Call FolderCreate(ActiveWorkbook.path & "\" & OUTPUT_DIR)
    
    Const VS_XSD1 = _
    "  <xsd:schema id=""root"" xmlns="""" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:msdata=""urn:schemas-microsoft-com:xml-msdata"">" & vbCrLf & _
    "    <xsd:import namespace=""http://www.w3.org/XML/1998/namespace"" />" & vbCrLf & _
    "    <xsd:element name=""root"" msdata:IsDataSet=""true"">" & vbCrLf & _
    "      <xsd:complexType>" & vbCrLf & _
    "        <xsd:choice maxOccurs=""unbounded"">" & vbCrLf & _
    "          <xsd:element name=""metadata"">" & vbCrLf & _
    "            <xsd:complexType>" & vbCrLf & _
    "              <xsd:sequence>" & vbCrLf & _
    "                <xsd:element name=""value"" type=""xsd:string"" minOccurs=""0"" />" & vbCrLf & _
    "              </xsd:sequence>" & vbCrLf & _
    "              <xsd:attribute name=""name"" use=""required"" type=""xsd:string"" />" & vbCrLf & _
    "              <xsd:attribute name=""type"" type=""xsd:string"" />" & vbCrLf & _
    "              <xsd:attribute name=""mimetype"" type=""xsd:string"" />" & vbCrLf & _
    "              <xsd:attribute ref=""xml:space"" />" & vbCrLf & _
    "            </xsd:complexType>" & vbCrLf & _
    "          </xsd:element>" & vbCrLf & _
    "          <xsd:element name=""assembly"">" & vbCrLf & _
    "            <xsd:complexType>" & vbCrLf & _
    "              <xsd:attribute name=""alias"" type=""xsd:string"" />" & vbCrLf & _
    "              <xsd:attribute name=""name"" type=""xsd:string"" />" & vbCrLf & _
    "            </xsd:complexType>" & vbCrLf & _
    "          </xsd:element>" & vbCrLf
    
    Const VS_XSD2 = _
    "          <xsd:element name=""data"">" & vbCrLf & _
    "            <xsd:complexType>" & vbCrLf & _
    "              <xsd:sequence>" & vbCrLf & _
    "                <xsd:element name=""value"" type=""xsd:string"" minOccurs=""0"" msdata:Ordinal=""1"" />" & vbCrLf & _
    "                <xsd:element name=""comment"" type=""xsd:string"" minOccurs=""0"" msdata:Ordinal=""2"" />" & vbCrLf & _
    "              </xsd:sequence>" & vbCrLf & _
    "              <xsd:attribute name=""name"" type=""xsd:string"" use=""required"" msdata:Ordinal=""1"" />" & vbCrLf & _
    "              <xsd:attribute name=""type"" type=""xsd:string"" msdata:Ordinal=""3"" />" & vbCrLf & _
    "              <xsd:attribute name=""mimetype"" type=""xsd:string"" msdata:Ordinal=""4"" />" & vbCrLf & _
    "              <xsd:attribute ref=""xml:space"" />" & vbCrLf & _
    "            </xsd:complexType>" & vbCrLf & _
    "          </xsd:element>" & vbCrLf & _
    "          <xsd:element name=""resheader"">" & vbCrLf & _
    "            <xsd:complexType>" & vbCrLf & _
    "              <xsd:sequence>" & vbCrLf & _
    "                <xsd:element name=""value"" type=""xsd:string"" minOccurs=""0"" msdata:Ordinal=""1"" />" & vbCrLf & _
    "              </xsd:sequence>" & vbCrLf & _
    "              <xsd:attribute name=""name"" type=""xsd:string"" use=""required"" />" & vbCrLf & _
    "            </xsd:complexType>" & vbCrLf & _
    "          </xsd:element>" & vbCrLf & _
    "        </xsd:choice>" & vbCrLf & _
    "      </xsd:complexType>" & vbCrLf & _
    "    </xsd:element>" & vbCrLf & _
    "  </xsd:schema>" & vbCrLf
    
    Const VS_XML_HEADER = _
    "  <resheader name=""resmimetype"">" & vbCrLf & _
    "    <value>text/microsoft-resx</value>" & vbCrLf & _
    "  </resheader>" & vbCrLf & _
    "  <resheader name=""version"">" & vbCrLf & _
    "    <value>2.0</value>" & vbCrLf & _
    "  </resheader>" & vbCrLf & _
    "  <resheader name=""reader"">" & vbCrLf & _
    "    <value>System.Resources.ResXResourceReader, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</value>" & vbCrLf & _
    "  </resheader>" & vbCrLf & _
    "  <resheader name=""writer"">" & vbCrLf & _
    "    <value>System.Resources.ResXResourceWriter, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</value>" & vbCrLf & _
    "  </resheader>" & vbCrLf
    
    Const DESIGN_CODE1 = _
    "//------------------------------------------------------------------------------" & vbCrLf & _
    "// <auto-generated>" & vbCrLf & _
    "//     This code was generated by a tool." & vbCrLf & _
    "//     Runtime Version:4.0.30319.42000" & vbCrLf & _
    "//" & vbCrLf & _
    "//     Changes to this file may cause incorrect behavior and will be lost if" & vbCrLf & _
    "//     the code is regenerated." & vbCrLf & _
    "// </auto-generated>" & vbCrLf & _
    "//------------------------------------------------------------------------------" & vbCrLf & _
    "" & vbCrLf & _
    "namespace MyProject1 {" & vbCrLf & _
    "    using System;" & vbCrLf & _
    "" & vbCrLf & _
    "" & vbCrLf & _
    "    /// <summary>" & vbCrLf & _
    "    ///   A strongly-typed resource class, for looking up localized strings, etc." & vbCrLf & _
    "    /// </summary>" & vbCrLf & _
    "    // This class was auto-generated by the StronglyTypedResourceBuilder" & vbCrLf & _
    "    // class via a tool like ResGen or Visual Studio." & vbCrLf & _
    "    // To add or remove a member, edit your .ResX file then rerun ResGen" & vbCrLf & _
    "    // with the /str option, or rebuild your VS project." & vbCrLf & _
    "    [global::System.CodeDom.Compiler.GeneratedCodeAttribute(""System.Resources.Tools.StronglyTypedResourceBuilder"", ""17.0.0.0"")]" & vbCrLf & _
    "    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]" & vbCrLf
    
    Const DESIGN_CODE2 = _
    "    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]" & vbCrLf & _
    "    internal class Strings {" & vbCrLf & _
    "" & vbCrLf & _
    "        private static global::System.Resources.ResourceManager resourceMan;" & vbCrLf & _
    "" & vbCrLf & _
    "        private static global::System.Globalization.CultureInfo resourceCulture;" & vbCrLf & _
    "" & vbCrLf & _
    "        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(""Microsoft.Performance"", ""CA1811:AvoidUncalledPrivateCode"")]" & vbCrLf & _
    "        internal Strings() {" & vbCrLf & _
    "        }" & vbCrLf & _
    "" & vbCrLf & _
    "        /// <summary>" & vbCrLf & _
    "        ///   Returns the cached ResourceManager instance used by this class." & vbCrLf & _
    "        /// </summary>" & vbCrLf & _
    "        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]" & vbCrLf & _
    "        internal static global::System.Resources.ResourceManager ResourceManager {" & vbCrLf & _
    "            get {" & vbCrLf & _
    "                if (object.ReferenceEquals(resourceMan, null)) {" & vbCrLf & _
    "                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager(""MyProject1.Strings"", typeof(Strings).Assembly);" & vbCrLf & _
    "                    resourceMan = temp;" & vbCrLf & _
    "                }" & vbCrLf & _
    "                return resourceMan;" & vbCrLf & _
    "            }" & vbCrLf
    
    Const DESIGN_CODE3 = _
    "        }" & vbCrLf & _
    "" & vbCrLf & _
    "        /// <summary>" & vbCrLf & _
    "        ///   Overrides the current thread's CurrentUICulture property for all" & vbCrLf & _
    "        ///   resource lookups using this strongly typed resource class." & vbCrLf & _
    "        /// </summary>" & vbCrLf & _
    "        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]" & vbCrLf & _
    "        internal static global::System.Globalization.CultureInfo Culture {" & vbCrLf & _
    "            get {" & vbCrLf & _
    "                return resourceCulture;" & vbCrLf & _
    "            }" & vbCrLf & _
    "            set {" & vbCrLf & _
    "                resourceCulture = value;" & vbCrLf & _
    "            }" & vbCrLf & _
    "        }" & vbCrLf
 
    ' check all worksheets
    For Each ws In Worksheets
 
        ' only the active worksheet
        If ws.Name = ActiveSheet.Name Then

            ' ---- Generate resource ResXC# for each language ----
            ' get all field definition
            For x = 2 To ws.Cells.SpecialCells(xlLastCell).Column
            
                ' use worksheet name and language code as filename
                strLangCode = LCase(ws.Cells(2, x))
                ' First language is Default language, so no language extension, i.e. "Strings.resx" instead of "Strings.en.resx"
                strLangExt = ""
                If (x > 2) Then strLangExt = "." & strLangCode
                strFilename = ActiveWorkbook.path & "\" & OUTPUT_DIR & OUTPUT_FILENAME & strLangExt & ".resx"
                
                ' initialise contents
                strContents = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf & "<root>" & vbCrLf
                
                ' initialise comment
                strContents = strContents & "  <!--" & vbCrLf
                strContents = strContents & "    Microsoft ResX Schema" & vbCrLf & vbCrLf
                strContents = strContents & "    Version 2.0" & vbCrLf & vbCrLf
                strContents = strContents & "    " & ws.Name & " (" & ws.Cells(1, x) & ")" & vbCrLf
                strContents = strContents & "    Translation by " & ws.Cells(4, x) & vbCrLf & vbCrLf

                strContents = strContents & "    Generated: " & Format(Now(), "dd-mmm-yyyy hh:nn") & vbCrLf
                strContents = strContents & "  -->" & vbCrLf
                
                strContents = strContents & VS_XSD1 & VS_XSD2 & VS_XML_HEADER

                For y = 6 To ws.Cells.SpecialCells(xlLastCell).Row
                    
                    ' get key and value
                    strKey = ws.Cells(y, 1)
                    strValue = ws.Cells(y, x)

                    If (strKey = "") Then
                        ' empty line skip
                    ElseIf (Left$(strKey, 2) = "//") Then
                        ' comment line
                        strContents = strContents & "  <!--" & Mid$(strKey, 3) & " -->" & vbCrLf
                    Else
                        ' language key value
                        strContents = strContents & "  <data name=""" & ReplaceXmlKey(strKey) & """>" & vbCrLf
                        strContents = strContents & "    <value>" & ReplaceXmlValue(strValue) & "</value>" & vbCrLf
                        strContents = strContents & "  </data>" & vbCrLf
                    End If
                Next y

                ' close XML
                strContents = strContents & "</root>"

                ' overwrite and save file
                If FileOrDirExists(strFilename) Then
                    Kill strFilename
                End If
                Call SaveToFile(strFilename, strContents)

            Next x
            
            ' ---- Generate C# Designer code file ----
            ' use worksheet name and language code as filename
            strFilename = ActiveWorkbook.path & "\" & OUTPUT_DIR & OUTPUT_FILENAME & ".Designer.resx"
            
            ' initialise contents
            strContents = DESIGN_CODE1 & DESIGN_CODE2 & DESIGN_CODE3
            
            For y = 6 To ws.Cells.SpecialCells(xlLastCell).Row
                
                ' get key and value
                strKey = ws.Cells(y, 1)
                strValue = ws.Cells(y, 2)

                If (strKey = "") Then
                    ' empty line skip
                ElseIf (Left$(strKey, 2) = "//") Then
                    ' comment line
                    strContents = strContents & "        // ---- " & Mid$(strKey, 3) & " ----" & vbCrLf
                Else
                    ' language key value
                    strContents = strContents & "        /// <summary>" & vbCrLf
                    strContents = strContents & "        ///   Looks up a localized string: " & strValue & vbCrLf
                    strContents = strContents & "        /// </summary>" & vbCrLf
                    strContents = strContents & "        internal static string " & strKey & " {" & vbCrLf
                    strContents = strContents & "            get {" & vbCrLf
                    strContents = strContents & "                return ResourceManager.GetString(""" & strKey & """, resourceCulture);" & vbCrLf
                    strContents = strContents & "            }" & vbCrLf
                    strContents = strContents & "        }" & vbCrLf
                End If
            Next y

            ' close C# code
            strContents = strContents & vbCrLf & "    }" & vbCrLf & "}"

            ' overwrite and save file
            If FileOrDirExists(strFilename) Then
                Kill strFilename
            End If
            Call SaveToFile(strFilename, strContents)
 
        End If

        MsgBox ("Translation files created in folder " & OUTPUT_DIR)

    Next ws
 
End Sub

Sub UniqueCharactersFromSelection()

    ' collect all unique characters from the selected cells, for when creating bitmap font

    Dim strFilename As String
    Dim cell As Range
    Dim txt As String
    Dim i As Long, j As Long, k As Long
    Dim ch As String
    Dim code As Long
    Dim codes() As Long
    Dim codeCount As Long
    Dim exists As Boolean
    Dim temp As Long
    Dim result As String
    
    strFilename = ActiveWorkbook.path & "\Unique_characters_output.txt"
    result = "; Manage-translations, all unique characters in text" & vbCrLf & "; useful when creating bitmap fonts" & vbCrLf & "; Generated on: " & Format(Now(), "dd-mmm-yyyy hh:nn")

    ' Combine all text from selected cells
    For Each cell In Selection
        If Not IsEmpty(cell.Value) Then
            txt = txt & cell.Value
        End If
    Next cell

    ' Do it twice
    For k = 1 To 2
    
        ' Add comment line
        result = result & vbCrLf & vbCrLf
        If (k = 1) Then result = result & "; unique characters in selection" & vbCrLf
        If (k = 2) Then result = result & "; unique characters + minial complete a-z A-Z 0-9" & vbCrLf
        
        ' Add minial complete a-z A-Z 0-9 characters
        If (k = 2) Then txt = txt & "1234567890!@#$%^&*()-=+ abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ"

        ' Loop through each character
        For i = 1 To Len(txt)
            code = Asc(Mid(txt, i, 1))
            exists = False
            
            ' Check if this code is already in array
            For j = 1 To codeCount
                If codes(j) = code Then
                    exists = True
                    Exit For
                End If
            Next j
            
            ' Add new unique code
            If Not exists Then
                codeCount = codeCount + 1
                ReDim Preserve codes(1 To codeCount)
                codes(codeCount) = code
            End If
        Next i
    
        ' --- Sort array (simple Bubble Sort for clarity) ---
        For i = 1 To codeCount - 1
            For j = i + 1 To codeCount
                If codes(i) > codes(j) Then
                    temp = codes(i)
                    codes(i) = codes(j)
                    codes(j) = temp
                End If
            Next j
        Next i
    
        ' Add all characters to output string
        For i = 1 To codeCount
            result = result & Chr(codes(i))
        Next i
    Next k
    
    ' overwrite and save file
    If FileOrDirExists(strFilename) Then
        Kill strFilename
    End If
    Call SaveToFile(strFilename, result)

    ' --- Show result ---
    MsgBox "Unique sorted characters:" & vbCrLf & result

End Sub


' --------------------------------------
' Helper subs and functions
' --------------------------------------
Function ReplaceXmlKey(sXmlKey As String) As String

    ' prepare xml-safe-key
    sXmlKey = Replace(sXmlKey, "-", "_")
    sXmlKey = Replace(sXmlKey, "/", "_")
    sXmlKey = Replace(sXmlKey, ".", " ")
    sXmlKey = Replace(sXmlKey, "<", " ")
    sXmlKey = Replace(sXmlKey, ">", " ")
    sXmlKey = Replace(sXmlKey, "?", " ")
    sXmlKey = Replace(sXmlKey, "&", " ")
    sXmlKey = Replace(sXmlKey, "'", " ")
    sXmlKey = Replace(sXmlKey, """", " ")
    sXmlKey = Replace(sXmlKey, "  ", " ")
    sXmlKey = Replace(sXmlKey, "  ", " ")
    sXmlKey = Replace(sXmlKey, "  ", " ")
    sXmlKey = Replace(sXmlKey, " ", "_")

    ReplaceXmlKey = sXmlKey
End Function

Function ReplaceXmlValue(sXmlValue As String) As String

    ' prepare xml-safe-value
    sXmlValue = Replace(sXmlValue, "&", "&amp;")
    sXmlValue = Replace(sXmlValue, "<", "&lt;")
    sXmlValue = Replace(sXmlValue, ">", "&gt;")
    sXmlValue = Replace(sXmlValue, """", "&quot;")
    sXmlValue = Replace(sXmlValue, "'", "&apos;")

    ReplaceXmlValue = sXmlValue
End Function

Function FormatCsvValue(sCsvValue As String) As String

    ' csv format
    sCsvValue = Replace(sCsvValue, """", """""") 'escape double quote
    sCsvValue = """" & sCsvValue & """"

    FormatCsvValue = sCsvValue
End Function
