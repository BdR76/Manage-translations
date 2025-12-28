REM  *****  BASIC  *****

' Multilanguage.ods by Bas de Reuver (bdr1976@gmail.com) 2012-2025 Free to use
'
' This VBA script writes content of Excel sheet to translation files
' which can be used in XCode, Eclipse and Visual Studio.
' Left most column contains the translations keys,
' each column contains the translation values.
'
' Manage translations words and phrases in spreadsheet
' and use this macro to export to different source files
' * GenerateLocalisationCSV - CSV file for example Godot game engine
' * GenerateLocalisationJson - JSON files for javascript
' * GenerateLocalisationXcode - XCode Localizable.strings files for iPhone
' * GenerateLocalisationEclipse - XML strings.xml files for Android
' * GenerateLocalisationVisualStudio - .resx xml files for Visual Studio

' global String LINE_BREAK = "vbCrLf" ' Chr(13) & Chr(10) gives syntax error

' --------------------------------------
' Export CSV files
' --------------------------------------
Sub GenerateLocalisationCSV

	Dim OUTPUT_DIR
	Dim LINE_BREAK
	Dim HEADER_ROWS
	LINE_BREAK = Chr(13) & Chr(10) ' doesn't gives syntax error when used here
	HEADER_ROWS = 5

	Dim oSheet as Object
	Dim oCursor
	Dim col_start
	Dim row_start
	Dim col_end
	Dim row_end
	Dim strtmp

	Dim sContent as String
	Dim col as Integer
	Dim row as Integer
	Dim str_key as String
	Dim str_val as String

	' create csv directory if not exist
	OUTPUT_DIR = "csv\"

	oSheet = ThisComponent.CurrentController.ActiveSheet

	' get cells range, all of sheet contents
	oCursor = oSheet.createCursor()
	oCursor.gotoStartOfUsedArea(False)
	oCursor.gotoEndOfUsedArea(True)

	' determine how many rows and columns
	col_start = oCursor.RangeAddress.StartColumn
	row_start = oCursor.RangeAddress.StartRow

	col_end = oCursor.RangeAddress.EndColumn
	row_end = oCursor.RangeAddress.EndRow

	' file content
	sContent = ""

	' output filename (directories will be automatically created if not exist)
	sFilename = FilePath() & OUTPUT_DIR & "multilanguage.csv"

	' build column header
	sContent = FormatCsvValue("Keys")
	For col = 1 to col_end
		' next language, get header language code
		LanguageCode = LCase(oSheet.getCellByPosition(col, 1).String)
		sContent = sContent & ",""" & LanguageCode & """"
	Next col

	' add all text string values
	For row = HEADER_ROWS to row_end

		' get key and value
		str_key = oSheet.getCellByPosition(0, row).String

		If (str_key = "") Then
			' empty line skip
		ElseIf (Left(str_key, 2) = "//") Then
			' comment line skip
		Else
			' key value
			sContent = sContent & LINE_BREAK & FormatCsvValue(str_key)

			' get all text string values
			For col = 1 to col_end
				str_val = FormatCsvValue(oSheet.getCellByPosition(col, row).String)
				sContent = sContent & "," & str_val
			Next col
		End If
	Next row

	' write to single languages file
	WriteToTextFile(sFilename, sContent)

	MsgBox ("Translation files for " & col_end & " language created in folder " & OUTPUT_DIR)

End Sub

' --------------------------------------
' Export JSON localization files
' --------------------------------------
Sub GenerateLocalisationJson

	Dim OUTPUT_DIR
	Dim LINE_BREAK
	Dim HEADER_ROWS
	LINE_BREAK = Chr(13) & Chr(10) ' doesn't gives syntax error when used here
	HEADER_ROWS = 5

	OUTPUT_DIR = "json\"

	Dim oSheet as Object
	Dim oCursor
	Dim col_start
	Dim row_start
	Dim col_end
	Dim row_end
	Dim strtmp

	Dim sContent as String
	Dim sContentAll as String
	Dim col as Integer
	Dim row as Integer
	Dim str_key as String
	Dim str_val as String

	oSheet = ThisComponent.CurrentController.ActiveSheet

	' get cells range, all of sheet contents
	oCursor = oSheet.createCursor()
	oCursor.gotoStartOfUsedArea(False)
	oCursor.gotoEndOfUsedArea(True)

	' determine how many rows and columns
	col_start = oCursor.RangeAddress.StartColumn
	row_start = oCursor.RangeAddress.StartRow

	col_end = oCursor.RangeAddress.EndColumn
	row_end = oCursor.RangeAddress.EndRow

	' file content
	sContent = ""
	sContentAll = ""

	For col = 1 to col_end
		' next language, get header info
		LanguageEn = oSheet.getCellByPosition(col, 0).String
		LanguageCode = LCase(oSheet.getCellByPosition(col, 1).String)
		LanguageDisplay = oSheet.getCellByPosition(col, 2).String
		LanguageTranslator = oSheet.getCellByPosition(col, 3).String

		' output filename (directories will be automatically created if not exist)
		sFilename = FilePath() & OUTPUT_DIR & LanguageEn & ".json"

		' start file content
		sContent = "	""" & LanguageCode & """: {"

		For row = HEADER_ROWS to row_end

			str_key = oSheet.getCellByPosition(0, row).String
			str_val = oSheet.getCellByPosition(col, row).String
			If (str_key <> "") Then
				If (Left(str_key, 2) = "//") Then
					' JSON format does not support comment lines
					' sContent = sContent & str_key & " COMMENT" & LINE_BREAK
				Else
					' phrase value
					sContent = sContent & LINE_BREAK & "		""" & Replace(str_key, """", "\""") & """: """ & Replace(str_val, """", "\""") & ""","
				End If
			End If

		Next row

		' remove last comma
		sContent = Left(sContent, Len(sContent) - 1)

		' close current language, and add to total
		sContent = sContent & LINE_BREAK & "	}"
		sContentAll = sContentAll & LINE_BREAK & sContent & ","

		' open and closing brackets
		sContent = "{" & LINE_BREAK & sContent & LINE_BREAK & "}"

		' write to single languages file
		WriteToTextFile(sFilename, sContent)
	Next col

	' write to all languages combined in one file
	sFilename = FilePath() & OUTPUT_DIR & "all_translations.json"

	' remove last comma
	sContentAll = Left(sContentAll, Len(sContentAll) - 1)

	' open and closing brackets
	sContentAll = "{" & sContentAll & LINE_BREAK & "}"

	WriteToTextFile(sFilename, sContentAll)

	MsgBox ("Translation files for " & col_end & " language created in folder " & OUTPUT_DIR)

End Sub

' --------------------------------------
' Export localization strings Xcode (iPhone)
' --------------------------------------
Sub GenerateLocalisationXcode

	Dim OUTPUT_DIR
	Dim LINE_BREAK
	Dim HEADER_ROWS
	LINE_BREAK = Chr(10) ' unix line breaks for Xcode
	HEADER_ROWS = 5

	OUTPUT_DIR = "xcode\"

	Dim oSheet as Object
	Dim oCursor
	Dim col_start
	Dim row_start
	Dim col_end
	Dim row_end
	Dim strtmp

	Dim sContent as String
	Dim col as Integer
	Dim row as Integer
	Dim str_key as String
	Dim str_val as String

	oSheet = ThisComponent.CurrentController.ActiveSheet

	' get cells range, all of sheet contents
	oCursor = oSheet.createCursor()
	oCursor.gotoStartOfUsedArea(False)
	oCursor.gotoEndOfUsedArea(True)

	' determine how many rows and columns
	col_start = oCursor.RangeAddress.StartColumn
	row_start = oCursor.RangeAddress.StartRow

	col_end = oCursor.RangeAddress.EndColumn
	row_end = oCursor.RangeAddress.EndRow

	' file content
	sContent = ""
	sContentAll = ""

	For col = 1 to col_end
		' next language, get header info
		LanguageEn = oSheet.getCellByPosition(col, 0).String
		LanguageCode = LCase(oSheet.getCellByPosition(col, 1).String)
		LanguageDisplay = oSheet.getCellByPosition(col, 2).String
		LanguageTranslator = oSheet.getCellByPosition(col, 3).String

		' output filename (directories will be automatically created if not exist)
		sFilename = FilePath() & OUTPUT_DIR & LanguageCode & ".lproj\Localizable.strings" ' example folder "en.lproj" file "Localizable.strings"

        ' initialise comment
		sContent = "/*" & LINE_BREAK
		sContent = sContent & "	Localizable.Strings" & LINE_BREAK
		sContent = sContent & "	" & oSheet.Name & " (" & LanguageEn & ")" & LINE_BREAK
		sContent = sContent & "	Translation by " & LanguageTranslator & LINE_BREAK & LINE_BREAK

		sContent = sContent & "	Generated: " & Format(Now(), "dd-mmm-yyyy hh:mm") & LINE_BREAK
		sContent = sContent & "*/" & LINE_BREAK & LINE_BREAK

		For row = HEADER_ROWS to row_end

			str_key = oSheet.getCellByPosition(0, row).String
			str_val = oSheet.getCellByPosition(col, row).String

			If (str_key = "") Then
				' empty line
				sContent = sContent & LINE_BREAK ' char 10 = Unix linefeed
			ElseIf (Left(str_key, 2) = "//") Then
				' comment line
				sContent = sContent & str_key & LINE_BREAK  ' char 10 = Unix linefeed
			Else
				' language key value
				sContent = sContent & """" & Replace(str_key, """", "\""") & """ = """ & Replace(str_val, """", "\""") & """;" & LINE_BREAK ' char 10 = Unix linefeed
			End If

		Next row

		WriteToTextFile(sFilename, sContent)
	Next col

	MsgBox ("Translation files for " & col_end & " language created in folder " & OUTPUT_DIR)

End Sub

' --------------------------------------
' Export XML localization files Eclipse (Android)
' --------------------------------------
Sub GenerateLocalisationEclipse

	Dim OUTPUT_DIR
	Dim LINE_BREAK
	Dim HEADER_ROWS
	LINE_BREAK = Chr(10) ' unix line breaks for Eclipse
	HEADER_ROWS = 5

	OUTPUT_DIR = "eclipse\"

	Dim oSheet as Object
	Dim oCursor
	Dim col_start
	Dim row_start
	Dim col_end
	Dim row_end
	Dim strtmp

	Dim sContent as String
	Dim col as Integer
	Dim row as Integer
	Dim str_key as String
	Dim str_val as String

	oSheet = ThisComponent.CurrentController.ActiveSheet

	' get cells range, all of sheet contents
	oCursor = oSheet.createCursor()
	oCursor.gotoStartOfUsedArea(False)
	oCursor.gotoEndOfUsedArea(True)

	' determine how many rows and columns
	col_start = oCursor.RangeAddress.StartColumn
	row_start = oCursor.RangeAddress.StartRow

	col_end = oCursor.RangeAddress.EndColumn
	row_end = oCursor.RangeAddress.EndRow

	' file content
	sContent = ""
	sContentAll = ""

	For col = 1 to col_end
		' next language, get header info
		LanguageEn = oSheet.getCellByPosition(col, 0).String
		LanguageCode = LCase(oSheet.getCellByPosition(col, 1).String)
		LanguageDisplay = oSheet.getCellByPosition(col, 2).String
		LanguageTranslator = oSheet.getCellByPosition(col, 3).String

		' create language directories if not exist
		If (LanguageCode = "en") Then
			sFilename = FilePath() & OUTPUT_DIR & "values" ' english, default
		Else
			sFilename = FilePath() & OUTPUT_DIR & "values-" & LanguageCode ' other languages
		End If

		' initialise
		sFilename = sFilename & "\strings.xml"

		' start file content
		sContent = "<?xml version=""1.0"" encoding=""utf-8""?>" & LINE_BREAK & "<resources>" & LINE_BREAK

        ' initialise comment
        sContent = sContent & "	<!--" & LINE_BREAK
        sContent = sContent & "	Localizable.Strings" & LINE_BREAK
        sContent = sContent & "	" & oSheet.Name & " (" & LanguageEn & ")" & LINE_BREAK
        sContent = sContent & "	Translation by " & LanguageTranslator & LINE_BREAK & LINE_BREAK

        sContent = sContent & "	Generated: " & Format(Now(), "dd-mmm-yyyy hh:mm") & LINE_BREAK
        sContent = sContent & "	-->" & LINE_BREAK & LINE_BREAK

		For row = HEADER_ROWS to row_end

			str_key = oSheet.getCellByPosition(0, row).String
			str_val = oSheet.getCellByPosition(col, row).String
			If (str_key = "") Then
				' empty line
				sContent = sContent & LINE_BREAK
			ElseIf (Left(str_key, 2) = "//") Then
				' comment lines
				sContent = sContent & "	<!-- " & Trim(Mid(str_key, 3)) & " -->" & LINE_BREAK
			Else
				' language key value
				sContent = sContent & "	<string name=""" & ReplaceXmlKey(LCase(str_key)) & """>" & ReplaceXmlValue(str_val) & "</string>" &LINE_BREAK
			End If

		Next row

		' close current language, and add to total
		sContent = sContent & "</resources>"

		WriteToTextFile(sFilename, sContent)
	Next col

	MsgBox ("Translation files for " & col_end & " language created in folder " & OUTPUT_DIR)

End Sub

' --------------------------------------
' Export .resx xml files Visual Studio
' --------------------------------------
Sub GenerateLocalisationVisualStudio

    Dim oDoc As Object
    Dim oSheet As Object
    Dim oSheets As Object

    Dim x As Long, y As Long
    Dim lastCol As Long, lastRow As Long

    Dim strFilename As String
    Dim strLangCode As String
    Dim strLangExt As String
    Dim strKey As String
    Dim strValue As String
    Dim strContents As String

	Dim LINE_BREAK
    LINE_BREAK = Chr(13) & Chr(10) ' windows line breaks

    Const OUTPUT_DIR = "visualstudio/"
    Const OUTPUT_FILENAME = "Strings"

    Dim VS_XSD1 As String
    Dim VS_XSD2 As String
	Dim VS_XSD3 As String
    Dim VS_XML_HEADER As String
	Dim DESIGN_CODE1 As String
	Dim DESIGN_CODE2 As String
	Dim DESIGN_CODE3 As String

    VS_XSD1 = _
    "  <xsd:schema id=""root"" xmlns="""" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:msdata=""urn:schemas-microsoft-com:xml-msdata"">" & LINE_BREAK & _
    "    <xsd:import namespace=""http://www.w3.org/XML/1998/namespace"" />" & LINE_BREAK & _
    "    <xsd:element name=""root"" msdata:IsDataSet=""true"">" & LINE_BREAK & _
    "      <xsd:complexType>" & LINE_BREAK & _
    "        <xsd:choice maxOccurs=""unbounded"">" & LINE_BREAK & _
    "          <xsd:element name=""metadata"">" & LINE_BREAK & _
    "            <xsd:complexType>" & LINE_BREAK & _
    "              <xsd:sequence>" & LINE_BREAK & _
    "                <xsd:element name=""value"" type=""xsd:string"" minOccurs=""0"" />" & LINE_BREAK & _
    "              </xsd:sequence>" & LINE_BREAK & _
    "              <xsd:attribute name=""name"" use=""required"" type=""xsd:string"" />" & LINE_BREAK & _
    "              <xsd:attribute name=""type"" type=""xsd:string"" />" & LINE_BREAK & _
    "              <xsd:attribute name=""mimetype"" type=""xsd:string"" />" & LINE_BREAK & _
    "              <xsd:attribute ref=""xml:space"" />" & LINE_BREAK & _
    "            </xsd:complexType>" & LINE_BREAK & _
    "          </xsd:element>" & LINE_BREAK & _
    "          <xsd:element name=""assembly"">" & LINE_BREAK & _
    "            <xsd:complexType>" & LINE_BREAK & _
    "              <xsd:attribute name=""alias"" type=""xsd:string"" />" & LINE_BREAK & _
    "              <xsd:attribute name=""name"" type=""xsd:string"" />" & LINE_BREAK & _
    "            </xsd:complexType>" & LINE_BREAK & _
    "          </xsd:element>" & LINE_BREAK

    VS_XSD2 = _
    "          <xsd:element name=""data"">" & LINE_BREAK & _
    "            <xsd:complexType>" & LINE_BREAK & _
    "              <xsd:sequence>" & LINE_BREAK & _
    "                <xsd:element name=""value"" type=""xsd:string"" minOccurs=""0"" msdata:Ordinal=""1"" />" & LINE_BREAK & _
    "                <xsd:element name=""comment"" type=""xsd:string"" minOccurs=""0"" msdata:Ordinal=""2"" />" & LINE_BREAK & _
    "              </xsd:sequence>" & LINE_BREAK & _
    "              <xsd:attribute name=""name"" type=""xsd:string"" use=""required"" msdata:Ordinal=""1"" />" & LINE_BREAK & _
    "              <xsd:attribute name=""type"" type=""xsd:string"" msdata:Ordinal=""3"" />" & LINE_BREAK & _
    "              <xsd:attribute name=""mimetype"" type=""xsd:string"" msdata:Ordinal=""4"" />" & LINE_BREAK & _
    "              <xsd:attribute ref=""xml:space"" />" & LINE_BREAK & _
    "            </xsd:complexType>" & LINE_BREAK & _
    "          </xsd:element>" & LINE_BREAK & _
    "          <xsd:element name=""resheader"">" & LINE_BREAK & _
    "            <xsd:complexType>" & LINE_BREAK & _
    "              <xsd:sequence>" & LINE_BREAK & _
    "                <xsd:element name=""value"" type=""xsd:string"" minOccurs=""0"" msdata:Ordinal=""1"" />" & LINE_BREAK & _
    "              </xsd:sequence>" & LINE_BREAK & _
    "              <xsd:attribute name=""name"" type=""xsd:string"" use=""required"" />" & LINE_BREAK & _
    "            </xsd:complexType>" & LINE_BREAK & _
    "          </xsd:element>" & LINE_BREAK & _
    "        </xsd:choice>" & LINE_BREAK & _
    "      </xsd:complexType>" & LINE_BREAK & _
    "    </xsd:element>" & LINE_BREAK & _
    "  </xsd:schema>" & LINE_BREAK

    VS_XML_HEADER = _
    "  <resheader name=""resmimetype"">" & LINE_BREAK & _
    "    <value>text/microsoft-resx</value>" & LINE_BREAK & _
    "  </resheader>" & LINE_BREAK & _
    "  <resheader name=""version"">" & LINE_BREAK & _
    "    <value>2.0</value>" & LINE_BREAK & _
    "  </resheader>" & LINE_BREAK & _
    "  <resheader name=""reader"">" & LINE_BREAK & _
    "    <value>System.Resources.ResXResourceReader, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</value>" & LINE_BREAK & _
    "  </resheader>" & LINE_BREAK & _
    "  <resheader name=""writer"">" & LINE_BREAK & _
    "    <value>System.Resources.ResXResourceWriter, System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</value>" & LINE_BREAK & _
    "  </resheader>" & LINE_BREAK

    DESIGN_CODE1 = _
    "//------------------------------------------------------------------------------" & LINE_BREAK & _
    "// <auto-generated>" & LINE_BREAK & _
    "//     This code was generated by a tool." & LINE_BREAK & _
    "//     Runtime Version:4.0.30319.42000" & LINE_BREAK & _
    "//" & LINE_BREAK & _
    "//     Changes to this file may cause incorrect behavior and will be lost if" & LINE_BREAK & _
    "//     the code is regenerated." & LINE_BREAK & _
    "// </auto-generated>" & LINE_BREAK & _
    "//------------------------------------------------------------------------------" & LINE_BREAK & _
    "" & LINE_BREAK & _
    "namespace MyProject1 {" & LINE_BREAK & _
    "    using System;" & LINE_BREAK & _
    "" & LINE_BREAK & _
    "" & LINE_BREAK & _
    "    /// <summary>" & LINE_BREAK & _
    "    ///   A strongly-typed resource class, for looking up localized strings, etc." & LINE_BREAK & _
    "    /// </summary>" & LINE_BREAK & _
    "    // This class was auto-generated by the StronglyTypedResourceBuilder" & LINE_BREAK & _
    "    // class via a tool like ResGen or Visual Studio." & LINE_BREAK & _
    "    // To add or remove a member, edit your .ResX file then rerun ResGen" & LINE_BREAK & _
    "    // with the /str option, or rebuild your VS project." & LINE_BREAK & _
    "    [global::System.CodeDom.Compiler.GeneratedCodeAttribute(""System.Resources.Tools.StronglyTypedResourceBuilder"", ""17.0.0.0"")]" & LINE_BREAK & _
    "    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]" & LINE_BREAK

    DESIGN_CODE2 = _
    "    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]" & LINE_BREAK & _
    "    internal class Strings {" & LINE_BREAK & _
    "" & LINE_BREAK & _
    "        private static global::System.Resources.ResourceManager resourceMan;" & LINE_BREAK & _
    "" & LINE_BREAK & _
    "        private static global::System.Globalization.CultureInfo resourceCulture;" & LINE_BREAK & _
    "" & LINE_BREAK & _
    "        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(""Microsoft.Performance"", ""CA1811:AvoidUncalledPrivateCode"")]" & LINE_BREAK & _
    "        internal Strings() {" & LINE_BREAK & _
    "        }" & LINE_BREAK & _
    "" & LINE_BREAK & _
    "        /// <summary>" & LINE_BREAK & _
    "        ///   Returns the cached ResourceManager instance used by this class." & LINE_BREAK & _
    "        /// </summary>" & LINE_BREAK & _
    "        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]" & LINE_BREAK & _
    "        internal static global::System.Resources.ResourceManager ResourceManager {" & LINE_BREAK & _
    "            get {" & LINE_BREAK & _
    "                if (object.ReferenceEquals(resourceMan, null)) {" & LINE_BREAK & _
    "                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager(""MyProject1.Strings"", typeof(Strings).Assembly);" & LINE_BREAK & _
    "                    resourceMan = temp;" & LINE_BREAK & _
    "                }" & LINE_BREAK & _
    "                return resourceMan;" & LINE_BREAK & _
    "            }" & LINE_BREAK

    DESIGN_CODE3 = _
    "        }" & LINE_BREAK & _
    "" & LINE_BREAK & _
    "        /// <summary>" & LINE_BREAK & _
    "        ///   Overrides the current thread's CurrentUICulture property for all" & LINE_BREAK & _
    "        ///   resource lookups using this strongly typed resource class." & LINE_BREAK & _
    "        /// </summary>" & LINE_BREAK & _
    "        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]" & LINE_BREAK & _
    "        internal static global::System.Globalization.CultureInfo Culture {" & LINE_BREAK & _
    "            get {" & LINE_BREAK & _
    "                return resourceCulture;" & LINE_BREAK & _
    "            }" & LINE_BREAK & _
    "            set {" & LINE_BREAK & _
    "                resourceCulture = value;" & LINE_BREAK & _
    "            }" & LINE_BREAK & _
    "        }" & LINE_BREAK

    oDoc = ThisComponent
    oSheet = ThisComponent.CurrentController.ActiveSheet

	' get cells range, all of sheet contents
	oCursor = oSheet.createCursor()
	oCursor.gotoStartOfUsedArea(False)
	oCursor.gotoEndOfUsedArea(True)

	' determine how many rows and columns
	col_start = oCursor.RangeAddress.StartColumn
	row_start = oCursor.RangeAddress.StartRow

	col_end = oCursor.RangeAddress.EndColumn
	row_end = oCursor.RangeAddress.EndRow

	' ---- Generate .resx files per language ----
	For x = 1 To col_end

	    strLangCode = LCase(oSheet.getCellByPosition(x, 1).String)

	    strLangExt = ""
	    If x > 1 Then strLangExt = "." & strLangCode

	    strFilename = FilePath() & OUTPUT_DIR & _
	                  OUTPUT_FILENAME & strLangExt & ".resx"

	    strContents = "<?xml version=""1.0"" encoding=""utf-8""?>" & LINE_BREAK & _
	                  "<root>" & LINE_BREAK

	    strContents = strContents & _
	                  "  <!--" & LINE_BREAK & _
	                  "    Microsoft ResX Schema" & LINE_BREAK & LINE_BREAK & _
	                  "    Version 2.0" & LINE_BREAK & LINE_BREAK & _
	                  "    " & oSheet.Name & " (" & _
	                  oSheet.getCellByPosition(x, 0).String & ")" & LINE_BREAK & _
	                  "    Translation by " & _
	                  oSheet.getCellByPosition(x, 3).String & LINE_BREAK & LINE_BREAK & _
	                  "    Generated: " & _
	                  Format(Now, "DD-MMM-YYYY HH:MM") & LINE_BREAK & _
	                  "  -->" & LINE_BREAK

	    strContents = strContents & VS_XSD1 & VS_XSD2 & VS_XML_HEADER

	    For y = 5 To row_end
	        strKey = oSheet.getCellByPosition(0, y).String
	        strValue = oSheet.getCellByPosition(x, y).String

	        If strKey = "" Then
	            ' skip
	        ElseIf Left(strKey, 2) = "//" Then
	            strContents = strContents & _
	                          "  <!--" & Mid(strKey, 3) & " -->" & LINE_BREAK
	        Else
	            strContents = strContents & _
	                          "  <data name=""" & ReplaceXmlKey(strKey) & """>" & LINE_BREAK & _
	                          "    <value>" & ReplaceXmlValue(strValue) & "</value>" & LINE_BREAK & _
	                          "  </data>" & LINE_BREAK
	        End If
	    Next y

	    strContents = strContents & "</root>"

		WriteToTextFile(strFilename, strContents)

	Next x

	' ---- Generate C# Designer file ----
	strFilename = FilePath() & OUTPUT_DIR & OUTPUT_FILENAME & ".Designer.resx"

	' initialise contents
	strContents = DESIGN_CODE1 & DESIGN_CODE2 & DESIGN_CODE3

	For y = 5 To row_end
	    strKey = oSheet.getCellByPosition(0, y).String
	    strValue = oSheet.getCellByPosition(1, y).String

	    If strKey = "" Then
	    ElseIf Left(strKey, 2) = "//" Then
	        strContents = strContents & _
	                      "        // ---- " & Mid(strKey, 3) & " ----" & LINE_BREAK
	    Else
	        strContents = strContents & _
	                      "        /// <summary>" & LINE_BREAK & _
	                      "        ///   Looks up a localized string: " & strValue & LINE_BREAK & _
	                      "        /// </summary>" & LINE_BREAK & _
	                      "        internal static string " & strKey & " {" & LINE_BREAK & _
	                      "            get {" & LINE_BREAK & _
	                      "                return ResourceManager.GetString(""" & _
	                      strKey & """, resourceCulture);" & LINE_BREAK & _
	                      "            }" & LINE_BREAK & _
	                      "        }" & LINE_BREAK
	    End If
	Next y

	strContents = strContents & LINE_BREAK & "    }" & LINE_BREAK & "}"

	WriteToTextFile(strFilename, strContents)

	MsgBox "Translation files created in folder " & OUTPUT_DIR, 64

End Sub

Sub UniqueCharactersFromSelection()
    ' Collect all unique characters from the selected cells
    ' Useful when creating bitmap fonts

    Dim oDoc As Object
    Dim oSel As Object
    Dim oCell As Object
    Dim oRange As Object

    Dim strFilename As String
    Dim txt As String
    Dim result As String

    Dim i As Long, j As Long, k As Long
    Dim code As Long
    Dim codes() As Long
    Dim codeCount As Long
    Dim exists As Boolean
    Dim temp As Long

    oDoc = ThisComponent
    oSel = oDoc.getCurrentSelection()

    ' Build output filename (same folder as document)
	strFilename = FilePath() & "Unique_characters_output.txt"

    ' content of output
    result = "; Manage-translations, all unique characters in text" & Chr(10) & _
             "; useful when creating bitmap fonts" & Chr(10) & _
             "; Generated on: " & Format(Now, "DD-MMM-YYYY HH:MM")

    ' Combine all text from selected cells

	Dim oRanges As Object

	oSel = ThisComponent.getCurrentSelection()

	' Single range
	If oSel.supportsService("com.sun.star.sheet.SheetCellRange") Then

	    For r = 0 To oSel.Rows.getCount() - 1
	        For c = 0 To oSel.Columns.getCount() - 1
	            oCell = oSel.getCellByPosition(c, r)
	            If oCell.String <> "" Then
	                txt = txt & oCell.String
	            End If
	        Next c
	    Next r
	Else
	    MsgBox "Unsupported selection type", 48
	    Exit Sub
	End If

    ' Do it twice
    For k = 1 To 2

        result = result & Chr(10) & Chr(10)

        If (k = 1) Then
            result = result & "; unique characters in selection" & Chr(10)
        Else
            result = result & "; unique characters + minimal complete a-z A-Z 0-9" & Chr(10)

            ' Add minial complete a-z A-Z 0-9 characters
            txt = txt & "1234567890!@#$%^&*()-=+ abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        End If

        ' Reset array for each pass
        Erase codes
        codeCount = 0

        ' Collect unique character codes
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
	WriteToTextFile(strFilename, result)

    MsgBox result, 64, "Done"

End Sub

' --------------------------------------
' Helper subs and functions
' --------------------------------------
Private Function ReplaceXmlKey(sXmlKey As String) As String

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

Private Function ReplaceXmlValue(sXmlValue As String) As String

	' prepare xml-safe-value
	sXmlValue = Replace(sXmlValue, "&", "&amp;")
	sXmlValue = Replace(sXmlValue, "<", "&lt;")
	sXmlValue = Replace(sXmlValue, ">", "&gt;")
	sXmlValue = Replace(sXmlValue, """", "&quot;")
	sXmlValue = Replace(sXmlValue, "'", "&apos;")

	ReplaceXmlValue = sXmlValue
End Function

Private Function FormatCsvValue(sCsvValue As String) As String

    ' csv format
    sCsvValue = Replace(sCsvValue, """", """""") 'escape double quote
    sCsvValue = """" & sCsvValue & """"

    FormatCsvValue = sCsvValue

End Function

Private Function FilePath() As String
	' Returns file path excluding trailing separator and file name
	' The result is an array of two elements.
	Dim sTemp As String  ' temporary string variable - the URL of the current document
	Dim aTemp As Variant ' temporary variable of the variant - later it will be an array
	Dim sFileName As String, sFilePath As String ' the purpose of these variables is clear from their names
	sTemp = ConvertFromURL( ThisComponent.getUrl() )	' get URL of current document
	If sTemp = "" Then ' if document is new (not saved) then URL is empty string
		sFileName = "file name undefined"
		sFilePath = "file path undefined"
	Else
		aTemp = Split( sTemp, GetPathSeparator() )
		sFileName = aTemp( Ubound(aTemp) ) ' last element of array is a part of string after last PathSeparator
		sFilePath = Left( sTemp, Len( sTemp ) - Len( sFileName ) ) ' rest of URL-string is path
	EndIf
	FilePath = sFilePath
End Function

Private Function WriteToTextFile(sFilename as String, sText as String) As Boolean

	Dim oSFA As Object, oOutText As Object
	Dim FileURL As String

	' delete if already exists
	If FileExists(sFilename) Then
		Kill(sFilename)
	End If

	' convert the filenaem to URL
	oSFA = createUNOService("com.sun.star.ucb.SimpleFileAccess")
	FileURL = ConvertToURL(sFilename)

	oOutText = createUNOService("com.sun.star.io.TextOutputStream")
	oOutText.setOutputStream(oSFA.openFileWrite(FileURL))

	' write to file
	' NOTE: this will also automatically create directories that don't exists yet
	oOutText.WriteString(sText)

	' flush buffers ans close
	oOutText.flush
End Function
