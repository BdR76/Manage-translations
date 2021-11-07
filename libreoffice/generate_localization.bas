REM  *****  BASIC  *****

' Multilanguage localization
' --------------------------
' Manage translations words and phrases in spreadsheet
' and use this macro to export to different source files
' * GenerateLocalisationJson - JSON files for javascript
' * GenerateLocalisationXcode - XCode Localizable.strings files for iPhone
' * GenerateLocalisationEclipse - XML strings.xml files for Android
' * GenerateLocalisationVisualStudio - .resx xml files for Visual Studio

' global String LINE_BREAK = "vbCrLf" ' Chr(13) & Chr(10) gives syntax error

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

		For row = HEADER_ROWS to row_end - 1
		
			str_key = oSheet.getCellByPosition(0, row).String
			str_val = oSheet.getCellByPosition(col, row).String
			If (str_key <> "") Then
				If (Left(str_key, 2) = "//") Then
					' JSON format does not support comment lines
					' sContent = sContent & str_key & " COMMENT" & LINE_BREAK
				Else
					' phrase value
					sContent = sContent & LINE_BREAK & "		""" & Replace(str_key, """", """""") & """: """ & Replace(str_val, """", """""") & ""","
				End If
			End If

		Next row

		' remove last comma
		sContent = Left(sContent, Len(sContent) - 1)

		' close current language, and add to total
		sContent = sContent & LINE_BREAK & "	}" & LINE_BREAK
		sContentAll = sContentAll & sContent
		
		' open and closing brackets
		sContent = "{" & LINE_BREAK & sContent & "}" & LINE_BREAK

		' write to single languages file
		WriteToTextFile(sFilename, sContent)
	Next col

	' write to all languages combined in one file
	sFilename = FilePath() & OUTPUT_DIR & "all_translations.json"

	' open and closing brackets
	sContentAll = "{" & LINE_BREAK & sContentAll & "}" & LINE_BREAK

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
		sContent = sContent & "	" & LanguageDisplay & " (" & LanguageEn & ")" & LINE_BREAK
		sContent = sContent & "	Translation by " & LanguageTranslator & LINE_BREAK & LINE_BREAK

		sContent = sContent & "	Generated: " & Format(Now(), "dd-mmm-yyyy hh:mm") & LINE_BREAK
		sContent = sContent & "*/" & LINE_BREAK & LINE_BREAK

		For row = HEADER_ROWS to row_end - 1
		
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
				sContent = sContent & """" & Replace(str_key, """", """""") & """ = """ & Replace(str_val, """", """""") & """;" & LINE_BREAK ' char 10 = Unix linefeed
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
        sContent = sContent & "	" & LanguageDisplay & " (" & LanguageEn & ")" & LINE_BREAK
        sContent = sContent & "	Translation by " & LanguageTranslator & LINE_BREAK & LINE_BREAK

        sContent = sContent & "	Generated: " & Format(Now(), "dd-mmm-yyyy hh:mm") & LINE_BREAK
        sContent = sContent & "	-->" & LINE_BREAK

		For row = HEADER_ROWS to row_end - 1
		
			str_key = oSheet.getCellByPosition(0, row).String
			str_val = oSheet.getCellByPosition(col, row).String
			If (str_key <> "") Then
				If (Left(str_key, 2) = "//") Then
					' comment lines
					sContent = sContent & "	<!-- " & Trim(Mid(str_key, 3)) & " -->" & LINE_BREAK
				Else
					' language key value
					sContent = sContent & "	<string name=""" & ReplaceXmlKey(str_key) & """>" & ReplaceXmlValue(str_val) & "</string>" &LINE_BREAK
				End If
			End If

		Next row

		' close current language, and add to total
		sContent = sContent & "</resources>" & LINE_BREAK
		
		WriteToTextFile(sFilename, sContent)
	Next col
	
	MsgBox ("Translation files for " & col_end & " language created in folder " & OUTPUT_DIR)
	
End Sub

' --------------------------------------
' Export .resx xml files Visual Studio
' --------------------------------------
Sub GenerateLocalisationVisualStudio

	Dim OUTPUT_DIR
	Dim LINE_BREAK
	Dim HEADER_ROWS
	LINE_BREAK = Chr(13) & Chr(10) ' windows line breaks
	HEADER_ROWS = 5
	
	OUTPUT_DIR = "visualstudio\"

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
		sFilename = FilePath() & OUTPUT_DIR & oSheet.Name & "." & LanguageCode & ".resx"

		' start file content
		sContent = "<?xml version=""1.0"" encoding=""utf-8""?>" & LINE_BREAK & "<root>" & LINE_BREAK
		
        ' initialise comment
        sContent = sContent & "	<!--" & LINE_BREAK
        sContent = sContent & "	Visual Studio localization resource" & LINE_BREAK
        sContent = sContent & "	" & LanguageDisplay & " (" & LanguageEn & ")" & LINE_BREAK
        sContent = sContent & "	Translation by " & LanguageTranslator & LINE_BREAK & LINE_BREAK

        sContent = sContent & "	Generated: " & Format(Now(), "dd-mmm-yyyy hh:mm") & LINE_BREAK
        sContent = sContent & "	-->" & LINE_BREAK

		For row = HEADER_ROWS to row_end - 1
		
			str_key = oSheet.getCellByPosition(0, row).String
			str_val = oSheet.getCellByPosition(col, row).String
			If (str_key <> "") Then
				If (Left(str_key, 2) = "//") Then
					' comment lines
					sContent = sContent & "	<!-- " & Trim(Mid(str_key, 3)) & " -->" & LINE_BREAK
				Else
					' language key value
					sContent = sContent & "	<data name=""" & ReplaceXmlKey(str_key) & """>" & LINE_BREAK
					sContent = sContent & "		<value>" & ReplaceXmlValue(str_val) & "</value>" & LINE_BREAK
					sContent = sContent & "	</data>" & LINE_BREAK
				End If
			End If

		Next row

		' close current language, and add to total
		sContent = sContent & "</root>" & LINE_BREAK
		
		WriteToTextFile(sFilename, sContent)
	Next col
	
	MsgBox ("Translation files for " & col_end & " language created in folder " & OUTPUT_DIR)
	
End Sub


Sub FormatLayoutColors

	Dim oSheet as Object
	Dim oCursor
	Dim col_start
	Dim row_start
	Dim col_end
	Dim row_end

	Dim col, row
	Dim str_key, iscomment

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
	
	For col = col_start to col_end
		For row = row_start to row_end
		
			str_key = oSheet.getCellByPosition(0, row).String
			if (Left(str_key, 2) = "//") Then
				iscomment = True
			else
				iscomment = False
			End If
					
			If (row < 4) Then
				oSheet.getCellByPosition(col, row).CellBackColor = RGB(0, 255, 255) ' top meta data
			ElseIf (iscomment = True) Then
				oSheet.getCellByPosition(col, row).CellBackColor = RGB(204, 255, 204) ' comment line
			ElseIf (col = 0) Then
				oSheet.getCellByPosition(col, row).CellBackColor = RGB(204, 255, 255) ' left side keys
			Else
				oSheet.getCellByPosition(col, row).CellBackColor = RGB(255, 255, 255) ' white
				oSheet.getCellByPosition(col, row).setPropertyValue( "IsCellBackgroundTransparent", True)
				SetCellBorder(col, row, False, False, True, False) ' top bottom left right
				
			End If
	
			
			'  other options for reference
			' SetColWidth(col, 2.7, oSheet) ' in inch or cm
			' SetRowHeight(row, 0.8, oSheet) ' in inch or cm
			' SetCellBorder(col, row, False, False, True, False) ' top bottom left right
			' oSheet.getCellByPosition(col, row).CellBackColor = RGB(255, 255, 0)
			' oSheet.getCellByPosition(col, row+j).setFormula("=$STOCKDATA.B" & artline)
			' oSheet.getCellByPosition(col, row+j).CharFontName = "Libre Barcode 128"
			' oSheet.getCellByPosition(col, row+j).CharHeight = 16
			' oSheet.getCellByPosition(col, row+j).HoriJustify = com.sun.star.table.CellHoriJustify.CENTER ' LEFT, CENTER, RIGHT
			' oSheet.getCellByPosition(col, row+j).VertJustify = com.sun.star.table.CellVertJustify.TOP ' TOP, CENTER, BOTTOM
					
			' oSheet.getCellByPosition(col+1, row).setPropertyValue("IsTextWrapped", True)
			' oSheet.getCellByPosition(col+1, row).ParaIndent = (3 * 35.28) ' 1pt = 0.352778 mm, value in 0.01mm
					
			' oSheet.getCellRangeByPosition(col+1, row, col+1, row+3).merge(True)
		Next row
	Next col
	
	MsgBox ("Formatting layout colors is ready")
	
End Sub

' --------------------------------------
' Helper subs and functions
' --------------------------------------
Private Function ReplaceXmlKey(sXmlKey As String) As String

	' prepare xml-safe-key
	sXmlKey = LCase(sXmlKey)
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
	sXmlValue = Replace(sXmlValue, "'", "\'")

	ReplaceXmlValue = sXmlValue
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

Private Sub WriteToTextFile(sFilename as String, sText as String)
	
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
End Sub

sub SetColWidth(col as integer, wcm as single, oSheet as object)
	Static oColumn As Object
	
	oColumn = oSheet.getColumns.getByIndex(col)
	oColumn.Width = (wcm * 1000.0)
	
end sub

sub SetRowHeight(row as integer, hcm as single, oSheet as object)
	Static oRow As Object
	
	oRow = oSheet.getRows.getByIndex(row)
	oRow.Height = (hcm * 1000.0)
	
end sub

Sub SetCellBorder(x As Integer, y As Integer, bTop As Boolean, bBottom As Boolean, bLeft As Boolean, bRight As Boolean)
	Dim BasicBorder as New com.sun.star.table.BorderLine
	Dim oRange As Object
	Dim oBorder As Object

	REM get border from range
	oRange = thiscomponent.getcurrentcontroller.activesheet.getCellRangeByPosition(x, y, x, y)
	oBorder = oRange.TableBorder

	REM set border lines properties
	BasicBorder.Color = RGB(0, 0, 0)
	BasicBorder.InnerLineWidth = 0
	BasicBorder.OuterLineWidth = 2
	BasicBorder.LineDistance = 0

	REM set lines
	If (bTop)    Then oBorder.TopLine	= BasicBorder
	If (bBottom) Then oBorder.BottomLine= BasicBorder
	If (bLeft)   Then oBorder.LeftLine	= BasicBorder
	If (bRight)  Then oBorder.RightLine	= BasicBorder

	REM set border onto range
	oRange.TableBorder = oBorder
End Sub
