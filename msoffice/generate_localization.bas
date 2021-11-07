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
 
    ' create eclipse directory if not exist
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
                            strContents = strContents & vbCrLf & vbTab & vbTab & """" & Replace(strKey, """", """""") & """: """ & Replace(strValue, """", """""") & ""","
                        End If
                    End If
                Next y

                ' remove last comma
                strContents = Left(strContents, Len(strContents) - 1)
                
                ' close current language, and add to total
                strContents = strContents & vbCrLf & vbTab & "}" & vbCrLf
                strContentsAll = strContentsAll & strContents

                ' close
                strContents = "{" & vbCrLf & strContents & "}" & vbCrLf

                ' write to single languages file
                If FileOrDirExists(strFilename) Then
                    Kill strFilename
                End If
                Call SaveToFile(strFilename, strContents)

            Next x
 
        End If
        
        ' open and closing brackets
        strContentsAll = "{" & vbCrLf & strContentsAll & "}" & vbCrLf
    
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
 
    ' create eclipse directory if not exist
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
                        strContents = strContents & """" & Replace(strKey, """", """""") & """ = """ & Replace(strValue, """", """""") & """;" & Chr$(10) ' char 10 = Unix linefeed
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
                        strContents = strContents & vbTab & "<string name=""" & ReplaceXmlKey(strXmlKey) & """>" & ReplaceXmlValue(strValue) & "</string>" & Chr$(10)  ' char 10 = Unix linefeed
                    End If
                Next y

                ' close XML
                strContents = strContents & Chr$(10) & "</resources>"

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
    Dim strCode As String
    Dim strKey As String
    Dim strXmlKey As String
    Dim strValue As String
    Dim strContents As String
 
    ' create eclipse directory if not exist
    Const OUTPUT_DIR = "visualstudio\"
    Call FolderCreate(ActiveWorkbook.path & "\" & OUTPUT_DIR)
 
    ' check all worksheets
    For Each ws In Worksheets
 
        ' only the active worksheet
        If ws.Name = ActiveSheet.Name Then

            ' get all field definition
            For x = 2 To ws.Cells.SpecialCells(xlLastCell).Column
            
                ' use worksheet name and language code as filename
                strLangCode = LCase(ws.Cells(2, x))
                strFilename = ActiveWorkbook.path & "\" & OUTPUT_DIR & "Resources." & strLangCode & ".resx"
                
                ' initialise contents
                strContents = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf & "<root>" & vbCrLf
                
                ' initialise comment
                strContents = strContents & vbTab & "<!--" & vbCrLf
                strContents = strContents & vbTab & "Visual Studio translation" & vbCrLf
                strContents = strContents & vbTab & ws.Name & " (" & ws.Cells(1, x) & ")" & vbCrLf
                strContents = strContents & vbTab & "Translation by " & ws.Cells(4, x) & vbCrLf & vbCrLf

                strContents = strContents & vbTab & "Generated: " & Format(Now(), "dd-mmm-yyyy hh:nn") & vbCrLf
                strContents = strContents & vbTab & "-->" & vbCrLf & vbCrLf

                For y = 6 To ws.Cells.SpecialCells(xlLastCell).Row
                    
                    ' get key and value
                    strKey = ws.Cells(y, 1)
                    strValue = ws.Cells(y, x)

                    If (strKey = "") Then
                        ' empty line
                        strContents = strContents & vbCrLf
                    ElseIf (Left$(strKey, 2) = "//") Then
                        ' comment line
                        strContents = strContents & vbTab & "<!--" & Mid$(strKey, 3) & " -->" & vbCrLf
                    Else
                        ' language key value
                        strContents = strContents & vbTab & "<data name=""" & ReplaceXmlKey(strKey) & """>" & vbCrLf
                        strContents = strContents & vbTab & vbTab & "<value>" & ReplaceXmlValue(strValue) & "</value>" & vbCrLf
                        strContents = strContents & vbTab & "</data>" & vbCrLf
                    End If
                Next y

                ' close XML
                strContents = strContents & vbCrLf & "</root>"

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


' --------------------------------------
' Helper subs and functions
' --------------------------------------
Function ReplaceXmlKey(sXmlKey As String) As String

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

Function ReplaceXmlValue(sXmlValue As String) As String

    ' prepare xml-safe-value
    sXmlValue = Replace(sXmlValue, "&", "&amp;")
    sXmlValue = Replace(sXmlValue, "<", "&lt;")
    sXmlValue = Replace(sXmlValue, ">", "&gt;")
    sXmlValue = Replace(sXmlValue, """", "&quot;")
    sXmlValue = Replace(sXmlValue, "'", "\'")

    ReplaceXmlValue = sXmlValue
End Function
