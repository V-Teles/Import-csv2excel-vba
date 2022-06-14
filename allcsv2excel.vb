Option Explicit

Sub Macro1()
'
' Macro1 Macro
'
    Dim strFile As String, strPath As String

    strPath = "C:\Users\victor\Desktop\DQ\"
    strFile = Dir(strPath & "*.csv")

    While strFile <> ""

        ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
        ActiveWorkbook.Sheets(Worksheets.Count).Name = Right(Left(strFile, InStr(strFile, ".csv") - 1), InStr(strFile, ".csv") - 7)
        
        With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & strPath & strFile, Destination:=Range("$A$1"))
    '        .CommandType = 0
            .Name = Left(strFile, InStr(strFile, ".csv") - 1)
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 437
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileOtherDelimiter = "|"
            .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
            2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2 _
            , 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
            
            
        End With
        
        Selection.AutoFilter
        strFile = Dir
        
    Wend

End Sub










