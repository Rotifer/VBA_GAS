Attribute VB_Name = "modMarkdownGenerator"
Option Explicit

' Utility function to return an array of strings for all the data values in the given range.
Private Function GetRangeCellsAsArray(rng As Range) As String()
    Dim arr() As String
    Dim cell As Range
    Dim i As Long: i = 0
    For Each cell In rng.Cells
        ReDim Preserve arr(i)
        arr(i) = CStr(cell.Value)
        i = i + 1
    Next cell
    GetRangeCellsAsArray = arr
End Function

' Create Markdown table hyphen line for a table with a given cell alignment ("left", "mid" or "right") and a given number of table columns.
Private Function GetHyphenLine(alignment As String, columnCount As Integer) As String
    Dim arrHyphens() As String
    Dim i As Integer
    Dim columnHyphen As String
    If alignment = "left" Then
        columnHyphen = ":---"
    ElseIf alignment = "right" Then
        columnHyphen = "---:"
    Else
        columnHyphen = ":--:"
    End If
    For i = 0 To columnCount
        ReDim Preserve arrHyphens(i)
        arrHyphens(i) = columnHyphen
    Next i
    GetHyphenLine = "|" & Join(arrHyphens, "|") & "|"
End Function
' Generate the data rows lines of the Markdown table.
' Construct a Markdown table line for each range row but skip the first row (assumed
'  to be the column names row) by building an array of Markdown data row lines which are
'  then concatenated to generate the returned text.
Private Function GetDataRowsMarkdown(inputRng As Range) As String
    Dim row As Range
    Dim i As Integer: i = 0
    Dim dataRow As String
    Dim dataRows() As String
    For Each row In inputRng.Rows
        i = i + 1
        If i > 0 Then ' skip the column headers row
            dataRow = "|" & Join(GetRangeCellsAsArray(row), "|") & "|"
            ReDim Preserve dataRows(i - 1)
            dataRows(i - 1) = dataRow
        End If
    Next row
    GetDataRowsMarkdown = Join(dataRows, vbCr)
End Function
' Generate a valid Markdown table definition from an input sheet range.
' The table is generated in three parts: (1) the column names, (2) the hyphen line and (3) the data rows
' The cell alignment can be one of "left", "mid" (default) or "right"
' The generated Markdown can be pasted into a Markdown document and then rendered as HTML.
' Remember to remove the encosing double quotes.
Public Function TABLE_TO_MARKDOWN(inputRng As Range, Optional ByRef alignment As String = "mid") As String
    Dim columnNamesRow As Range: Set columnNamesRow = inputRng.Rows(1)
    Dim columnNames() As String: columnNames = GetRangeCellsAsArray(columnNamesRow)
    Dim columnNamesLine As String: columnNamesLine = "|" & Join(columnNames, "|") & "|"
    Dim columnCount As Integer: columnCount = inputRng.Columns.Count
    Dim hyphenLine As String: hyphenLine = GetHyphenLine(alignment, columnCount)
    Dim dataRowsMarkdown As String: dataRowsMarkdown = GetDataRowsMarkdown(inputRng)
    TABLE_TO_MARKDOWN = columnNamesLine & vbCr & hyphenLine & vbCr & dataRowsMarkdown
End Function

