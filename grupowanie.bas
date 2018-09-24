Attribute VB_Name = "Module21"
Option Explicit
Sub Grupowanie()
    Dim bBreak As Boolean
    Dim i As Long
    Dim n As Range
    Dim lRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim WS As Worksheet
    Dim dict As Object

    Set WS = ActiveSheet

    'Specify the range which suits your purpose
    Set n = Range("a1", ActiveSheet.Range("a65536").End(xlUp))
    
    'Size of rows and columns if needed
    WS.Rows.RowHeight = 15
    WS.Columns("A:N").ColumnWidth = 13
    
    With n
        Selection.NumberFormat = "0"
        .Value = .Value
    End With

    With WS
        lRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With

    'Define a dictionary to store values that should be grouped
    Set dict = CreateObject("scripting.dictionary")

    dict.Add 150, ""
    dict.Add 155, ""
    dict.Add 130, ""
    dict.Add 115, ""
    dict.Add 110, ""
    dict.Add 117, ""

    'Loop through the values in Column A and group them
    For i = 1 To lRow
        Set cell = WS.Cells(i, 1)
        If dict.exists(cell.Value) Then 'cell value is among values store in the dict, so it should be grouped
            If rng Is Nothing Then
                Set rng = cell 'first cell with a value to be grouped
            Else
                Set rng = Application.Union(rng, cell)
            End If
        Else 'cell is a breaker value
            If rng Is Nothing Then 'there is nothing to be grouped, continue
                'do nothing
            Else 'this is the last cell of a range that should be grouped
                rng.Rows.Group 'group only rows
                Set rng = Nothing 'reset the range
            End If
        End If
    Next i

    'If the last cell is not a breaker, the last cells will not be grouped
    If Not rng Is Nothing Then
        rng.Rows.Group 'group only rows
    End If


End Sub
