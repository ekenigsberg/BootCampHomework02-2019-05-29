Sub SummarizeSingleWorksheet(wrk As Worksheet)
    ' SET VARS
    ' ranges for iteration
    Dim rngTkrSource, rngTkrSummary As Range
    ' ranges for building summary-table formulas
    Dim rngPriceOpen, rngPriceClose As Range
    ' var for Total Stock Volume. I show Total Stock Volume on the spreadsheet as a formula, but I
    ' use this variable to prove I understand how to compute a running total in VBA.
    Dim dblVolume As Double
    ' var for last row in Summary table
    Dim dblLastSummRow As Double
    ' array for "Greatest" table
    Dim varGreatest(0 To 2, 0 To 1) As Variant
    
    ' PREPARE SHEET
    With wrk
        .Activate
        ' freeze panes
        .[B2].Select
        ActiveWindow.FreezePanes = True
        ' create names
        .[A1].CurrentRegion.CreateNames Top:=True, Bottom:=False, Left:=False, Right:=False
        ' establish summary sections' titles
        .[I1:L1] = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        .[O2:O4] = Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
        .[P1:Q1] = Array("Ticker", "Value")
        ' set rngTkrSummary
        Set rngTkrSummary = .[I2]
    End With
    
    ' ITERATE THROUGH EACH CELL IN "<ticker>" RANGE AND POPULATE TICKER-LEVEL SUMMARY CELLS
    For Each rngTkrSource In ActiveSheet.[ticker]
        ' IF we're on first row of a given ticker's data THEN set rngPriceOpen
        If rngTkrSource <> rngTkrSource.Offset(-1, 0) Then Set rngPriceOpen = rngTkrSource.Offset(0, 2)
        ' increment dblVolume
        dblVolume = dblVolume + rngTkrSource.Offset(0, 6)
        ' IF we're on last row of a given ticker's data THEN (1) set rngPriceClose, (2) populate summary cells, (3) reset dblVolume, (4) populate varGreatest(), (5) increment rngTkrSummary
        If rngTkrSource <> rngTkrSource.Offset(1, 0) Then
            ' 1) set rngPriceClose
            Set rngPriceClose = rngTkrSource.Offset(0, 5)
            ' 2a) populate summary cells. protect against #DIV/0! errors by arbitrarily making Percent Change equal zero
            rngTkrSummary = rngTkrSource
            rngTkrSummary.Offset(0, 1) = "=" & rngPriceClose.Address(False, False) & " - " & rngPriceOpen.Address(False, False)
            rngTkrSummary.Offset(0, 2) = "=IF(" & rngPriceOpen.Address(False, False) & " = 0, 0, " & rngTkrSummary.Offset(0, 1).Address(False, False) & " / " & rngPriceOpen.Address(False, False) & ")"
            rngTkrSummary.Offset(0, 3) = "=SUM(" & rngPriceOpen.Offset(0, 4).Address(False, False) & ":" & rngPriceClose.Offset(0, 1).Address(False, False) & ")"
            ' 2b) use dblVolume for something! confirm that the "Total Stock Volume" cell matches dblVolume
            If rngTkrSummary.Offset(0, 3) <> dblVolume Then
                MsgBox "There's been an error. The Total Stock Volume computed for " & rngTkrSource & " (" & dblVolume & _
                 " does not match the Total Stock Volume returned in the spreadsheet (" & rngTkrSummary.Offset(0, 3) & ")."
                 Exit Sub
            End If
            ' 3) reset dblVolume
            dblVolume = 0
            ' 4) populate varGreatest() array
            If rngTkrSummary.Offset(0, 2) > varGreatest(0, 1) Then ' Greatest % Increase
                varGreatest(0, 0) = rngTkrSummary
                varGreatest(0, 1) = rngTkrSummary.Offset(0, 2)
            End If
            If rngTkrSummary.Offset(0, 2) < varGreatest(1, 1) Then ' Greatest % Decrease
                varGreatest(1, 0) = rngTkrSummary
                varGreatest(1, 1) = rngTkrSummary.Offset(0, 2)
            End If
            If rngTkrSummary.Offset(0, 3) > varGreatest(2, 1) Then ' Greatest Total Volume
                varGreatest(2, 0) = rngTkrSummary
                varGreatest(2, 1) = rngTkrSummary.Offset(0, 3)
            End If
            ' 5) increment rngTkrSummary
            Set rngTkrSummary = rngTkrSummary.Offset(1, 0)
        End If
    Next rngTkrSource
    
    ' FORMAT SUMMARY SECTIONS
    ' set dblLastSummRow
    dblLastSummRow = wrk.[I2].CurrentRegion.Rows.Count
    ' create Conditional Formatting for "Yearly Change" (col J)
    With wrk.Range("J2:J" & dblLastSummRow)
        With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0")
            With .Interior
                .PatternColorIndex = xlAutomatic
                .Color = 5287936
                .TintAndShade = 0
            End With
        End With
        With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
            With .Interior
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
            End With
        End With
    End With
    With wrk
        ' format "Yearly Change" (col J)
        .Columns("J").NumberFormat = "#,##0.00"
        ' format "Percent Change" (col K)
        .Columns("K").NumberFormat = "#,##0.0%"
        ' format "Total Stock Volume" (col L)
        .Columns("L").NumberFormat = "#,##0"
        ' populate sheet-level summary
        .[Q2:P4] = varGreatest
        ' format sheet-level summary cells
        .[Q2:Q3].NumberFormat = "#,##0.0%"
        .[Q4].NumberFormat = "#,##0"
        ' autofit all columns
        .Columns("A:Q").AutoFit
    End With
End Sub
