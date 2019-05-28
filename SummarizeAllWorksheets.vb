Sub SummarizeAllWorksheets()
    Dim fd As FileDialog
    Dim strFilePath As String
    Dim wbk As Workbook
    Dim wrk As Worksheet
    
    ' prompt user for Workbook
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Please select the file to summarize."
        .Filters.Clear
        .Filters.Add "Excel", "*.xls?"
        If .Show = True Then
            strFilePath = Dir(.SelectedItems(1))
            Set wbk = Workbooks.Open(strFilePath)
            ' if Workbook opened successfully, iterate through the sheets
            If wbk.Name <> "" Then
                For Each wrk In wbk.Worksheets
                    Call SummarizeSingleWorksheet(wrk)
                Next wrk
             End If
        End If
    End With
End Sub
