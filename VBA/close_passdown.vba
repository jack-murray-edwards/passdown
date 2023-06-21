Sub close_passdown ()
'create a new instance of Outlook
    Dim appOutlook As Object
    Dim mailItem As Object
    Set appOutlook = CreateObject("Outlook.Application")

    'clean up objects
    Set mailItem = Nothing
    Set appOutlook = Nothing

    'Select all cells that are not empty
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select

    'Copy the selected cells as a picture
    Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    'Create a temporary chart object (same size as shape)
    Set tempChart = ActiveSheet.ChartObjects.Add( _
        Left:=550, _
        Width:=ActiveShape.Width, _
        Top:=0, _
        Height:=ActiveShape.Height)
    
    'Format temporary chart to have a transparent background
    tempChart.ShapeRange.Fill.Visible = msoFalse
    tempChart.ShapeRange.Line.Visible = msoFalse

    'Copy/Paste Shape inside temporary chart
    ActiveShape.Copy
    tempChart.Activate
    ActiveChart.Paste

    'Save chart to User's temp folder as PNG File
    tempChart.Chart.Export savePath

    'Delete temporary objects
    tempChart.Delete
    ActiveShape.Delete

    'create a new instance of Outlook
    Set appOutlook = CreateObject("Outlook.Application")

    'create a new e-mail item
    Set mailItem = appOutlook.CreateItem(0)
        With mailItem
        .To = "figure_this_out@later.com"
        .CC = "figure_this_out@later.com"
        .Subject = "passdown screenshot for " & ActiveSheet.Name
        .Body = "Attached is the passdown image for " & ActiveSheet.Name
        .Attachments.Add savePath
        .Display
    End With

    'clean up objects
    Set mailItem = Nothing
    Set appOutlook = Nothing

    'check if the passdown has been closed already
    Range("AA1").Select
    If ActiveCell.Value = "passdown closed" Then
        MsgBox "Passdown already closed" & vbNewLine & "Current user is: " & Application.UserName
        Exit Sub
    End If

    'select all cells in the passdown table that are not empty
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select 'TODO set this for just the table

    'copy selected cells to the first open row in the passdown summary sheet
    Selection.Copy
    Sheets("Passdown Summary").Select
    Range("A" & Rows.Count).End(xlUp).Offset(1).Select
    ActiveSheet.Paste

    'create a messagebox to confirm that the passdown has been copied to the summary
    MsgBox "Passdown copied to summary" & vbNewLine & "Current user is: " & Application.UserName

    'create a flag in cell AA1 to indicate that the passdown has been closed
    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "passdown closed"
    Range("AA1").Select
    Selection.Style = "Good"
    Selection.Font.Bold = True
    



End Sub


