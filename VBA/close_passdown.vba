Sub close_passdown ()
'Macro to close out the passdown, copy todays jobs to the summary sheet
'and send out a passdown screenshot with one button
'Jack.Murray@edwardsvacuum.com
'v1.0 2023-06-21

    Dim mailItem As Object
    Set appOutlook = CreateObject("Outlook.Application")
    Dim savePath As String

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

    'Setup savepath as Users local temp folder with sheet name as filename
    Set fSys = CreateObject("Scripting.FileSystemObject")
    Const tempFolder = 2
    tempFolderPath = fSys.GetSpecialFolder(tempFolder)
    savePath = tempFolderPath & "\" & ActiveSheet.Name & ".jpg"

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
        .To = "figure_this_out@later.com" 'TODO create loop here to add all emails from a range on the sheet
        .CC = "figure_this_out@later.com" 'TODO edwards iq dist.
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
        MsgBox "Passdown already closed"
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
    ActiveCell.FormulaR1C1 = "passdown closed" 'FIXME does this need to be ActiveCell.Value?
    Range("AA1").Select
    Selection.Style = "Good"
    Selection.Font.Bold = True
    



End Sub


