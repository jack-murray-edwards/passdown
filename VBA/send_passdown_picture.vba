Sub send_passdown_picture()
'Macro to send out a passdown screenshot with one button
'Jack.Murray@edwardsvacuum.com
'v1.2 2023-06-21
  
  'Setup savepath as Users local temp folder with sheet name as filename
  Dim savePath As String
  Set fSys = CreateObject("Scripting.FileSystemObject")
  Const tempFolder = 2
  tempFolderPath = fSys.GetSpecialFolder(tempFolder)
  savePath = tempFolderPath & "\" & ActiveSheet.Name & ".jpg"
  'MsgBox "savePath is" & savePath
  
  'Select the range for the screenshot
  lastRowToday = Range("A7").End(xlDown).Row
  lastRowLookAhead = Range("I7").End(xlDown).Row
  If (lastRowToday > lastRowLookAhead) Then
      Range("A1:N" & lastRowToday).Select
  Else
      Range("A1:N" & lastRowLookAhead).Select
  End If

  'Set up temporary chart for image export
  Dim tempChart As ChartObject
  Dim ActiveShape As Shape

  'Confirm if a Cell Range is currently selected
  If TypeName(Selection) <> "Range" Then
    MsgBox "No Cell Range Selected.  Screenshot Sending Failed."
    Exit Sub
  End If
  
  'Copy/Paste Cell Range as a Picture
  Selection.Copy
  ActiveSheet.Pictures.Paste(link:=False).Select
  Set ActiveShape = ActiveSheet.Shapes(ActiveWindow.Selection.Name)
  
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
  'ActiveChart.Delete

  'create a new instance of Outlook
  Dim appOutlook As Object
  Dim mailItem As Object
  Set appOutlook = CreateObject("Outlook.Application")
  Set mailItem = appOutlook.CreateItem(0)
  With mailItem
    .To = "robert.nolan@edwardsvacuum.com"
    '.CC = "jack.murray@edwardsvacuum.com"
    .Subject = "passdown screenshot for " & ActiveSheet.Name
    .Body = "Attached is the passdown image for " & ActiveSheet.Name
    .Attachments.Add savePath
    .Display
  End With
  'MsgBox "screenshot maybe sent out" & vbNewLine & "Current user is: " & Application.UserName
  
  'clean up objects
  Set mailItem = Nothing
  Set appOutlook = Nothing
  
End Sub



