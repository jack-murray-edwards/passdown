Sub close_passdown ()
'create a new instance of Outlook
    Dim appOutlook As Object
    Dim mailItem As Object
    Set appOutlook = CreateObject("Outlook.Application")
    Set mailItem = appOutlook.CreateItem(0)
    With mailItem
        .To = "figure_this_out@later.com"
        .CC = "figure_this_out@later.com"
        .Subject = "passdown screenshot for " & ActiveSheet.Name
        .Body = "Attached is the passdown image for " & ActiveSheet.Name
        .Attachments.Add savePath
        .Display
    End With

    
