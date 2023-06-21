For Each ws In Sheets
        If ws.Name <> "Contents" And ws.Name <> "Passdown Summary" And ws.Name <> "passdown_assets" Then
            Range("P12").Select
            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
                "'Passdown Summary'!A1", TextToDisplay:="'Passdown Summary'!A1"
            Selection.Hyperlinks(1).TextToDisplay = "Passdown Summary"
            Range("P12").Select
            Selection.Style = "Hyperlink"
            Selection.Font.Bold = True
            Range("P12").Select
            With Selection.Font
                .Name = "Calibri"
                .Size = 12
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleSingle
                .ThemeColor = xlThemeColorHyperlink
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
        End If
    Next ws
End Sub