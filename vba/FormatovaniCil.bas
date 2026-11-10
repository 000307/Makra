Attribute VB_Name = "Module4"
Option Explicit

Private FBStage As String

' Keeps the historical entry-point name; this chart now uses SOC WITH overflows (main column P).
Public Sub FormatovaniBaterkaBezPretokuProcenta()
    Dim priorUpdating As Boolean
    Dim priorEvents As Boolean
    Dim ws As Worksheet
    Dim co As ChartObject
    Dim graf As chart
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim dataField As PivotField
    Dim item As Series
    Dim pass As Long
    Dim kind As Long
    Dim counts(1 To 3) As Long
    Dim foundConsumption As Boolean
    Dim fieldsChanged As Boolean
    Dim maxBattery As Double
    Dim errText As String
    Dim errCode As Long

    priorUpdating = Application.ScreenUpdating
    priorEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    FBStage = "Vyhledani zdrojoveho grafu"
    Set ws = FBTargetSheet()
    Set co = FBTargetChart(ws)
    Set graf = co.chart
    If graf.PivotLayout Is Nothing Then
        Err.Raise vbObjectError + 5200, , "Cilovy graf neni kontingencni."
    End If
    Set pt = graf.PivotLayout.PivotTable
    FBStage = "Vyber stavu baterie s pretoky"
    FBEnsureBatteryField pt, fieldsChanged

    ' Add the original-consumption measure only when it is absent.
    For Each dataField In pt.DataFields
        If FBNorm(dataField.sourceName) = "spotreba(mwh)" Then
            foundConsumption = True
            Exit For
        End If
    Next dataField
    If Not foundConsumption Then
        For Each pf In pt.PivotFields
            If FBNorm(pf.Name) = "spotreba(mwh)" Then
                Set dataField = pt.AddDataField(pf, "Sou" & ChrW(269) & "et z Spot" & ChrW(345) & "eba (MWh)", xlSum)
                dataField.NumberFormat = "0.000000000"
                foundConsumption = True
                Exit For
            End If
        Next pf
        If Not foundConsumption Then
            Err.Raise vbObjectError + 5201, , "Ve zdroji chybi pole Spotreba (MWh)."
        End If
        fieldsChanged = True
    End If
    If fieldsChanged Then
        If Not pt.RefreshTable Then
            Err.Raise vbObjectError + 5202, , "Kontingencni tabulku se nepodarilo aktualizovat."
        End If
    End If

    ' Types first, axes second: Excel can regroup series after changing a type.
    For Each item In graf.SeriesCollection
        kind = FBKind(item.Name)
        If kind = 1 Then
            If item.ChartType <> xlColumnClustered Then item.ChartType = xlColumnClustered
        ElseIf kind = 2 Or kind = 3 Then
            If item.ChartType <> xlLine Then item.ChartType = xlLine
        End If
    Next item

    For pass = 1 To 2
        For Each item In graf.SeriesCollection
            kind = FBKind(item.Name)
            If kind = 1 Then
                If item.AxisGroup <> xlSecondary Then item.AxisGroup = xlSecondary
            ElseIf kind = 2 Or kind = 3 Then
                If item.AxisGroup <> xlPrimary Then item.AxisGroup = xlPrimary
            End If
        Next item
    Next pass

    For Each item In graf.SeriesCollection
        kind = FBKind(item.Name)
        If kind > 0 Then counts(kind) = counts(kind) + 1
        If kind = 1 Then
            item.Format.Fill.Solid
            item.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
            item.Format.Fill.Transparency = 0.25
            item.Format.Line.Visible = msoFalse
            FBMaximum item.values, maxBattery
        ElseIf kind = 2 Or kind = 3 Then
            item.Format.Line.Visible = msoTrue
            item.Format.Line.Weight = 1.25
            item.MarkerStyle = xlMarkerStyleNone
            item.Smooth = False
            If kind = 2 Then
                item.Format.Line.ForeColor.RGB = RGB(0, 0, 255)
            Else
                item.Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            End If
        End If
    Next item

    If counts(1) = 0 Or counts(2) = 0 Or counts(3) = 0 Then
        Err.Raise vbObjectError + 5203, , "V grafu nebyly rozpoznany vsechny tri ukazatele."
    End If

    ' Redraw any regrouped series first, then restore ALL four axes last.
    graf.Refresh
    FBConfigureAxes graf, maxBattery

    FBStage = "Kontrola vyslednych os"
    FBValidateAxes graf
    FBStage = "Aktualizace tlacitka"
    FBEnsureButton ws, co
    Application.EnableEvents = priorEvents
    Application.ScreenUpdating = priorUpdating
    Exit Sub

Failed:
    errText = Err.Description
    errCode = Err.Number
    On Error Resume Next
    If Not graf Is Nothing Then FBHideSecondaryX graf
    Application.EnableEvents = priorEvents
    Application.ScreenUpdating = priorUpdating
    On Error GoTo 0
    MsgBox "Graf se nepodarilo zformatovat." & vbCrLf & "Krok: " & FBStage & vbCrLf & "Chyba " & errCode & ": " & errText, vbExclamation, "Formatovani grafu"
End Sub


Private Sub FBEnsureBatteryField(ByVal pt As PivotTable, ByRef fieldsChanged As Boolean)
    Dim i As Long
    Dim dataField As PivotField
    Dim sourceField As PivotField
    Dim item As PivotField
    Dim sourceName As String
    Dim desiredField As PivotField

    ' Remove only the other battery scenario from this specific PivotTable.
    For i = pt.DataFields.Count To 1 Step -1
        Set dataField = pt.DataFields(i)
        sourceName = FBNorm(dataField.sourceName)
        If sourceName = "aktualnistavbaterkybezpretoku(%)" Then
            dataField.Orientation = xlHidden
            fieldsChanged = True
        ElseIf sourceName = "aktualnistavbaterkyspretoky(%)" Then
            Set desiredField = dataField
        End If
    Next i
    If Not desiredField Is Nothing Then
        If desiredField.Position <> 1 Then
            desiredField.Position = 1
            fieldsChanged = True
        End If
        Exit Sub
    End If

    For Each item In pt.PivotFields
        If FBNorm(item.Name) = "aktualnistavbaterkyspretoky(%)" Then
            Set sourceField = item
            Exit For
        End If
    Next item
    If sourceField Is Nothing Then
        Err.Raise vbObjectError + 5218, , "Ve zdroji chybi stav baterie s pretoky ze sloupce P."
    End If
    Set dataField = pt.AddDataField(sourceField, "Sou" & ChrW(269) & "et z " & sourceField.Name, xlSum)
    dataField.NumberFormat = "0%"
    dataField.Position = 1
    fieldsChanged = True
End Sub

Private Sub FBConfigureAxes(ByVal graf As chart, ByVal maxBattery As Double)
    Dim decimalSeparator As String
    decimalSeparator = CStr(Application.International(xlDecimalSeparator))
    If Not Application.UseSystemSeparators Then
        decimalSeparator = Application.decimalSeparator
    End If

    FBStage = "Vytvoreni os a skryti druhe osy X"
    graf.HasAxis(xlCategory, xlPrimary) = True
    graf.HasAxis(xlCategory, xlSecondary) = True
    graf.HasAxis(xlValue, xlPrimary) = True
    graf.HasAxis(xlValue, xlSecondary) = True
    ' Hide the extra X immediately, before any other formatting operation.
    FBHideSecondaryX graf

    FBStage = "Osa X: casy a dny"
    With graf.Axes(xlCategory, xlPrimary)
        .ReversePlotOrder = False
        .CategoryType = xlCategoryScale
        .AxisBetweenCategories = True
        ' MultiLevel belongs to TickLabels, NOT Axis.
        .TickLabels.MultiLevel = True
        .TickLabels.Orientation = xlUpward
        .TickLabels.Font.Size = 8
        .TickLabels.Offset = 100
        .TickLabelSpacingIsAuto = True
        .TickLabelPosition = xlTickLabelPositionLow
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .Format.Line.Visible = msoTrue
        .Crosses = xlAxisCrossesMinimum
    End With
    With graf.Axes(xlCategory, xlSecondary)
        .ReversePlotOrder = False
        .CategoryType = xlCategoryScale
        .AxisBetweenCategories = True
        .Crosses = xlAxisCrossesMaximum
    End With

    FBStage = "Leva osa Y: spotreba v MWh"
    With graf.Axes(xlValue, xlPrimary)
        .ReversePlotOrder = False
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
        .MajorUnitIsAuto = True
        .TickLabels.Orientation = 0
        .DisplayUnit = xlNone
        .HasDisplayUnitLabel = False
        .TickLabelPosition = xlTickLabelPositionLow
        .TickLabels.NumberFormatLinked = False
        .TickLabels.NumberFormatLocal = "0" & decimalSeparator & "000######"
        .Crosses = xlAxisCrossesMinimum
    End With

    FBStage = "Prava osa Y: stav baterie v procentech"
    With graf.Axes(xlValue, xlSecondary)
        .ReversePlotOrder = False
        .MinimumScale = 0
        If maxBattery <= 1 Then
            .MaximumScale = 1
        Else
            .MaximumScale = -Int(-maxBattery * 10) / 10
        End If
        .MajorUnitIsAuto = True
        .TickLabels.Orientation = 0
        .DisplayUnit = xlNone
        .HasDisplayUnitLabel = False
        .TickLabelPosition = xlTickLabelPositionHigh
        .TickLabels.NumberFormatLinked = False
        .TickLabels.NumberFormatLocal = "0%"
        .Crosses = xlAxisCrossesMinimum
    End With

    ' Re-fetch both axes after all geometry changes and write titles LAST.
    FBStage = "Umisteni a nazvy os Y"
    graf.Axes(xlCategory, xlPrimary).Crosses = xlAxisCrossesMinimum
    graf.Axes(xlCategory, xlSecondary).Crosses = xlAxisCrossesMaximum
    FBHideSecondaryX graf
    With graf.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.text = "Spot" & ChrW(345) & "eba (MWh)"
    End With
    With graf.Axes(xlValue, xlSecondary)
        .HasTitle = True
        .AxisTitle.text = "Stav baterie (%)"
    End With
End Sub

Private Sub FBHideSecondaryX(ByVal graf As chart)
    If graf.HasAxis(xlCategory, xlSecondary) Then
        With graf.Axes(xlCategory, xlSecondary)
            .TickLabelPosition = xlTickLabelPositionNone
            .MajorTickMark = xlTickMarkNone
            .MinorTickMark = xlTickMarkNone
            .Format.Line.Visible = msoFalse
        End With
    End If
End Sub

Private Sub FBValidateAxes(ByVal graf As chart)
    Dim item As Series
    Dim kind As Long
    ' Postconditions are read back in Excel whenever the button is clicked.
    If graf.Axes(xlCategory, xlPrimary).Crosses <> xlAxisCrossesMinimum Then
        Err.Raise vbObjectError + 5210, , "Spotreba neni umistena vlevo."
    End If
    If graf.Axes(xlCategory, xlSecondary).Crosses <> xlAxisCrossesMaximum Then
        Err.Raise vbObjectError + 5211, , "Baterie neni umistena vpravo."
    End If
    If graf.Axes(xlCategory, xlPrimary).TickLabelPosition <> xlTickLabelPositionLow Then
        Err.Raise vbObjectError + 5212, , "Popisky osy X nejsou dole."
    End If
    If graf.Axes(xlCategory, xlSecondary).TickLabelPosition <> xlTickLabelPositionNone Then
        Err.Raise vbObjectError + 5213, , "Druha osa X neni skryta."
    End If
    If graf.Axes(xlValue, xlPrimary).AxisTitle.text <> "Spot" & ChrW(345) & "eba (MWh)" Then
        Err.Raise vbObjectError + 5214, , "Leva osa Y nema spravny nazev."
    End If
    If graf.Axes(xlValue, xlSecondary).AxisTitle.text <> "Stav baterie (%)" Then
        Err.Raise vbObjectError + 5215, , "Prava osa Y nema spravny nazev."
    End If
    For Each item In graf.SeriesCollection
        kind = FBKind(item.Name)
        If kind = 1 Then
            If item.AxisGroup <> xlSecondary Then Err.Raise vbObjectError + 5216, , "Baterie nema pravou osu Y."
        ElseIf kind = 2 Or kind = 3 Then
            If item.AxisGroup <> xlPrimary Then Err.Raise vbObjectError + 5217, , "Spotreba nema levou osu Y."
        End If
    Next item
End Sub

Private Function FBTargetSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If FBNorm(ws.Name) = "spotr.+bat.+pret." Then
            Set FBTargetSheet = ws
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 5204, , "List spotr. + bat. + pret. nebyl nalezen."
End Function

Private Function FBTargetChart(ByVal ws As Worksheet) As ChartObject
    Dim co As ChartObject
    Dim fallback As ChartObject
    For Each co In ws.ChartObjects
        If co.Name = "Graf 10" Then Set fallback = co
        If co.chart.HasTitle Then
            If FBNorm(co.chart.ChartTitle.text) = "kapacitabaterkyvs.spotreba" Then
                Set FBTargetChart = co
                Exit Function
            End If
        End If
    Next co
    If Not fallback Is Nothing Then
        Set FBTargetChart = fallback
        Exit Function
    End If
    Err.Raise vbObjectError + 5205, , "Graf Kapacita baterky vs. spotreba nebyl nalezen."
End Function

Private Sub FBEnsureButton(ByVal ws As Worksheet, ByVal co As ChartObject)
    Dim button As Shape
    On Error Resume Next
    Set button = ws.Shapes("btnFormatovaniBaterkaBezPretoku")
    On Error GoTo 0
    If button Is Nothing Then
        Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, co.Left + co.Width + 12, co.Top, 150, 32)
        button.Name = "btnFormatovaniBaterkaBezPretoku"
    End If
    With button
        .TextFrame.Characters.text = "Zform" & ChrW(225) & "tovat graf"
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!Module4.FormatovaniBaterkaBezPretokuProcenta"
        .AlternativeText = "FBP_TARGET_BUTTON"
        .Left = co.Left + co.Width + 12
        .Top = co.Top
        .Placement = xlFreeFloating
    End With
End Sub

Private Function FBKind(ByVal caption As String) As Long
    Dim n As String
    n = FBNorm(caption)
    If InStr(n, "aktualnistavbaterkyspretoky") > 0 Or InStr(n, "aktualnistavbateriespretoky") > 0 Then
        FBKind = 1
    ElseIf InStr(n, "spotreba+baterka+pretoky(mwh)") > 0 Then
        FBKind = 2
    ElseIf InStr(n, "spotreba(mwh)") > 0 Then
        FBKind = 3
    End If
End Function

Private Function FBNorm(ByVal text As String) As String
    Dim codes As Variant, letters As Variant, i As Long
    codes = Array(225, 269, 271, 233, 283, 237, 328, 243, 345, 353, 357, 250, 367, 253, 382)
    letters = Array("a", "c", "d", "e", "e", "i", "n", "o", "r", "s", "t", "u", "u", "y", "z")
    text = LCase$(text)
    For i = LBound(codes) To UBound(codes)
        text = Replace(text, ChrW(codes(i)), letters(i))
    Next i
    text = Replace(text, " ", "")
    text = Replace(text, ChrW(160), "")
    text = Replace(text, vbTab, "")
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    FBNorm = Replace(text, "sporeba", "spotreba")
End Function

Private Sub FBMaximum(ByVal values As Variant, ByRef maximum As Double)
    Dim item As Variant
    If IsArray(values) Then
        For Each item In values
            FBMaximum item, maximum
        Next item
    ElseIf Not IsError(values) And Not IsNull(values) And Not IsEmpty(values) Then
        If IsNumeric(values) Then
            If CDbl(values) > maximum Then maximum = CDbl(values)
        End If
    End If
End Sub


' Dedicated formatter for the second chart; original exports procedures above are unchanged.
Public Sub FormatovaniBaterkaBezPretokuNaListu()
    Dim priorUpdating As Boolean
    Dim priorEvents As Boolean
    Dim ws As Worksheet
    Dim co As ChartObject
    Dim graf As chart
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim dataField As PivotField
    Dim item As Series
    Dim pass As Long
    Dim kind As Long
    Dim counts(1 To 3) As Long
    Dim foundConsumption As Boolean
    Dim fieldsChanged As Boolean
    Dim maxBattery As Double
    Dim errText As String
    Dim errCode As Long

    priorUpdating = Application.ScreenUpdating
    priorEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    FBStage = "Vyhledani zdrojoveho grafu"
    Set ws = NBTargetSheet()
    Set co = NBTargetChart(ws)
    Set graf = co.chart
    If graf.PivotLayout Is Nothing Then
        Err.Raise vbObjectError + 5200, , "Cilovy graf neni kontingencni."
    End If
    Set pt = graf.PivotLayout.PivotTable
    FBStage = "Vyber ukazatelu baterie bez pretoku"
    NBEnsureFields pt, fieldsChanged
    If fieldsChanged Then
        If Not pt.RefreshTable Then
            Err.Raise vbObjectError + 5300, , "Kontingencni tabulku se nepodarilo aktualizovat."
        End If
    End If

    ' Types first, axes second: Excel can regroup series after changing a type.
    For Each item In graf.SeriesCollection
        kind = NBKind(item.Name)
        If kind = 1 Then
            If item.ChartType <> xlColumnClustered Then item.ChartType = xlColumnClustered
        ElseIf kind = 2 Or kind = 3 Then
            If item.ChartType <> xlLine Then item.ChartType = xlLine
        End If
    Next item

    For pass = 1 To 2
        For Each item In graf.SeriesCollection
            kind = NBKind(item.Name)
            If kind = 1 Then
                If item.AxisGroup <> xlSecondary Then item.AxisGroup = xlSecondary
            ElseIf kind = 2 Or kind = 3 Then
                If item.AxisGroup <> xlPrimary Then item.AxisGroup = xlPrimary
            End If
        Next item
    Next pass

    For Each item In graf.SeriesCollection
        kind = NBKind(item.Name)
        If kind > 0 Then counts(kind) = counts(kind) + 1
        If kind = 1 Then
            item.Format.Fill.Solid
            item.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
            item.Format.Fill.Transparency = 0.25
            item.Format.Line.Visible = msoFalse
            FBMaximum item.values, maxBattery
        ElseIf kind = 2 Or kind = 3 Then
            item.Format.Line.Visible = msoTrue
            item.Format.Line.Weight = 1.25
            item.MarkerStyle = xlMarkerStyleNone
            item.Smooth = False
            If kind = 2 Then
                item.Format.Line.ForeColor.RGB = RGB(0, 0, 255)
            Else
                item.Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            End If
        End If
    Next item

    If counts(1) = 0 Or counts(2) = 0 Or counts(3) = 0 Then
        Err.Raise vbObjectError + 5203, , "V grafu nebyly rozpoznany vsechny tri ukazatele."
    End If

    ' Redraw any regrouped series first, then restore ALL four axes last.
    graf.Refresh
    FBConfigureAxes graf, maxBattery
    graf.Axes(xlCategory, xlPrimary).TickLabels.Font.Size = 11

    FBStage = "Kontrola vyslednych os"
    FBValidateAxes graf
    NBValidateSeries graf
    FBStage = "Aktualizace tlacitka"
    NBEnsureButton ws, co
    Application.EnableEvents = priorEvents
    Application.ScreenUpdating = priorUpdating
    Exit Sub

Failed:
    errText = Err.Description
    errCode = Err.Number
    On Error Resume Next
    If Not graf Is Nothing Then FBHideSecondaryX graf
    Application.EnableEvents = priorEvents
    Application.ScreenUpdating = priorUpdating
    On Error GoTo 0
    MsgBox "Graf se nepodarilo zformatovat." & vbCrLf & "Krok: " & FBStage & vbCrLf & "Chyba " & errCode & ": " & errText, vbExclamation, "Formatovani grafu"
End Sub


Private Function NBTargetSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If FBNorm(ws.Name) = "spotreba+baterka" Then
            Set NBTargetSheet = ws
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 5301, , "List spotreba + baterka nebyl nalezen."
End Function

Private Function NBTargetChart(ByVal ws As Worksheet) As ChartObject
    Dim co As ChartObject
    Dim fallback As ChartObject
    Dim title As String
    For Each co In ws.ChartObjects
        If co.Name = "Graf 1" Then Set fallback = co
        If co.Chart.HasTitle Then
            title = FBNorm(co.Chart.ChartTitle.Text)
            Do While Right$(title, 1) = "."
                title = Left$(title, Len(title) - 1)
            Loop
            If title = "kapacitabaterkyvs.spotreba" Then
                Set NBTargetChart = co
                Exit Function
            End If
        End If
    Next co
    If Not fallback Is Nothing Then
        Set NBTargetChart = fallback
        Exit Function
    End If
    Err.Raise vbObjectError + 5302, , "Graf Kapacita baterky vs. spotreba nebyl nalezen."
End Function

Private Function NBSourceField(ByVal pt As PivotTable, ByVal key As String) As PivotField
    Dim item As PivotField
    For Each item In pt.PivotFields
        If FBNorm(item.Name) = key Then
            Set NBSourceField = item
            Exit Function
        End If
    Next item
    Err.Raise vbObjectError + 5303, , "Chybi zdrojove pole: " & key
End Function

Private Sub NBEnsureFields(ByVal pt As PivotTable, ByRef fieldsChanged As Boolean)
    Dim keys As Variant
    Dim i As Long
    Dim j As Long
    Dim item As PivotField
    Dim source As PivotField
    Dim name As String
    Dim caption As String
    Dim found As Boolean
    keys = Array("aktualnistavbaterkybezpretoku(%)", "spotreba+baterkabezpretoku(mwh)", "spotreba(mwh)")

    ' Verify all sources before changing the existing field list.
    For i = 0 To 2
        Set source = NBSourceField(pt, CStr(keys(i)))
    Next i
    ' Keep this chart consistently in the WITHOUT-overflows scenario.
    For j = pt.DataFields.Count To 1 Step -1
        Set item = pt.DataFields(j)
        name = FBNorm(item.SourceName)
        If name = "aktualnistavbaterkyspretoky(%)" Or name = "spotreba+baterka+pretoky(mwh)" Then
            item.Orientation = xlHidden
            fieldsChanged = True
        End If
    Next j
    For i = 0 To 2
        found = False
        For Each item In pt.DataFields
            If FBNorm(item.SourceName) = CStr(keys(i)) Then
                found = True
                Exit For
            End If
        Next item
        If Not found Then
            Set source = NBSourceField(pt, CStr(keys(i)))
            caption = "Sou" & ChrW(269) & "et z " & source.Name
            If i = 2 Then caption = "Sou" & ChrW(269) & "et z Spot" & ChrW(345) & "eba (MWh)"
            Set item = pt.AddDataField(source, caption, xlSum)
            If i = 0 Then item.NumberFormat = "0%" Else item.NumberFormat = "0.000000000"
            fieldsChanged = True
        End If
    Next i
    ' Reorder only after all removals/additions; never during a reverse removal loop.
    For i = 0 To 2
        For Each item In pt.DataFields
            If FBNorm(item.SourceName) = CStr(keys(i)) Then
                If item.Position <> i + 1 Then
                    item.Position = i + 1
                    fieldsChanged = True
                End If
                Exit For
            End If
        Next item
    Next i
End Sub

Private Function NBKind(ByVal caption As String) As Long
    Dim name As String
    name = FBNorm(caption)
    If InStr(name, "aktualnistavbaterkybezpretoku") > 0 Or InStr(name, "aktualnistavbateriebezpretoku") > 0 Then
        NBKind = 1
    ElseIf InStr(name, "spotreba+baterkabezpretoku(mwh)") > 0 Then
        NBKind = 2
    ElseIf InStr(name, "spotreba(mwh)") > 0 Then
        NBKind = 3
    End If
End Function

Private Sub NBValidateSeries(ByVal graf As Chart)
    Dim item As Series
    Dim kind As Long
    Dim counts(1 To 3) As Long
    For Each item In graf.SeriesCollection
        kind = NBKind(item.Name)
        If kind > 0 Then counts(kind) = counts(kind) + 1
        If kind = 1 Then
            If item.AxisGroup <> xlSecondary Then Err.Raise vbObjectError + 5304, , "Baterie bez pretoku nema pravou osu Y."
        ElseIf kind = 2 Or kind = 3 Then
            If item.AxisGroup <> xlPrimary Then Err.Raise vbObjectError + 5305, , "Spotreba nema levou osu Y."
        End If
    Next item
    If counts(1) = 0 Or counts(2) = 0 Or counts(3) = 0 Then
        Err.Raise vbObjectError + 5306, , "V grafu chybi nektery ukazatel scenare bez pretoku."
    End If
End Sub

Private Sub NBEnsureButton(ByVal ws As Worksheet, ByVal co As ChartObject)
    Dim button As Shape
    On Error Resume Next
    Set button = ws.Shapes("btnFormatovaniBaterkaBezPretoku")
    On Error GoTo 0
    If button Is Nothing Then
        Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, co.Left + co.Width + 12, co.Top, 150, 32)
        button.Name = "btnFormatovaniBaterkaBezPretoku"
    End If
    With button
        .TextFrame.Characters.Text = "Zform" & ChrW(225) & "tovat graf"
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!Module4.FormatovaniBaterkaBezPretokuNaListu"
        .AlternativeText = "FBP_NO_EXPORT_BUTTON"
        .Left = co.Left + co.Width + 12
        .Top = co.Top
        .Placement = xlFreeFloating
    End With
End Sub
