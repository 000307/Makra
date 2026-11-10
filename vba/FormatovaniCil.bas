Attribute VB_Name = "Module4"
Option Explicit

' Formats the existing PivotChart; does not create another chart.
Public Sub FormatovaniBaterkaBezPretokuProcenta()
    Dim priorUpdating As Boolean
    Dim priorEvents As Boolean
    Dim ws As Worksheet
    Dim co As ChartObject
    Dim graf As Chart
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim dataField As PivotField
    Dim item As Series
    Dim pass As Long
    Dim kind As Long
    Dim counts(1 To 3) As Long
    Dim foundConsumption As Boolean
    Dim maxBattery As Double
    Dim errText As String

    priorUpdating = Application.ScreenUpdating
    priorEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = FBTargetSheet()
    Set co = FBTargetChart(ws)
    Set graf = co.Chart
    If graf.PivotLayout Is Nothing Then
        Err.Raise vbObjectError + 5200, , "Cilovy graf neni kontingencni."
    End If
    Set pt = graf.PivotLayout.PivotTable

    ' Add the original-consumption measure only when it is absent.
    For Each dataField In pt.DataFields
        If FBNorm(dataField.SourceName) = "spotreba(mwh)" Then
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
            FBMaximum item.Values, maxBattery
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

    FBEnsureButton ws, co
    Application.EnableEvents = priorEvents
    Application.ScreenUpdating = priorUpdating
    Exit Sub

Failed:
    errText = Err.Description
    On Error Resume Next
    Application.EnableEvents = priorEvents
    Application.ScreenUpdating = priorUpdating
    On Error GoTo 0
    MsgBox "Graf se nepodarilo zformatovat: " & errText, vbExclamation, "Formatovani grafu"
End Sub


Private Sub FBConfigureAxes(ByVal graf As Chart, ByVal maxBattery As Double)
    Dim decimalSeparator As String
    decimalSeparator = CStr(Application.International(xlDecimalSeparator))
    If Not Application.UseSystemSeparators Then
        decimalSeparator = Application.DecimalSeparator
    End If

    ' A primary/secondary group is not itself a left/right position.
    ' Category-axis crossings determine where the two Y axes are drawn.
    graf.HasAxis(xlCategory, xlPrimary) = True
    graf.HasAxis(xlCategory, xlSecondary) = True
    graf.HasAxis(xlValue, xlPrimary) = True
    graf.HasAxis(xlValue, xlSecondary) = True

    With graf.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Spot" & ChrW(345) & "eba (MWh)"
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

    With graf.Axes(xlValue, xlSecondary)
        .HasTitle = True
        .AxisTitle.Text = "Stav baterie (%)"
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

    With graf.Axes(xlCategory, xlPrimary)
        .ReversePlotOrder = False
        .CategoryType = xlCategoryScale
        .AxisBetweenCategories = True
        .HasMultiLevelLabels = True
        .TickLabelSpacingIsAuto = True
        .TickLabelPosition = xlTickLabelPositionLow
        .MajorTickMark = xlTickMarkOutside
        .MinorTickMark = xlTickMarkNone
        .TickLabels.Orientation = 90
        .Format.Line.Visible = msoTrue
        .Crosses = xlAxisCrossesMinimum
    End With

    With graf.Axes(xlCategory, xlSecondary)
        .ReversePlotOrder = False
        .CategoryType = xlCategoryScale
        .AxisBetweenCategories = True
        .TickLabelPosition = xlTickLabelPositionNone
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .Format.Line.Visible = msoFalse
        .Crosses = xlAxisCrossesMaximum
    End With
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
        If co.Chart.HasTitle Then
            If FBNorm(co.Chart.ChartTitle.Text) = "kapacitabaterkyvs.spotreba" Then
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
        .TextFrame.Characters.Text = "Zform" & ChrW(225) & "tovat graf"
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
    If InStr(n, "aktualnistavbaterkybezpretoku") > 0 Or InStr(n, "aktualnistavbateriebezpretoku") > 0 Then
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
