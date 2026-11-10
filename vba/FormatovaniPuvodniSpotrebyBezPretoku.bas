Attribute VB_Name = "Module1"
Option Explicit

' =========================================================
' 1. RUCNI FORMATOVANI OZNACENEHO GRAFU
' =========================================================
Public Sub ZformatovatGrafBezPretoku()
    Dim graf As Chart
    On Error GoTo Chyba
    Set graf = NajdiPuvodniSpotrebuBezPretoku()
    FormatujGrafBezPretoku graf
    Exit Sub
Chyba:
    MsgBox "Nelze zformatovat graf Puvodni spotreba: " & Err.Description, vbExclamation
End Sub

' =========================================================
' 2. FORMATOVANI POMOCI TLACITKA
' =========================================================
Public Sub SpustitFormatovaniBezPretoku()
    Dim graf As Chart
    On Error GoTo Chyba
    Set graf = NajdiPuvodniSpotrebuBezPretoku()
    FormatujGrafBezPretoku graf
    Exit Sub
Chyba:
    MsgBox "Nelze zformatovat graf Puvodni spotreba: " & Err.Description, vbExclamation
End Sub

' =========================================================
' 3. HLAVNI FORMATOVACI MAKRO
' =========================================================
Private Sub FormatujGrafBezPretoku(ByVal graf As chart)
    Dim rada As Series
    Dim osa As Axis
    Dim radySpot As Collection
    Dim radyCervene As Collection
    Dim radyModre As Collection
    Dim radyHnede As Collection
    Dim nazev As String
    Dim i As Long
    Dim pocetOs As Long
    Dim pocetNazvu As Long
    Dim puvodniZobrazeni As Boolean
    Dim nulySjednoceny As Boolean
    Dim puvodniUdalosti As Boolean

    puvodniZobrazeni = Application.ScreenUpdating
    puvodniUdalosti = Application.EnableEvents
    On Error GoTo Chyba
    If Not JeCilovyGrafBezPretoku(graf) Then Err.Raise vbObjectError + 5401, , "Toto makro smi upravit jen graf Puvodni spotreba na listu spotreba + baterka."
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set radySpot = New Collection
    Set radyCervene = New Collection
    Set radyModre = New Collection
    Set radyHnede = New Collection

    ' Trideni vsech rad (vcetne jednotlivych tydnu/let).
    For Each rada In graf.SeriesCollection
        nazev = LCase$(rada.name)

        If InStr(NormCilBezPretoku(nazev), "cenyspot") > 0 Or InStr(NormCilBezPretoku(nazev), "cenaspot") > 0 Then
            radySpot.Add rada

        ElseIf JeBaterkaBezPretoku(nazev) Then
            radyModre.Add rada

        ElseIf JePretokDoSite(nazev) Then
            radyHnede.Add rada

        ElseIf InStr(1, nazev, "sou", vbTextCompare) > 0 _
            And InStr(1, nazev, "spo", vbTextCompare) > 0 _
            And InStr(1, nazev, "(mwh)", vbTextCompare) > 0 Then
            radyCervene.Add rada
        End If
    Next rada

    If radySpot.Count = 0 Or radyCervene.Count = 0 Or radyModre.Count = 0 Then
        Err.Raise vbObjectError + 5402, , "V cilovem grafu musi byt cena SPOT a obe spotreby. Zadny jiny graf nebyl zmenen."
    End If

    ' Ceny SPOT - zelena, prava osa Y.
    For i = 1 To radySpot.Count
        Set rada = radySpot(i)
        rada.AxisGroup = xlSecondary
        NastavBarvuARozmerCary rada, RGB(0, 220, 0), 1
    Next i

    graf.HasAxis(xlValue, xlSecondary) = True

    ' Puvodni spotreba - cervena, leva osa Y.
    For i = 1 To radyCervene.Count
        Set rada = radyCervene(i)
        rada.AxisGroup = xlPrimary
        NastavBarvuARozmerCary rada, RGB(230, 0, 0), 1
    Next i

    ' Spotreba + baterka BEZ pretoku - modra, leva osa Y.
    For i = 1 To radyModre.Count
        Set rada = radyModre(i)
        rada.AxisGroup = xlPrimary
        NastavBarvuARozmerCary rada, RGB(0, 0, 230), 1
    Next i

    ' Pretoky do site (pokud jsou v grafu) - hneda #9C6E30.
    For i = 1 To radyHnede.Count
        Set rada = radyHnede(i)
        rada.AxisGroup = xlPrimary
        NastavBarvuARozmerCary rada, RGB(156, 110, 48), 1
    Next i

    graf.Refresh
    PripravOsyPuvodniSpotreby graf

    ' Nuly obou hodnotovych os na jedne vodorovne urovni.
    nulySjednoceny = False
    If radySpot.Count > 0 Then
        SjednotNulyOs graf
        nulySjednoceny = True
    End If

    ' Write titles only after all geometry/range changes.
    graf.Axes(xlValue, xlPrimary).HasTitle = True
    graf.Axes(xlValue, xlPrimary).AxisTitle.Text = "Spot" & ChrW(345) & "eba (MWh)"
    graf.Axes(xlValue, xlSecondary).HasTitle = True
    graf.Axes(xlValue, xlSecondary).AxisTitle.Text = "Cena SPOT (K" & ChrW(269) & "/MWh)"
    ProverOsyPuvodniSpotreby graf, radySpot, radyCervene, radyModre

    ' Popisky a nazvy vsech os.
    pocetOs = 0
    pocetNazvu = 0
    For Each osa In graf.Axes
        With osa.TickLabels.Font
            .Color = RGB(0, 0, 0)
            .Size = 10.5
        End With
        pocetOs = pocetOs + 1

        If osa.HasTitle Then
            With osa.AxisTitle.Font
                .Color = RGB(0, 0, 0)
                .Size = 10.5
                .Bold = True
            End With
            pocetNazvu = pocetNazvu + 1
        End If
    Next osa

    Application.EnableEvents = puvodniUdalosti
    Application.ScreenUpdating = puvodniZobrazeni

    MsgBox "Graf BEZ PRETOKU byl naformatovan!" & vbCrLf & vbCrLf & _
           "Zelene rady SPOT: " & radySpot.Count & vbCrLf & _
           "Cervene rady spotreby: " & radyCervene.Count & vbCrLf & _
           "Modre rady bez pretoku: " & radyModre.Count & vbCrLf & _
           "Hnede rady pretoku do site: " & radyHnede.Count & vbCrLf & vbCrLf & _
           "Zarovnani nul obou os: " & IIf(nulySjednoceny, "ANO", "NE - chybi SPOT") & vbCrLf & _
           "Upravene osy: " & pocetOs & vbCrLf & _
           "Upravene nazvy os: " & pocetNazvu, _
           vbInformation, "Formatovani BEZ pretoku"
    Exit Sub

Chyba:
    Application.EnableEvents = puvodniUdalosti
    Application.ScreenUpdating = puvodniZobrazeni
    MsgBox "Chyba pri formatovani grafu: " & Err.Description, vbExclamation
End Sub

' =========================================================
' 4. POMOCNE PROCEDURY
' =========================================================
Private Sub NastavBarvuARozmerCary(ByVal rada As Series, _
                                  ByVal barva As Long, _
                                  ByVal tloustka As Single)
    With rada.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = barva
        .DashStyle = msoLineSolid
        .Weight = tloustka
        .Transparency = 0
    End With
    rada.MarkerStyle = xlMarkerStyleNone
End Sub

Private Function JeBaterkaBezPretoku(ByVal nazev As String) As Boolean
    JeBaterkaBezPretoku = InStr(NormCilBezPretoku(nazev), "spotreba+baterkabezpretoku(mwh)") > 0
End Function

Private Function JePretokDoSite(ByVal nazev As String) As Boolean
    JePretokDoSite = _
        (InStr(1, nazev, "pretok", vbTextCompare) > 0 Or _
         InStr(1, nazev, "přetok", vbTextCompare) > 0) And _
        (InStr(1, nazev, "site", vbTextCompare) > 0 Or _
         InStr(1, nazev, "sítě", vbTextCompare) > 0)
End Function

' Vypocte podil vzdalenosti nuly od spodniho okraje (0-1),
' a pro obe osy zvoli rozsahy, ktere vsechny hodnoty zachovaji.
Private Sub SjednotNulyOs(ByVal graf As chart)
    Dim osaLeva As Axis
    Dim osaPrava As Axis
    Dim minL As Double, maxL As Double
    Dim minP As Double, maxP As Double
    Dim podil As Double

    Set osaLeva = graf.Axes(xlValue, xlPrimary)
    Set osaPrava = graf.Axes(xlValue, xlSecondary)

    ' Nejdrive znovu nacti skutecne rozsahy po zmene filtru.
    osaLeva.MinimumScaleIsAuto = True
    osaLeva.MaximumScaleIsAuto = True
    osaPrava.MinimumScaleIsAuto = True
    osaPrava.MaximumScaleIsAuto = True

    minL = osaLeva.MinimumScale
    maxL = osaLeva.MaximumScale
    minP = osaPrava.MinimumScale
    maxP = osaPrava.MaximumScale

    If maxL <= minL Or maxP <= minP Then
        Err.Raise vbObjectError + 521, , "Osy nemaji platny rozsah."
    End If

    If minL < 0 And maxL > 0 Then
        ' Pokud leva osa obsahuje obe znamenka, zachovej jeji pozici nuly.
        podil = -minL / (maxL - minL)
    ElseIf minP < 0 And maxP > 0 Then
        ' Jinak zachovej pozici nuly na prave ose, pokud ma obe znamenka.
        podil = -minP / (maxP - minP)
    ElseIf minL >= 0 And minP >= 0 Then
        ' Obe osy jen kladne: nula dole.
        podil = 0
    ElseIf maxL <= 0 And maxP <= 0 Then
        ' Obe osy jen zaporne: nula nahore.
        podil = 1
    Else
        ' Jedna osa kladna a druha zaporna: nula uprostred.
        podil = 0.5
    End If

    NastavOsuSeStejnouNulou osaLeva, minL, maxL, podil
    NastavOsuSeStejnouNulou osaPrava, minP, maxP, podil
End Sub

Private Sub NastavOsuSeStejnouNulou(ByVal osa As Axis, _
                                    ByVal minHod As Double, _
                                    ByVal maxHod As Double, _
                                    ByVal podil As Double)
    Dim rozsah As Double
    Dim kladnyRozsah As Double
    Dim zapornyRozsah As Double

    If podil <= 0 Then
        ' Nula na spodnim okraji.
        osa.MinimumScale = 0
        osa.MaximumScale = IIf(maxHod > 0, maxHod, 1)

    ElseIf podil >= 1 Then
        ' Nula na hornim okraji.
        osa.MinimumScale = IIf(minHod < 0, minHod, -1)
        osa.MaximumScale = 0

    Else
        ' Pro jednotnou vysku nuly plati:
        ' -minimum / (maximum - minimum) = podil.
        kladnyRozsah = 0
        zapornyRozsah = 0

        If maxHod > 0 Then kladnyRozsah = maxHod / (1 - podil)
        If minHod < 0 Then zapornyRozsah = -minHod / podil

        If kladnyRozsah > zapornyRozsah Then
            rozsah = kladnyRozsah
        Else
            rozsah = zapornyRozsah
        End If

        If rozsah <= 0 Then rozsah = 1
        rozsah = rozsah * 1.02 ' Mala rezerva na okrajich.

        ' Nastaveni v tomto poradi vychazi i s mezitimnim auto-prizpusobenim.
        osa.MinimumScale = -podil * rozsah
        osa.MaximumScale = (1 - podil) * rozsah
    End If
End Sub

' =========================================================
' 5. VYTVORENI / AKTUALIZACE TLACITKA
' =========================================================
Public Sub VytvorTlacitkoBezPretoku()
    Dim grafObj As ChartObject
    Dim graf As Chart
    Dim ws As Worksheet
    Dim tlacitko As Shape

    On Error GoTo Chyba

    Set graf = NajdiPuvodniSpotrebuBezPretoku()
    Set grafObj = graf.Parent
    Set ws = grafObj.Parent

    On Error Resume Next
    Set tlacitko = ws.Shapes("btnFormatBezPretoku")
    On Error GoTo Chyba

    If tlacitko Is Nothing Then
        Set tlacitko = ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            grafObj.Left + grafObj.Width + 15, _
            grafObj.Top + 20, 160, 40)
        tlacitko.name = "btnFormatBezPretoku"
    End If

    With tlacitko
        .Left = grafObj.Left + grafObj.Width + 15
        .Top = grafObj.Top + 20
        .Width = 160
        .height = 40
        .TextFrame.Characters.text = "Zformátovat graf"
        .Fill.ForeColor.RGB = RGB(0, 102, 204)
        .Line.Visible = msoFalse

        With .TextFrame.Characters.Font
            .name = "Calibri"
            .Size = 11
            .Bold = True
            .Color = RGB(255, 255, 255)
        End With

        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter

        ' Na tlacitko uloz nazev prirazeneho grafu.
        .AlternativeText = grafObj.name
        .OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & _
                    "'!Module1.SpustitFormatovaniBezPretoku"
        .Placement = xlFreeFloating
    End With

    MsgBox "Tlacitko pro graf BEZ PRETOKU bylo vytvoreno / aktualizovano.", vbInformation
    Exit Sub

Chyba:
    MsgBox "Chyba pri vytvareni tlacitka: " & Err.Description, vbExclamation
End Sub



Private Function NormCilBezPretoku(ByVal value As String) As String
    Dim codes As Variant, letters As Variant, i As Long
    codes = Array(225, 269, 271, 233, 283, 237, 328, 243, 345, 353, 357, 250, 367, 253, 382)
    letters = Array("a", "c", "d", "e", "e", "i", "n", "o", "r", "s", "t", "u", "u", "y", "z")
    value = LCase$(value)
    For i = LBound(codes) To UBound(codes)
        value = Replace(value, ChrW(codes(i)), letters(i))
    Next i
    value = Replace(value, " ", "")
    value = Replace(value, ChrW(160), "")
    value = Replace(value, vbTab, "")
    value = Replace(value, vbCr, "")
    value = Replace(value, vbLf, "")
    value = Replace(value, ".", "")
    NormCilBezPretoku = Replace(value, "sporeba", "spotreba")
End Function

Private Function JeCilovyGrafBezPretoku(ByVal graf As Chart) As Boolean
    Dim obj As ChartObject
    Dim ws As Worksheet
    If Not TypeOf graf.Parent Is ChartObject Then Exit Function
    Set obj = graf.Parent
    Set ws = obj.Parent
    If NormCilBezPretoku(ws.Name) <> "spotreba+baterka" Then Exit Function
    If Not graf.HasTitle Then Exit Function
    JeCilovyGrafBezPretoku = NormCilBezPretoku(graf.ChartTitle.Text) = "puvodnispotrebavsspotreba+na/vybijeni"
End Function

Private Function NajdiPuvodniSpotrebuBezPretoku() As Chart
    Dim ws As Worksheet
    Dim item As ChartObject
    Dim found As Chart
    Dim count As Long
    For Each ws In ThisWorkbook.Worksheets
        If NormCilBezPretoku(ws.Name) = "spotreba+baterka" Then
            For Each item In ws.ChartObjects
                If JeCilovyGrafBezPretoku(item.Chart) Then
                    Set found = item.Chart
                    count = count + 1
                End If
            Next item
        End If
    Next ws
    If count <> 1 Then Err.Raise vbObjectError + 5403, , "Cilovy graf nebyl jednoznacne nalezen; ostatni grafy nebyly zmeneny."
    Set NajdiPuvodniSpotrebuBezPretoku = found
End Function

Private Sub PripravOsyPuvodniSpotreby(ByVal graf As Chart)
    Dim decimalSeparator As String
    decimalSeparator = CStr(Application.International(xlDecimalSeparator))
    If Not Application.UseSystemSeparators Then decimalSeparator = Application.DecimalSeparator
    graf.HasAxis(xlCategory, xlPrimary) = True
    graf.HasAxis(xlCategory, xlSecondary) = True
    graf.HasAxis(xlValue, xlPrimary) = True
    graf.HasAxis(xlValue, xlSecondary) = True
    With graf.Axes(xlCategory, xlSecondary)
        .TickLabelPosition = xlTickLabelPositionNone
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .Format.Line.Visible = msoFalse
        .ReversePlotOrder = False
        .CategoryType = xlCategoryScale
        .AxisBetweenCategories = True
        .Crosses = xlAxisCrossesMaximum
    End With
    With graf.Axes(xlCategory, xlPrimary)
        .ReversePlotOrder = False
        .CategoryType = xlCategoryScale
        .AxisBetweenCategories = True
        .TickLabels.MultiLevel = True
        .TickLabels.Orientation = xlUpward
        .TickLabelSpacingIsAuto = True
        .TickLabelPosition = xlTickLabelPositionLow
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .Format.Line.Visible = msoTrue
        .Crosses = xlAxisCrossesMinimum
    End With
    With graf.Axes(xlValue, xlPrimary)
        .ReversePlotOrder = False
        .DisplayUnit = xlNone
        .HasDisplayUnitLabel = False
        .TickLabelPosition = xlTickLabelPositionLow
        .TickLabels.NumberFormatLinked = False
        .TickLabels.NumberFormatLocal = "0" & decimalSeparator & "0000"
        .Crosses = xlAxisCrossesMinimum
    End With
    With graf.Axes(xlValue, xlSecondary)
        .ReversePlotOrder = False
        .DisplayUnit = xlNone
        .HasDisplayUnitLabel = False
        .TickLabelPosition = xlTickLabelPositionHigh
        .TickLabels.NumberFormatLinked = False
        .TickLabels.NumberFormatLocal = "0" & decimalSeparator & "00"
        .Crosses = xlAxisCrossesMinimum
    End With
End Sub

Private Sub ProverOsyPuvodniSpotreby(ByVal graf As Chart, ByVal spot As Collection, ByVal cervene As Collection, ByVal modre As Collection)
    Dim rada As Series
    If Not graf.HasAxis(xlValue, xlSecondary) Then Err.Raise vbObjectError + 5404, , "Chybi prava osa ceny SPOT."
    If graf.Axes(xlCategory, xlPrimary).Crosses <> xlAxisCrossesMinimum Then Err.Raise vbObjectError + 5405, , "Spotreba neni umistena vlevo."
    If graf.Axes(xlCategory, xlSecondary).Crosses <> xlAxisCrossesMaximum Then Err.Raise vbObjectError + 5406, , "Cena SPOT neni umistena vpravo."
    For Each rada In spot
        If rada.AxisGroup <> xlSecondary Then Err.Raise vbObjectError + 5407, , "Rada SPOT nema pravou osu."
    Next rada
    For Each rada In cervene
        If rada.AxisGroup <> xlPrimary Then Err.Raise vbObjectError + 5408, , "Spotreba nema levou osu."
    Next rada
    For Each rada In modre
        If rada.AxisGroup <> xlPrimary Then Err.Raise vbObjectError + 5409, , "Spotreba s baterii nema levou osu."
    Next rada
End Sub
