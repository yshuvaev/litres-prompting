Attribute VB_Name = "DashboardModule"
Option Explicit
Private Type ANGroup
    week As String
    region As String
    book As Long
    sold As Double
    gross As Double
    returned As Double
    amount As Double
End Type
Private Type ANBook
    id As String
    title As String
    sold As Double
    returned As Double
    amount As Double
    reasonUnits(1 To 5) As Double
    reasonAmount(1 To 5) As Double
End Type

Private Function LVNum(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    LVNum = (VarType(v) = vbDouble Or VarType(v) = vbInteger Or VarType(v) = vbLong Or VarType(v) = vbCurrency)
End Function
Private Function LVRead(ByVal name As String, ByVal headers As Variant) As Variant
    Dim s As Worksheet, c As Long, n As Long
    Set s = ThisWorkbook.Worksheets(name)
    For c = 0 To UBound(headers)
        If IsError(s.Cells(1, c + 1).Value) Then Err.Raise 5, , "Invalid header: " & name
        If s.Cells(1, c + 1).Value <> headers(c) Then Err.Raise 5, , "Wrong headers: " & name
        n = Application.Max(n, s.Cells(s.Rows.Count, c + 1).End(xlUp).Row)
    Next c
    If n > 100001 Then Err.Raise 5, , "Training limit: 100000 rows per source"
    LVRead = s.Range(s.Cells(1, 1), s.Cells(n, UBound(headers) + 1)).Value2
End Function
Private Function LVOut(ByVal name As String) As Worksheet
    Dim s As Worksheet
    Select Case name
        Case "OUT_Checked", "OUT_Issues", "OUT_ModelAudit", "OUT_Dashboard", "OUT_Regions", "OUT_Books", "OUT_Reasons", "OUT_Partner", "OUT_Rejected"
        Case Else: Err.Raise 5, , "Output is not reserved"
    End Select
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If s Is Nothing Then
        Set s = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        s.name = name
    End If
    s.UsedRange.ClearContents
    s.UsedRange.Interior.Pattern = xlNone
    Do While s.ChartObjects.Count > 0
        s.ChartObjects(1).Delete
    Loop
    Set LVOut = s
End Function
Private Sub LVWrite(ByVal s As Worksheet, ByVal a As Variant)
    Dim i As Long, j As Long
    For i = 1 To UBound(a, 1)
        For j = 1 To UBound(a, 2)
            If VarType(a(i, j)) = vbString Then
                If Left$(a(i, j), 1) = "=" Then a(i, j) = "'" & a(i, j)
            End If
        Next j
    Next i
    s.Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value2 = a
    s.Rows(1).Font.Bold = True
    s.Columns.AutoFit
End Sub
Private Sub LVLog(ByVal name As String, ByVal headers As Variant, ByVal items As Collection)
    Dim a() As Variant, i As Long, c As Long, r As Variant, s As Worksheet
    ReDim a(1 To items.Count + 1, 1 To UBound(headers) + 1)
    For c = 0 To UBound(headers): a(1, c + 1) = headers(c): Next c
    i = 2
    For Each r In items
        For c = 0 To UBound(headers): a(i, c + 1) = r(c): Next c
        i = i + 1
    Next r
    Set s = LVOut(name): LVWrite s, a
End Sub
Private Function LVSeen(ByVal seen As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    Err.Clear
    seen.Add True, "k" & key
    LVSeen = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
End Function

' Paste this complete module once. Run either of the two public procedures.
Private Function ANFind(ByVal index As Collection, ByVal key As String) As Long
    On Error Resume Next
    ANFind = index("k" & key)
    On Error GoTo 0
End Function
Private Function ANID(ByVal v As Variant) As String
    Dim s As String, i As Long, c As String
    If VarType(v) <> vbString Then Err.Raise 5, , "Expected text ID"
    s = CStr(v)
    If Len(s) = 0 Then Err.Raise 5, , "Empty ID"
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-", c, vbBinaryCompare) = 0 Then Err.Raise 5, , "Expected uppercase ID: " & s
    Next i
    ANID = s
End Function
Private Function ANWeek(ByVal v As Variant) As String
    Dim s As String, n As Long
    If VarType(v) <> vbString Then Err.Raise 5, , "Expected text week"
    s = CStr(v)
    If Not s Like "2026-W##" Then Err.Raise 5, , "Expected week 2026-W27..W39"
    n = CLng(Right$(s, 2))
    If n < 27 Or n > 39 Then Err.Raise 5, , "Expected week 2026-W27..W39"
    ANWeek = s
End Function
Private Function ANNumber(ByVal v As Variant, ByVal whole As Boolean) As Double
    If Not LVNum(v) Then Err.Raise 5, , "Expected numeric value, not text"
    If CDbl(v) < 0 Then Err.Raise 5, , "Negative value"
    If whole Then
        If CDbl(v) <> Fix(CDbl(v)) Then Err.Raise 5, , "Expected integer units"
    End If
    ANNumber = CDbl(v)
End Function
Private Function ANText(ByVal v As Variant) As String
    If VarType(v) <> vbString Then Err.Raise 5, , "Expected text"
    If Len(Trim$(CStr(v))) = 0 Then Err.Raise 5, , "Empty text"
    ANText = CStr(v)
End Function
Private Function ANNonempty(ByRef a As Variant, ByVal row As Long) As Boolean
    Dim c As Long
    For c = 1 To UBound(a, 2)
        If IsError(a(row, c)) Then ANNonempty = True: Exit Function
        If Not IsEmpty(a(row, c)) Then
            If CStr(a(row, c)) <> "" Then ANNonempty = True: Exit Function
        End If
    Next c
End Function
Private Function ANReasons() As Variant
    ANReasons = Array("Повреждение", "Ошибка комплектации", "Отказ покупателя", "Излишек запаса", "Полиграфический брак")
End Function
Private Sub ANRead(ByRef weekly As Variant, ByRef regional As Variant, ByRef ranked As Variant, ByRef reasonsOut As Variant)
    On Error GoTo Failed
    Dim sales As Variant, returns As Variant, outlets As Variant, reasonNames As Variant
    Dim byID As New Collection, byCode As New Collection, groupIndex As New Collection, bookIndex As New Collection
    Dim weekIndex As New Collection, regionIndex As New Collection
    Dim g() As ANGroup, b() As ANBook, wg() As Double, rg() As Double, wn() As String, rn() As String, rw() As String
    Dim i As Long, j As Long, k As Long, oi As Long, gi As Long, bi As Long, ri As Long, wi As Long, cap As Long
    Dim reasonOrder(1 To 5) As Long, pos As Long, tempReason As Long
    Dim ng As Long, nb As Long, nw As Long, nr As Long, where As String, ky As String, week As String, returnWeek As String
    Dim outlet As String, point As String, bid As String, title As String, qty As Double, amount As Double, reason As String
    sales = LVRead("Sales", Array("ShipmentWeek", "OutletID", "BookID", "Book", "Units", "GrossAmount"))
    returns = LVRead("Returns", Array("ShipmentWeek", "ReturnWeek", "ReturnPointCode", "BookID", "Units", "ReturnAmount", "Reason"))
    outlets = LVRead("Outlets", Array("OutletID", "ReturnPointCode", "Region"))
    cap = UBound(sales, 1)
    ReDim g(1 To cap): ReDim b(1 To cap)
    ReDim wg(1 To cap, 1 To 2): ReDim rg(1 To cap, 1 To 2)
    ReDim wn(1 To cap): ReDim rn(1 To cap): ReDim rw(1 To cap)
    reasonNames = ANReasons()
    For i = 2 To UBound(outlets, 1)
        where = "Outlets row " & i
        If ANNonempty(outlets, i) Then
            outlet = ANID(outlets(i, 1)): point = ANID(outlets(i, 2)): title = ANText(outlets(i, 3))
            If ANFind(byID, outlet) > 0 Or ANFind(byCode, point) > 0 Then Err.Raise 5, , "Duplicate outlet key"
            byID.Add i, "k" & outlet: byCode.Add i, "k" & point
        End If
    Next i
    For i = 2 To UBound(sales, 1)
        where = "Sales row " & i
        If ANNonempty(sales, i) Then
            week = ANWeek(sales(i, 1)): outlet = ANID(sales(i, 2)): bid = ANID(sales(i, 3)): title = ANText(sales(i, 4))
            qty = ANNumber(sales(i, 5), True): amount = ANNumber(sales(i, 6), False)
            oi = ANFind(byID, outlet)
            If oi = 0 Then Err.Raise 5, , "Unknown OutletID"
            bi = ANFind(bookIndex, bid)
            If bi = 0 Then
                nb = nb + 1: bi = nb: bookIndex.Add bi, "k" & bid: b(bi).id = bid: b(bi).title = title
            ElseIf StrComp(b(bi).title, title, vbBinaryCompare) <> 0 Then
                Err.Raise 5, , "BookID has conflicting names"
            End If
            b(bi).sold = b(bi).sold + qty
            ky = week & "|" & outlet & "|" & bid: gi = ANFind(groupIndex, ky)
            If gi = 0 Then
                ng = ng + 1: gi = ng: groupIndex.Add gi, "k" & ky
                g(gi).week = week: g(gi).region = outlets(oi, 3): g(gi).book = bi
            End If
            g(gi).sold = g(gi).sold + qty: g(gi).gross = g(gi).gross + amount
        End If
    Next i
    For i = 2 To UBound(returns, 1)
        where = "Returns row " & i
        If ANNonempty(returns, i) Then
            week = ANWeek(returns(i, 1)): returnWeek = ANWeek(returns(i, 2)): point = ANID(returns(i, 3)): bid = ANID(returns(i, 4))
            qty = ANNumber(returns(i, 5), True): amount = ANNumber(returns(i, 6), False): reason = ANText(returns(i, 7))
            If returnWeek < week Then Err.Raise 5, , "Return precedes shipment"
            oi = ANFind(byCode, point)
            If oi = 0 Then Err.Raise 5, , "Unknown ReturnPointCode"
            ri = 0
            For j = 0 To 4
                If StrComp(reason, CStr(reasonNames(j)), vbBinaryCompare) = 0 Then ri = j + 1
            Next j
            If ri = 0 Then Err.Raise 5, , "Unknown return reason"
            ky = week & "|" & CStr(outlets(oi, 1)) & "|" & bid: gi = ANFind(groupIndex, ky)
            If gi = 0 Then Err.Raise 5, , "No corresponding shipment group"
            g(gi).returned = g(gi).returned + qty: g(gi).amount = g(gi).amount + amount
            bi = g(gi).book: b(bi).returned = b(bi).returned + qty: b(bi).amount = b(bi).amount + amount
            b(bi).reasonUnits(ri) = b(bi).reasonUnits(ri) + qty: b(bi).reasonAmount(ri) = b(bi).reasonAmount(ri) + amount
        End If
    Next i
    For gi = 1 To ng
        where = g(gi).week & "/" & g(gi).region & "/" & b(g(gi).book).id
        If g(gi).returned > g(gi).sold Or g(gi).amount > g(gi).gross + 0.005 Then Err.Raise 5, , "Returns exceed shipments"
        wi = ANFind(weekIndex, g(gi).week)
        If wi = 0 Then
            nw = nw + 1: wi = nw: weekIndex.Add wi, "k" & g(gi).week: wn(wi) = g(gi).week
        End If
        ky = g(gi).week & "|" & ANKeyText(g(gi).region): ri = ANFind(regionIndex, ky)
        If ri = 0 Then
            nr = nr + 1: ri = nr: regionIndex.Add ri, "k" & ky: rw(ri) = g(gi).week: rn(ri) = g(gi).region
        End If
        wg(wi, 1) = wg(wi, 1) + g(gi).gross: wg(wi, 2) = wg(wi, 2) + g(gi).amount
        rg(ri, 1) = rg(ri, 1) + g(gi).gross: rg(ri, 2) = rg(ri, 2) + g(gi).amount
    Next gi
    ReDim weekly(1 To nw + 1, 1 To 4): ReDim regional(1 To nr + 1, 1 To 5): ReDim ranked(1 To nb + 1, 1 To 6): ReDim reasonsOut(1 To nb * 5 + 1, 1 To 5)
    ANHeader weekly, Array("ShipmentWeek", "GrossAmount", "ReturnAmount", "NetAmount")
    ANHeader regional, Array("ShipmentWeek", "Region", "GrossAmount", "ReturnAmount", "NetAmount")
    ANHeader ranked, Array("BookID", "Book", "ShippedUnits", "ReturnedUnits", "ReturnRate", "ReturnAmount")
    ANHeader reasonsOut, Array("BookID", "Book", "Reason", "ReturnedUnits", "ReturnAmount")
    For i = 1 To nw
        weekly(i + 1, 1) = wn(i): weekly(i + 1, 2) = Application.Round(wg(i, 1), 2): weekly(i + 1, 3) = Application.Round(wg(i, 2), 2): weekly(i + 1, 4) = Application.Round(wg(i, 1) - wg(i, 2), 2)
    Next i
    For i = 1 To nr
        regional(i + 1, 1) = rw(i): regional(i + 1, 2) = rn(i): regional(i + 1, 3) = Application.Round(rg(i, 1), 2): regional(i + 1, 4) = Application.Round(rg(i, 2), 2): regional(i + 1, 5) = Application.Round(rg(i, 1) - rg(i, 2), 2)
    Next i
    For i = 1 To nb
        ranked(i + 1, 1) = b(i).id: ranked(i + 1, 2) = b(i).title: ranked(i + 1, 3) = b(i).sold: ranked(i + 1, 4) = b(i).returned: ranked(i + 1, 5) = 0: ranked(i + 1, 6) = Application.Round(b(i).amount, 2)
        If b(i).sold > 0 Then ranked(i + 1, 5) = b(i).returned / b(i).sold
    Next i
    If nw > 1 Then ANSort weekly, 2, nw + 1, 1
    If nr > 1 Then ANSort regional, 2, nr + 1, 2
    If nb > 1 Then ANSort ranked, 2, nb + 1, 3
    k = 2
    For i = 2 To nb + 1
        bi = ANFind(bookIndex, CStr(ranked(i, 1)))
        For j = 1 To 5: reasonOrder(j) = j: Next j
        For j = 2 To 5
            pos = j
            Do While pos > 1
                If b(bi).reasonUnits(reasonOrder(pos - 1)) >= b(bi).reasonUnits(reasonOrder(pos)) Then Exit Do
                tempReason = reasonOrder(pos): reasonOrder(pos) = reasonOrder(pos - 1): reasonOrder(pos - 1) = tempReason: pos = pos - 1
            Loop
        Next j
        For j = 1 To 5
            ri = reasonOrder(j)
            reasonsOut(k, 1) = b(bi).id: reasonsOut(k, 2) = b(bi).title: reasonsOut(k, 3) = reasonNames(ri - 1): reasonsOut(k, 4) = b(bi).reasonUnits(ri): reasonsOut(k, 5) = Application.Round(b(bi).reasonAmount(ri), 2): k = k + 1
        Next j
    Next i
    Exit Sub
Failed:
    Err.Raise 5, , where & ": " & Err.Description
End Sub
Private Sub ANHeader(ByRef a As Variant, ByVal names As Variant)
    Dim c As Long
    For c = 0 To UBound(names): a(1, c + 1) = names(c): Next c
End Sub
Private Function ANCompare(ByRef a As Variant, ByVal row As Long, ByRef pivot As Variant, ByVal kind As Long) As Long
    Dim cmp As Long
    If kind = 3 Then
        If a(row, 5) > pivot(5) Then ANCompare = -1: Exit Function
        If a(row, 5) < pivot(5) Then ANCompare = 1: Exit Function
        If a(row, 4) > pivot(4) Then ANCompare = -1: Exit Function
        If a(row, 4) < pivot(4) Then ANCompare = 1: Exit Function
    End If
    cmp = StrComp(CStr(a(row, 1)), CStr(pivot(1)), vbBinaryCompare)
    If cmp = 0 And kind = 2 Then cmp = StrComp(CStr(a(row, 2)), CStr(pivot(2)), vbBinaryCompare)
    ANCompare = cmp
End Function
Private Sub ANSort(ByRef a As Variant, ByVal lo As Long, ByVal hi As Long, ByVal kind As Long)
    Dim i As Long, j As Long, c As Long, mid As Long, temp As Variant, pivot As Variant
    ReDim pivot(1 To UBound(a, 2))
    mid = (lo + hi) \ 2
    For c = 1 To UBound(a, 2): pivot(c) = a(mid, c): Next c
    i = lo: j = hi
    Do While i <= j
        Do While ANCompare(a, i, pivot, kind) < 0: i = i + 1: Loop
        Do While ANCompare(a, j, pivot, kind) > 0: j = j - 1: Loop
        If i <= j Then
            For c = 1 To UBound(a, 2): temp = a(i, c): a(i, c) = a(j, c): a(j, c) = temp: Next c
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then ANSort a, lo, j, kind
    If i < hi Then ANSort a, i, hi, kind
End Sub
Public Sub BuildDashboard()
    On Error GoTo Failed
    Dim weekly As Variant, regional As Variant, ranked As Variant, reasons As Variant
    Dim s As Worksheet, t As Worksheet, co As ChartObject, n As Long
    ANRead weekly, regional, ranked, reasons
    Set s = LVOut("OUT_Dashboard"): LVWrite s, weekly
    Set t = LVOut("OUT_Regions"): LVWrite t, regional
    n = UBound(weekly, 1)
    If n < 2 Then Exit Sub
    s.Range("B2:D" & n).NumberFormat = "#,##0.00"
    Set co = s.ChartObjects.Add(s.Range("G2").Left, s.Range("G2").Top, 800, 380)
    With co.Chart
        .ChartType = xlLineMarkers
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        With .SeriesCollection.NewSeries
            .Name = "Gross shipments": .XValues = s.Range("A2:A" & n): .Values = s.Range("B2:B" & n)
        End With
        With .SeriesCollection.NewSeries
            .Name = "Net of known returns": .XValues = s.Range("A2:A" & n): .Values = s.Range("D2:D" & n)
        End With
        .HasTitle = True: .ChartTitle.Text = "Shipments and sales net of known returns"
        .Axes(xlValue).MinimumScale = 0: .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "RUB"
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Original shipment week"
    End With
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "BuildDashboard"
End Sub
Public Sub AnalyseReturns()
    On Error GoTo Failed
    Dim weekly As Variant, regional As Variant, ranked As Variant, reasons As Variant
    Dim s As Worksheet, t As Worksheet, co As ChartObject, n As Long
    ANRead weekly, regional, ranked, reasons
    Set s = LVOut("OUT_Books"): LVWrite s, ranked
    Set t = LVOut("OUT_Reasons"): LVWrite t, reasons
    If UBound(ranked, 1) < 2 Then Exit Sub
    s.Range("E2:E" & UBound(ranked, 1)).NumberFormat = "0.0%"
    n = Application.Min(11, UBound(ranked, 1))
    Set co = s.ChartObjects.Add(s.Range("I2").Left, s.Range("I2").Top, 800, 420)
    With co.Chart
        .ChartType = xlBarClustered
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        With .SeriesCollection.NewSeries
            .Name = "Returned / shipped units": .XValues = s.Range("B2:B" & n): .Values = s.Range("E2:E" & n)
        End With
        .HasTitle = True: .ChartTitle.Text = "Top 10 books by return rate"
        .Axes(xlValue).MinimumScale = 0: .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .Axes(xlCategory).ReversePlotOrder = True
    End With
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "AnalyseReturns"
End Sub

Private Function ANKeyText(ByVal s As String) As String
    Dim i As Long, result As String
    For i = 1 To Len(s): result = result & Right$("0000" & Hex$(AscW(Mid$(s, i, 1)) And &HFFFF&), 4): Next i
    ANKeyText = result
End Function
