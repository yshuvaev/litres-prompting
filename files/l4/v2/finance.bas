Attribute VB_Name = "FinanceModule"
Option Explicit
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
        Case "OUT_Checked", "OUT_Issues", "OUT_ModelAudit", "OUT_Dashboard", "OUT_Partner", "OUT_Rejected"
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

Public Sub BuildFinanceModel()
    On Error GoTo Failed
    Dim inputSheet As Worksheet, calc As Worksheet, summary As Worksheet, s As Worksheet
    Dim a As Variant, names As Variant, i As Long, name As Variant, data(1 To 26, 1 To 14) As Variant
    Set inputSheet = ThisWorkbook.Worksheets("Inputs")
    a = inputSheet.Range("A2:B12").Value2
    names = Array("Тираж", "Розничная цена", "Скидка торговому партнёру", "Печать одного экземпляра", "Подготовка книги", "Стартовое продвижение", "Ежемесячные расходы", "Роялти с поступлений", "Спрос первого месяца", "Ежемесячное изменение спроса", "Доля возврата")
    For i = 1 To 11
        If IsError(a(i, 1)) Then Err.Raise 5, , "Invalid parameter label"
        If CStr(a(i, 1)) <> names(i - 1) Or Not LVNum(a(i, 2)) Then Err.Raise 5, , "Check Inputs row " & i + 1
        If i <> 10 Then
            If a(i, 2) < 0 Then Err.Raise 5, , "Negative parameter"
        End If
    Next i
    If a(1, 2) <> Fix(a(1, 2)) Or a(9, 2) <> Fix(a(9, 2)) Then Err.Raise 5, , "Print run and demand must be integers"
    If a(3, 2) > 1 Or a(8, 2) > 1 Or a(11, 2) > 1 Or a(10, 2) < -1 Or a(10, 2) > 1 Then Err.Raise 5, , "Invalid percentage"
    For Each name In Array("Calc", "Summary")
        Set s = Nothing
        On Error Resume Next
        Set s = ThisWorkbook.Worksheets(CStr(name))
        On Error GoTo Failed
        If Not s Is Nothing Then
            If IsError(s.Range("P1").Value) Then Err.Raise 5, , "Output sheet name conflict"
            If s.Range("P1").Value2 <> "L4_FINANCE_V1" Then Err.Raise 5, , "Sheet already exists: " & CStr(name)
        End If
    Next name
    inputSheet.Range("D2").Formula = "=IFERROR(IF(AND(COUNT(B2:B12)=11,MIN(B2:B10)>=0,B2=INT(B2),B10=INT(B10),B4<=1,B9<=1,B11>=-1,B11<=1,B12>=0,B12<=1),""OK"",""Проверьте параметры""),""Проверьте параметры"")"
    inputSheet.Range("A14").Formula = "=D2"
    Set calc = FinOutput("Calc"): Set summary = FinOutput("Summary")
    data(1, 1) = "Месяц"
    data(1, 2) = "Остаток на начало"
    data(1, 3) = "Спрос"
    data(1, 4) = "Отгружено"
    data(1, 5) = "Возвращено"
    data(1, 6) = "Продано после возвратов"
    data(1, 7) = "Поступления"
    data(1, 8) = "Роялти"
    data(1, 9) = "Расходы месяца"
    data(1, 10) = "Поток месяца"
    data(1, 11) = "Накопленный поток"
    data(1, 12) = "Остаток на конец"
    data(1, 13) = "Месяц окупаемости"
    data(1, 14) = "Нулевая линия"
    data(2, 1) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(2, 2) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(2, 3) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(2, 4) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(2, 5) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(2, 6) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(2, 7) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(2, 8) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(2, 9) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(2, 10) = "=IF(Inputs!$D$2=""OK"",-(Inputs!B2*Inputs!B5+Inputs!B6+Inputs!B7),NA())"
    data(2, 11) = "=IF(Inputs!$D$2=""OK"",J2,NA())"
    data(2, 12) = "=IF(Inputs!$D$2=""OK"",Inputs!B2,NA())"
    data(2, 13) = "=IF(Inputs!$D$2=""OK"","""",NA())"
    data(2, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(3, 1) = "=IF(Inputs!$D$2=""OK"",A2+1,NA())"
    data(3, 2) = "=IF(Inputs!$D$2=""OK"",L2,NA())"
    data(3, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A3-1),0)),NA())"
    data(3, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B3,C3),NA())"
    data(3, 5) = "=IF(Inputs!$D$2=""OK"",D3*Inputs!$B$12,NA())"
    data(3, 6) = "=IF(Inputs!$D$2=""OK"",D3-E3,NA())"
    data(3, 7) = "=IF(Inputs!$D$2=""OK"",F3*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(3, 8) = "=IF(Inputs!$D$2=""OK"",G3*Inputs!$B$9,NA())"
    data(3, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(3, 10) = "=IF(Inputs!$D$2=""OK"",G3-H3-I3,NA())"
    data(3, 11) = "=IF(Inputs!$D$2=""OK"",K2+J3,NA())"
    data(3, 12) = "=IF(Inputs!$D$2=""OK"",B3-D3,NA())"
    data(3, 13) = "=IF(Inputs!$D$2=""OK"",IF(K3>=0,A3,""""),NA())"
    data(3, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(4, 1) = "=IF(Inputs!$D$2=""OK"",A3+1,NA())"
    data(4, 2) = "=IF(Inputs!$D$2=""OK"",L3,NA())"
    data(4, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A4-1),0)),NA())"
    data(4, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B4,C4),NA())"
    data(4, 5) = "=IF(Inputs!$D$2=""OK"",D4*Inputs!$B$12,NA())"
    data(4, 6) = "=IF(Inputs!$D$2=""OK"",D4-E4,NA())"
    data(4, 7) = "=IF(Inputs!$D$2=""OK"",F4*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(4, 8) = "=IF(Inputs!$D$2=""OK"",G4*Inputs!$B$9,NA())"
    data(4, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(4, 10) = "=IF(Inputs!$D$2=""OK"",G4-H4-I4,NA())"
    data(4, 11) = "=IF(Inputs!$D$2=""OK"",K3+J4,NA())"
    data(4, 12) = "=IF(Inputs!$D$2=""OK"",B4-D4,NA())"
    data(4, 13) = "=IF(Inputs!$D$2=""OK"",IF(K4>=0,A4,""""),NA())"
    data(4, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(5, 1) = "=IF(Inputs!$D$2=""OK"",A4+1,NA())"
    data(5, 2) = "=IF(Inputs!$D$2=""OK"",L4,NA())"
    data(5, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A5-1),0)),NA())"
    data(5, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B5,C5),NA())"
    data(5, 5) = "=IF(Inputs!$D$2=""OK"",D5*Inputs!$B$12,NA())"
    data(5, 6) = "=IF(Inputs!$D$2=""OK"",D5-E5,NA())"
    data(5, 7) = "=IF(Inputs!$D$2=""OK"",F5*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(5, 8) = "=IF(Inputs!$D$2=""OK"",G5*Inputs!$B$9,NA())"
    data(5, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(5, 10) = "=IF(Inputs!$D$2=""OK"",G5-H5-I5,NA())"
    data(5, 11) = "=IF(Inputs!$D$2=""OK"",K4+J5,NA())"
    data(5, 12) = "=IF(Inputs!$D$2=""OK"",B5-D5,NA())"
    data(5, 13) = "=IF(Inputs!$D$2=""OK"",IF(K5>=0,A5,""""),NA())"
    data(5, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(6, 1) = "=IF(Inputs!$D$2=""OK"",A5+1,NA())"
    data(6, 2) = "=IF(Inputs!$D$2=""OK"",L5,NA())"
    data(6, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A6-1),0)),NA())"
    data(6, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B6,C6),NA())"
    data(6, 5) = "=IF(Inputs!$D$2=""OK"",D6*Inputs!$B$12,NA())"
    data(6, 6) = "=IF(Inputs!$D$2=""OK"",D6-E6,NA())"
    data(6, 7) = "=IF(Inputs!$D$2=""OK"",F6*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(6, 8) = "=IF(Inputs!$D$2=""OK"",G6*Inputs!$B$9,NA())"
    data(6, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(6, 10) = "=IF(Inputs!$D$2=""OK"",G6-H6-I6,NA())"
    data(6, 11) = "=IF(Inputs!$D$2=""OK"",K5+J6,NA())"
    data(6, 12) = "=IF(Inputs!$D$2=""OK"",B6-D6,NA())"
    data(6, 13) = "=IF(Inputs!$D$2=""OK"",IF(K6>=0,A6,""""),NA())"
    data(6, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(7, 1) = "=IF(Inputs!$D$2=""OK"",A6+1,NA())"
    data(7, 2) = "=IF(Inputs!$D$2=""OK"",L6,NA())"
    data(7, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A7-1),0)),NA())"
    data(7, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B7,C7),NA())"
    data(7, 5) = "=IF(Inputs!$D$2=""OK"",D7*Inputs!$B$12,NA())"
    data(7, 6) = "=IF(Inputs!$D$2=""OK"",D7-E7,NA())"
    data(7, 7) = "=IF(Inputs!$D$2=""OK"",F7*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(7, 8) = "=IF(Inputs!$D$2=""OK"",G7*Inputs!$B$9,NA())"
    data(7, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(7, 10) = "=IF(Inputs!$D$2=""OK"",G7-H7-I7,NA())"
    data(7, 11) = "=IF(Inputs!$D$2=""OK"",K6+J7,NA())"
    data(7, 12) = "=IF(Inputs!$D$2=""OK"",B7-D7,NA())"
    data(7, 13) = "=IF(Inputs!$D$2=""OK"",IF(K7>=0,A7,""""),NA())"
    data(7, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(8, 1) = "=IF(Inputs!$D$2=""OK"",A7+1,NA())"
    data(8, 2) = "=IF(Inputs!$D$2=""OK"",L7,NA())"
    data(8, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A8-1),0)),NA())"
    data(8, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B8,C8),NA())"
    data(8, 5) = "=IF(Inputs!$D$2=""OK"",D8*Inputs!$B$12,NA())"
    data(8, 6) = "=IF(Inputs!$D$2=""OK"",D8-E8,NA())"
    data(8, 7) = "=IF(Inputs!$D$2=""OK"",F8*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(8, 8) = "=IF(Inputs!$D$2=""OK"",G8*Inputs!$B$9,NA())"
    data(8, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(8, 10) = "=IF(Inputs!$D$2=""OK"",G8-H8-I8,NA())"
    data(8, 11) = "=IF(Inputs!$D$2=""OK"",K7+J8,NA())"
    data(8, 12) = "=IF(Inputs!$D$2=""OK"",B8-D8,NA())"
    data(8, 13) = "=IF(Inputs!$D$2=""OK"",IF(K8>=0,A8,""""),NA())"
    data(8, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(9, 1) = "=IF(Inputs!$D$2=""OK"",A8+1,NA())"
    data(9, 2) = "=IF(Inputs!$D$2=""OK"",L8,NA())"
    data(9, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A9-1),0)),NA())"
    data(9, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B9,C9),NA())"
    data(9, 5) = "=IF(Inputs!$D$2=""OK"",D9*Inputs!$B$12,NA())"
    data(9, 6) = "=IF(Inputs!$D$2=""OK"",D9-E9,NA())"
    data(9, 7) = "=IF(Inputs!$D$2=""OK"",F9*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(9, 8) = "=IF(Inputs!$D$2=""OK"",G9*Inputs!$B$9,NA())"
    data(9, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(9, 10) = "=IF(Inputs!$D$2=""OK"",G9-H9-I9,NA())"
    data(9, 11) = "=IF(Inputs!$D$2=""OK"",K8+J9,NA())"
    data(9, 12) = "=IF(Inputs!$D$2=""OK"",B9-D9,NA())"
    data(9, 13) = "=IF(Inputs!$D$2=""OK"",IF(K9>=0,A9,""""),NA())"
    data(9, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(10, 1) = "=IF(Inputs!$D$2=""OK"",A9+1,NA())"
    data(10, 2) = "=IF(Inputs!$D$2=""OK"",L9,NA())"
    data(10, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A10-1),0)),NA())"
    data(10, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B10,C10),NA())"
    data(10, 5) = "=IF(Inputs!$D$2=""OK"",D10*Inputs!$B$12,NA())"
    data(10, 6) = "=IF(Inputs!$D$2=""OK"",D10-E10,NA())"
    data(10, 7) = "=IF(Inputs!$D$2=""OK"",F10*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(10, 8) = "=IF(Inputs!$D$2=""OK"",G10*Inputs!$B$9,NA())"
    data(10, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(10, 10) = "=IF(Inputs!$D$2=""OK"",G10-H10-I10,NA())"
    data(10, 11) = "=IF(Inputs!$D$2=""OK"",K9+J10,NA())"
    data(10, 12) = "=IF(Inputs!$D$2=""OK"",B10-D10,NA())"
    data(10, 13) = "=IF(Inputs!$D$2=""OK"",IF(K10>=0,A10,""""),NA())"
    data(10, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(11, 1) = "=IF(Inputs!$D$2=""OK"",A10+1,NA())"
    data(11, 2) = "=IF(Inputs!$D$2=""OK"",L10,NA())"
    data(11, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A11-1),0)),NA())"
    data(11, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B11,C11),NA())"
    data(11, 5) = "=IF(Inputs!$D$2=""OK"",D11*Inputs!$B$12,NA())"
    data(11, 6) = "=IF(Inputs!$D$2=""OK"",D11-E11,NA())"
    data(11, 7) = "=IF(Inputs!$D$2=""OK"",F11*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(11, 8) = "=IF(Inputs!$D$2=""OK"",G11*Inputs!$B$9,NA())"
    data(11, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(11, 10) = "=IF(Inputs!$D$2=""OK"",G11-H11-I11,NA())"
    data(11, 11) = "=IF(Inputs!$D$2=""OK"",K10+J11,NA())"
    data(11, 12) = "=IF(Inputs!$D$2=""OK"",B11-D11,NA())"
    data(11, 13) = "=IF(Inputs!$D$2=""OK"",IF(K11>=0,A11,""""),NA())"
    data(11, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(12, 1) = "=IF(Inputs!$D$2=""OK"",A11+1,NA())"
    data(12, 2) = "=IF(Inputs!$D$2=""OK"",L11,NA())"
    data(12, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A12-1),0)),NA())"
    data(12, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B12,C12),NA())"
    data(12, 5) = "=IF(Inputs!$D$2=""OK"",D12*Inputs!$B$12,NA())"
    data(12, 6) = "=IF(Inputs!$D$2=""OK"",D12-E12,NA())"
    data(12, 7) = "=IF(Inputs!$D$2=""OK"",F12*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(12, 8) = "=IF(Inputs!$D$2=""OK"",G12*Inputs!$B$9,NA())"
    data(12, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(12, 10) = "=IF(Inputs!$D$2=""OK"",G12-H12-I12,NA())"
    data(12, 11) = "=IF(Inputs!$D$2=""OK"",K11+J12,NA())"
    data(12, 12) = "=IF(Inputs!$D$2=""OK"",B12-D12,NA())"
    data(12, 13) = "=IF(Inputs!$D$2=""OK"",IF(K12>=0,A12,""""),NA())"
    data(12, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(13, 1) = "=IF(Inputs!$D$2=""OK"",A12+1,NA())"
    data(13, 2) = "=IF(Inputs!$D$2=""OK"",L12,NA())"
    data(13, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A13-1),0)),NA())"
    data(13, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B13,C13),NA())"
    data(13, 5) = "=IF(Inputs!$D$2=""OK"",D13*Inputs!$B$12,NA())"
    data(13, 6) = "=IF(Inputs!$D$2=""OK"",D13-E13,NA())"
    data(13, 7) = "=IF(Inputs!$D$2=""OK"",F13*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(13, 8) = "=IF(Inputs!$D$2=""OK"",G13*Inputs!$B$9,NA())"
    data(13, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(13, 10) = "=IF(Inputs!$D$2=""OK"",G13-H13-I13,NA())"
    data(13, 11) = "=IF(Inputs!$D$2=""OK"",K12+J13,NA())"
    data(13, 12) = "=IF(Inputs!$D$2=""OK"",B13-D13,NA())"
    data(13, 13) = "=IF(Inputs!$D$2=""OK"",IF(K13>=0,A13,""""),NA())"
    data(13, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(14, 1) = "=IF(Inputs!$D$2=""OK"",A13+1,NA())"
    data(14, 2) = "=IF(Inputs!$D$2=""OK"",L13,NA())"
    data(14, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A14-1),0)),NA())"
    data(14, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B14,C14),NA())"
    data(14, 5) = "=IF(Inputs!$D$2=""OK"",D14*Inputs!$B$12,NA())"
    data(14, 6) = "=IF(Inputs!$D$2=""OK"",D14-E14,NA())"
    data(14, 7) = "=IF(Inputs!$D$2=""OK"",F14*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(14, 8) = "=IF(Inputs!$D$2=""OK"",G14*Inputs!$B$9,NA())"
    data(14, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(14, 10) = "=IF(Inputs!$D$2=""OK"",G14-H14-I14,NA())"
    data(14, 11) = "=IF(Inputs!$D$2=""OK"",K13+J14,NA())"
    data(14, 12) = "=IF(Inputs!$D$2=""OK"",B14-D14,NA())"
    data(14, 13) = "=IF(Inputs!$D$2=""OK"",IF(K14>=0,A14,""""),NA())"
    data(14, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(15, 1) = "=IF(Inputs!$D$2=""OK"",A14+1,NA())"
    data(15, 2) = "=IF(Inputs!$D$2=""OK"",L14,NA())"
    data(15, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A15-1),0)),NA())"
    data(15, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B15,C15),NA())"
    data(15, 5) = "=IF(Inputs!$D$2=""OK"",D15*Inputs!$B$12,NA())"
    data(15, 6) = "=IF(Inputs!$D$2=""OK"",D15-E15,NA())"
    data(15, 7) = "=IF(Inputs!$D$2=""OK"",F15*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(15, 8) = "=IF(Inputs!$D$2=""OK"",G15*Inputs!$B$9,NA())"
    data(15, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(15, 10) = "=IF(Inputs!$D$2=""OK"",G15-H15-I15,NA())"
    data(15, 11) = "=IF(Inputs!$D$2=""OK"",K14+J15,NA())"
    data(15, 12) = "=IF(Inputs!$D$2=""OK"",B15-D15,NA())"
    data(15, 13) = "=IF(Inputs!$D$2=""OK"",IF(K15>=0,A15,""""),NA())"
    data(15, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(16, 1) = "=IF(Inputs!$D$2=""OK"",A15+1,NA())"
    data(16, 2) = "=IF(Inputs!$D$2=""OK"",L15,NA())"
    data(16, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A16-1),0)),NA())"
    data(16, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B16,C16),NA())"
    data(16, 5) = "=IF(Inputs!$D$2=""OK"",D16*Inputs!$B$12,NA())"
    data(16, 6) = "=IF(Inputs!$D$2=""OK"",D16-E16,NA())"
    data(16, 7) = "=IF(Inputs!$D$2=""OK"",F16*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(16, 8) = "=IF(Inputs!$D$2=""OK"",G16*Inputs!$B$9,NA())"
    data(16, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(16, 10) = "=IF(Inputs!$D$2=""OK"",G16-H16-I16,NA())"
    data(16, 11) = "=IF(Inputs!$D$2=""OK"",K15+J16,NA())"
    data(16, 12) = "=IF(Inputs!$D$2=""OK"",B16-D16,NA())"
    data(16, 13) = "=IF(Inputs!$D$2=""OK"",IF(K16>=0,A16,""""),NA())"
    data(16, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(17, 1) = "=IF(Inputs!$D$2=""OK"",A16+1,NA())"
    data(17, 2) = "=IF(Inputs!$D$2=""OK"",L16,NA())"
    data(17, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A17-1),0)),NA())"
    data(17, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B17,C17),NA())"
    data(17, 5) = "=IF(Inputs!$D$2=""OK"",D17*Inputs!$B$12,NA())"
    data(17, 6) = "=IF(Inputs!$D$2=""OK"",D17-E17,NA())"
    data(17, 7) = "=IF(Inputs!$D$2=""OK"",F17*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(17, 8) = "=IF(Inputs!$D$2=""OK"",G17*Inputs!$B$9,NA())"
    data(17, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(17, 10) = "=IF(Inputs!$D$2=""OK"",G17-H17-I17,NA())"
    data(17, 11) = "=IF(Inputs!$D$2=""OK"",K16+J17,NA())"
    data(17, 12) = "=IF(Inputs!$D$2=""OK"",B17-D17,NA())"
    data(17, 13) = "=IF(Inputs!$D$2=""OK"",IF(K17>=0,A17,""""),NA())"
    data(17, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(18, 1) = "=IF(Inputs!$D$2=""OK"",A17+1,NA())"
    data(18, 2) = "=IF(Inputs!$D$2=""OK"",L17,NA())"
    data(18, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A18-1),0)),NA())"
    data(18, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B18,C18),NA())"
    data(18, 5) = "=IF(Inputs!$D$2=""OK"",D18*Inputs!$B$12,NA())"
    data(18, 6) = "=IF(Inputs!$D$2=""OK"",D18-E18,NA())"
    data(18, 7) = "=IF(Inputs!$D$2=""OK"",F18*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(18, 8) = "=IF(Inputs!$D$2=""OK"",G18*Inputs!$B$9,NA())"
    data(18, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(18, 10) = "=IF(Inputs!$D$2=""OK"",G18-H18-I18,NA())"
    data(18, 11) = "=IF(Inputs!$D$2=""OK"",K17+J18,NA())"
    data(18, 12) = "=IF(Inputs!$D$2=""OK"",B18-D18,NA())"
    data(18, 13) = "=IF(Inputs!$D$2=""OK"",IF(K18>=0,A18,""""),NA())"
    data(18, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(19, 1) = "=IF(Inputs!$D$2=""OK"",A18+1,NA())"
    data(19, 2) = "=IF(Inputs!$D$2=""OK"",L18,NA())"
    data(19, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A19-1),0)),NA())"
    data(19, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B19,C19),NA())"
    data(19, 5) = "=IF(Inputs!$D$2=""OK"",D19*Inputs!$B$12,NA())"
    data(19, 6) = "=IF(Inputs!$D$2=""OK"",D19-E19,NA())"
    data(19, 7) = "=IF(Inputs!$D$2=""OK"",F19*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(19, 8) = "=IF(Inputs!$D$2=""OK"",G19*Inputs!$B$9,NA())"
    data(19, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(19, 10) = "=IF(Inputs!$D$2=""OK"",G19-H19-I19,NA())"
    data(19, 11) = "=IF(Inputs!$D$2=""OK"",K18+J19,NA())"
    data(19, 12) = "=IF(Inputs!$D$2=""OK"",B19-D19,NA())"
    data(19, 13) = "=IF(Inputs!$D$2=""OK"",IF(K19>=0,A19,""""),NA())"
    data(19, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(20, 1) = "=IF(Inputs!$D$2=""OK"",A19+1,NA())"
    data(20, 2) = "=IF(Inputs!$D$2=""OK"",L19,NA())"
    data(20, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A20-1),0)),NA())"
    data(20, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B20,C20),NA())"
    data(20, 5) = "=IF(Inputs!$D$2=""OK"",D20*Inputs!$B$12,NA())"
    data(20, 6) = "=IF(Inputs!$D$2=""OK"",D20-E20,NA())"
    data(20, 7) = "=IF(Inputs!$D$2=""OK"",F20*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(20, 8) = "=IF(Inputs!$D$2=""OK"",G20*Inputs!$B$9,NA())"
    data(20, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(20, 10) = "=IF(Inputs!$D$2=""OK"",G20-H20-I20,NA())"
    data(20, 11) = "=IF(Inputs!$D$2=""OK"",K19+J20,NA())"
    data(20, 12) = "=IF(Inputs!$D$2=""OK"",B20-D20,NA())"
    data(20, 13) = "=IF(Inputs!$D$2=""OK"",IF(K20>=0,A20,""""),NA())"
    data(20, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(21, 1) = "=IF(Inputs!$D$2=""OK"",A20+1,NA())"
    data(21, 2) = "=IF(Inputs!$D$2=""OK"",L20,NA())"
    data(21, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A21-1),0)),NA())"
    data(21, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B21,C21),NA())"
    data(21, 5) = "=IF(Inputs!$D$2=""OK"",D21*Inputs!$B$12,NA())"
    data(21, 6) = "=IF(Inputs!$D$2=""OK"",D21-E21,NA())"
    data(21, 7) = "=IF(Inputs!$D$2=""OK"",F21*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(21, 8) = "=IF(Inputs!$D$2=""OK"",G21*Inputs!$B$9,NA())"
    data(21, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(21, 10) = "=IF(Inputs!$D$2=""OK"",G21-H21-I21,NA())"
    data(21, 11) = "=IF(Inputs!$D$2=""OK"",K20+J21,NA())"
    data(21, 12) = "=IF(Inputs!$D$2=""OK"",B21-D21,NA())"
    data(21, 13) = "=IF(Inputs!$D$2=""OK"",IF(K21>=0,A21,""""),NA())"
    data(21, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(22, 1) = "=IF(Inputs!$D$2=""OK"",A21+1,NA())"
    data(22, 2) = "=IF(Inputs!$D$2=""OK"",L21,NA())"
    data(22, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A22-1),0)),NA())"
    data(22, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B22,C22),NA())"
    data(22, 5) = "=IF(Inputs!$D$2=""OK"",D22*Inputs!$B$12,NA())"
    data(22, 6) = "=IF(Inputs!$D$2=""OK"",D22-E22,NA())"
    data(22, 7) = "=IF(Inputs!$D$2=""OK"",F22*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(22, 8) = "=IF(Inputs!$D$2=""OK"",G22*Inputs!$B$9,NA())"
    data(22, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(22, 10) = "=IF(Inputs!$D$2=""OK"",G22-H22-I22,NA())"
    data(22, 11) = "=IF(Inputs!$D$2=""OK"",K21+J22,NA())"
    data(22, 12) = "=IF(Inputs!$D$2=""OK"",B22-D22,NA())"
    data(22, 13) = "=IF(Inputs!$D$2=""OK"",IF(K22>=0,A22,""""),NA())"
    data(22, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(23, 1) = "=IF(Inputs!$D$2=""OK"",A22+1,NA())"
    data(23, 2) = "=IF(Inputs!$D$2=""OK"",L22,NA())"
    data(23, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A23-1),0)),NA())"
    data(23, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B23,C23),NA())"
    data(23, 5) = "=IF(Inputs!$D$2=""OK"",D23*Inputs!$B$12,NA())"
    data(23, 6) = "=IF(Inputs!$D$2=""OK"",D23-E23,NA())"
    data(23, 7) = "=IF(Inputs!$D$2=""OK"",F23*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(23, 8) = "=IF(Inputs!$D$2=""OK"",G23*Inputs!$B$9,NA())"
    data(23, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(23, 10) = "=IF(Inputs!$D$2=""OK"",G23-H23-I23,NA())"
    data(23, 11) = "=IF(Inputs!$D$2=""OK"",K22+J23,NA())"
    data(23, 12) = "=IF(Inputs!$D$2=""OK"",B23-D23,NA())"
    data(23, 13) = "=IF(Inputs!$D$2=""OK"",IF(K23>=0,A23,""""),NA())"
    data(23, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(24, 1) = "=IF(Inputs!$D$2=""OK"",A23+1,NA())"
    data(24, 2) = "=IF(Inputs!$D$2=""OK"",L23,NA())"
    data(24, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A24-1),0)),NA())"
    data(24, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B24,C24),NA())"
    data(24, 5) = "=IF(Inputs!$D$2=""OK"",D24*Inputs!$B$12,NA())"
    data(24, 6) = "=IF(Inputs!$D$2=""OK"",D24-E24,NA())"
    data(24, 7) = "=IF(Inputs!$D$2=""OK"",F24*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(24, 8) = "=IF(Inputs!$D$2=""OK"",G24*Inputs!$B$9,NA())"
    data(24, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(24, 10) = "=IF(Inputs!$D$2=""OK"",G24-H24-I24,NA())"
    data(24, 11) = "=IF(Inputs!$D$2=""OK"",K23+J24,NA())"
    data(24, 12) = "=IF(Inputs!$D$2=""OK"",B24-D24,NA())"
    data(24, 13) = "=IF(Inputs!$D$2=""OK"",IF(K24>=0,A24,""""),NA())"
    data(24, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(25, 1) = "=IF(Inputs!$D$2=""OK"",A24+1,NA())"
    data(25, 2) = "=IF(Inputs!$D$2=""OK"",L24,NA())"
    data(25, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A25-1),0)),NA())"
    data(25, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B25,C25),NA())"
    data(25, 5) = "=IF(Inputs!$D$2=""OK"",D25*Inputs!$B$12,NA())"
    data(25, 6) = "=IF(Inputs!$D$2=""OK"",D25-E25,NA())"
    data(25, 7) = "=IF(Inputs!$D$2=""OK"",F25*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(25, 8) = "=IF(Inputs!$D$2=""OK"",G25*Inputs!$B$9,NA())"
    data(25, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(25, 10) = "=IF(Inputs!$D$2=""OK"",G25-H25-I25,NA())"
    data(25, 11) = "=IF(Inputs!$D$2=""OK"",K24+J25,NA())"
    data(25, 12) = "=IF(Inputs!$D$2=""OK"",B25-D25,NA())"
    data(25, 13) = "=IF(Inputs!$D$2=""OK"",IF(K25>=0,A25,""""),NA())"
    data(25, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    data(26, 1) = "=IF(Inputs!$D$2=""OK"",A25+1,NA())"
    data(26, 2) = "=IF(Inputs!$D$2=""OK"",L25,NA())"
    data(26, 3) = "=IF(Inputs!$D$2=""OK"",MAX(0,ROUND(Inputs!$B$10*(1+Inputs!$B$11)^(A26-1),0)),NA())"
    data(26, 4) = "=IF(Inputs!$D$2=""OK"",MIN(B26,C26),NA())"
    data(26, 5) = "=IF(Inputs!$D$2=""OK"",D26*Inputs!$B$12,NA())"
    data(26, 6) = "=IF(Inputs!$D$2=""OK"",D26-E26,NA())"
    data(26, 7) = "=IF(Inputs!$D$2=""OK"",F26*Inputs!$B$3*(1-Inputs!$B$4),NA())"
    data(26, 8) = "=IF(Inputs!$D$2=""OK"",G26*Inputs!$B$9,NA())"
    data(26, 9) = "=IF(Inputs!$D$2=""OK"",Inputs!$B$8,NA())"
    data(26, 10) = "=IF(Inputs!$D$2=""OK"",G26-H26-I26,NA())"
    data(26, 11) = "=IF(Inputs!$D$2=""OK"",K25+J26,NA())"
    data(26, 12) = "=IF(Inputs!$D$2=""OK"",B26-D26,NA())"
    data(26, 13) = "=IF(Inputs!$D$2=""OK"",IF(K26>=0,A26,""""),NA())"
    data(26, 14) = "=IF(Inputs!$D$2=""OK"",0,NA())"
    calc.Range("A1:N26").Formula = data
    summary.Range("A1").Value2 = "Показатель": summary.Range("B1").Value2 = "Значение"
    summary.Range("A2").Value2 = "Первоначальные вложения, ₽": summary.Range("B2").Formula = "=-Calc!K2"
    summary.Range("A3").Value2 = "Первый месяц окупаемости": summary.Range("B3").Formula = "=IF(Inputs!D2<>""OK"",""Проверьте параметры"",IF(B2=0,0,IF(COUNT(Calc!M3:M26)=0,""За 24 месяца не окупается"",MIN(Calc!M3:M26))))"
    summary.Range("A4").Value2 = "Накопленный поток за 24 месяца, ₽": summary.Range("B4").Formula = "=Calc!K26"
    summary.Range("A5").Value2 = "Поступления за 24 месяца, ₽": summary.Range("B5").Formula = "=SUM(Calc!G3:G26)"
    summary.Range("A6").Value2 = "Продано после возвратов, экз.": summary.Range("B6").Formula = "=SUM(Calc!F3:F26)"
    summary.Range("A7").Value2 = "Остаток тиража, экз.": summary.Range("B7").Formula = "=Calc!L26"
    inputSheet.Range("B2:B12").Interior.Color = RGB(255, 241, 184)
    calc.Range("A1:N1").Interior.Color = RGB(35, 115, 77): calc.Range("A1:N1").Font.Color = vbWhite
    summary.Range("A1:B1").Interior.Color = RGB(35, 115, 77): summary.Range("A1:B1").Font.Color = vbWhite
    calc.Rows(1).Font.Bold = True: summary.Rows(1).Font.Bold = True
    calc.Columns("A:N").AutoFit: summary.Columns("A:B").AutoFit
    calc.Range("G2:K26").NumberFormat = "#,##0.00"
    summary.Range("B2").NumberFormat = "#,##0.00": summary.Range("B4:B5").NumberFormat = "#,##0.00"
    ' Automatic calculation is necessary for subsequent parameter edits.
    Application.Calculation = xlCalculationAutomatic
    inputSheet.Calculate: calc.Calculate: summary.Calculate
    Dim co As ChartObject
    For i = inputSheet.ChartObjects.Count To 1 Step -1
        If inputSheet.ChartObjects(i).Name = "L4FinanceChart" Then inputSheet.ChartObjects(i).Delete
    Next i
    Set co = inputSheet.ChartObjects.Add(inputSheet.Range("D2").Left, inputSheet.Range("D2").Top, 800, 400)
    co.Name = "L4FinanceChart"
    With co.Chart
        .ChartType = xlLine
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        With .SeriesCollection.NewSeries
            .Name = "Накопленный денежный поток": .XValues = calc.Range("A2:A26"): .Values = calc.Range("K2:K26")
        End With
        With .SeriesCollection.NewSeries
            .Name = "Нулевая линия": .XValues = calc.Range("A2:A26"): .Values = calc.Range("N2:N26")
        End With
        .HasTitle = True: .ChartTitle.Text = "Окупаемость книги"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "Рубли"
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Месяц"
    End With
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "BuildFinanceModel"
End Sub
Private Function FinOutput(ByVal name As String) As Worksheet
    Dim s As Worksheet
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If s Is Nothing Then
        Set s = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)): s.Name = name
    End If
    s.UsedRange.ClearContents
    Do While s.ChartObjects.Count > 0: s.ChartObjects(1).Delete: Loop
    s.Range("P1").Value2 = "L4_FINANCE_V1"
    Set FinOutput = s
End Function
