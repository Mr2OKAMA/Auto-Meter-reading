Attribute VB_Name = "InspectionJsonImporter"
Option Explicit

Private Const TARGET_SHEET_NAME As String = "記録"
Private Const DATE_HEADER_ROW As Long = 4
Private Const LABEL_SCAN_COLUMNS As Long = 12
Private Const msoFileDialogFilePicker As Long = 3

Public Sub Json点検データ取込_記録シート()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)

    Dim jsonPath As String
    jsonPath = PickJsonFilePath()
    If Len(jsonPath) = 0 Then Exit Sub

    Dim jsonText As String
    jsonText = ReadUtf8TextFile(jsonPath)
    If Len(Trim$(jsonText)) = 0 Then
        MsgBox "JSONファイルが空です。", vbExclamation
        Exit Sub
    End If

    Dim root As Object
    Set root = ParseJsonObject(jsonText)

    If Not root.Exists("点検日") Then
        MsgBox "JSONに「点検日」がありません。", vbExclamation
        Exit Sub
    End If

    If Not root.Exists("データ一覧") Then
        MsgBox "JSONに「データ一覧」がありません。", vbExclamation
        Exit Sub
    End If

    Dim inspectionDateKey As String
    inspectionDateKey = NormalizeDateKey(CStr(root("点検日")))
    If Len(inspectionDateKey) = 0 Then
        MsgBox "JSONの「点検日」を日付として解釈できません。値: " & CStr(root("点検日")), vbExclamation
        Exit Sub
    End If

    Dim targetColumn As Long
    targetColumn = FindDateColumn(ws, DATE_HEADER_ROW, inspectionDateKey)
    If targetColumn = 0 Then
        MsgBox "記録シート4行目に、JSONの点検日 " & inspectionDateKey & " と一致する列が見つかりません。既存データは変更していません。", vbExclamation
        Exit Sub
    End If

    Dim facilities As Object
    Set facilities = BuildFacilityMap()

    Dim itemAliasMap As Object
    Set itemAliasMap = BuildItemAliasMap()

    Dim layout As Object
    Set layout = BuildSheetLayoutMap(ws, facilities)

    Dim measurementRows As Object
    Set measurementRows = layout("measurement")
    Dim timeRows As Object
    Set timeRows = layout("time")

    Dim writePlans As Collection
    Set writePlans = New Collection

    Dim skippedEmptyCount As Long
    Dim skippedUnknownFacility As Long
    Dim skippedUnknownItem As Long
    Dim skippedAmbiguousTime As Long
    Dim skippedInvalidTime As Long

    BuildWritePlans root("データ一覧"), facilities, itemAliasMap, measurementRows, timeRows, targetColumn, _
                    writePlans, skippedEmptyCount, skippedUnknownFacility, skippedUnknownItem, skippedAmbiguousTime, skippedInvalidTime

    If writePlans.Count = 0 Then
        MsgBox "書き込み対象がありませんでした。" & vbCrLf & _
               "（空値スキップ: " & skippedEmptyCount & "件 / 施設未一致: " & skippedUnknownFacility & "件 / 項目未一致: " & skippedUnknownItem & "件）", vbInformation
        Exit Sub
    End If

    Dim message As String
    message = "点検日 " & inspectionDateKey & " の列（" & ColumnLetter(ws, targetColumn) & "列）に " & writePlans.Count & " 件書き込みます。" & vbCrLf & vbCrLf & _
              "空値スキップ: " & skippedEmptyCount & "件" & vbCrLf & _
              "施設未一致: " & skippedUnknownFacility & "件" & vbCrLf & _
              "項目未一致: " & skippedUnknownItem & "件" & vbCrLf & _
              "点検時間スキップ（欄未特定）: " & skippedAmbiguousTime & "件" & vbCrLf & _
              "点検時間スキップ（時刻形式不正）: " & skippedInvalidTime & "件" & vbCrLf & vbCrLf & _
              "続行しますか？"

    If MsgBox(message, vbQuestion + vbYesNo, "JSON取込確認") <> vbYes Then Exit Sub

    ApplyWritePlans ws, writePlans

    MsgBox "取込が完了しました。" & vbCrLf & _
           "書き込み件数: " & writePlans.Count & "件" & vbCrLf & _
           "空値スキップ: " & skippedEmptyCount & "件" & vbCrLf & _
           "施設未一致: " & skippedUnknownFacility & "件" & vbCrLf & _
           "項目未一致: " & skippedUnknownItem & "件" & vbCrLf & _
           "点検時間スキップ（欄未特定）: " & skippedAmbiguousTime & "件" & vbCrLf & _
           "点検時間スキップ（時刻形式不正）: " & skippedInvalidTime & "件", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "JSON取込中にエラーが発生しました。" & vbCrLf & _
           "原因: " & Err.Description, vbCritical
End Sub

Public Sub Json点検データ取込_セルフテスト()
    On Error GoTo ErrorHandler

    AssertEquals "2026-09-25", NormalizeDateKey("2026-09-25"), "日付正規化(yyyy-mm-dd)"
    AssertEquals "2026-09-25", NormalizeDateKey("2026/9/25"), "日付正規化(yyyy/m/d)"
    AssertEquals "", NormalizeDateKey("20261340"), "不正日付は不一致"

    Dim aliases As Object
    Set aliases = BuildItemAliasMap()
    AssertEquals "電力量", CanonicalItemName("電力", aliases), "項目別名"
    AssertEquals "流量（深田）", CanonicalItemName("流量(深田)", aliases), "括弧揺れ"

    Dim facilityMap As Object
    Set facilityMap = BuildFacilityMap()
    Dim facilityEntry As Object
    Set facilityEntry = CreateObject("Scripting.Dictionary")
    facilityEntry("施設名") = "Ｄ点（中屋川排水放流口）"
    AssertEquals "d_point", ResolveFacilityKey(facilityEntry, facilityMap), "施設名全角フォールバック"
    facilityEntry("施設名") = "D点（中屋川排水放流口）"
    AssertEquals "d_point", ResolveFacilityKey(facilityEntry, facilityMap), "施設名別名フォールバック"
    AssertEquals "kawashima_daini_nonino", ResolveFacilityFromRow(TokenizeNormalizedText("川島第二の二（環境楽園）"), NormalizeLabel("川島第二の二（環境楽園）"), facilityMap), "施設名最長一致"

    Dim okRoot As Object
    Set okRoot = ParseJsonObject("{""点検日"":""2026-09-25"",""データ一覧"":[]}")
    AssertTrue okRoot.Exists("点検日"), "正常JSON解析"

    Dim largeRoot As Object
    Set largeRoot = ParseJsonObject("{""n"":1000000000000001}")
    AssertEquals "1000000000000001", CStr(largeRoot("n")), "大きな整数は文字列保持"
    AssertEquals "1000000000000001", CStr(NormalizeCellValue(largeRoot("n"))), "書き込み前も精度保持"

    Dim decimalRoot As Object
    Set decimalRoot = ParseJsonObject("{""n"":0.1234567890123456789}")
    AssertEquals "0.1234567890123456789", CStr(decimalRoot("n")), "高精度小数は文字列保持"
    AssertEquals "09:05", CStr(NormalizeTimeValue("　9:05　")), "点検時間の全角空白トリム"

    AssertParseFail "{""点検日"":""2026-09-25""}garbage", "末尾ゴミ検知"
    AssertParseFail "{""n"":+1}", "不正数値(先頭プラス)検知"
    AssertParseFail "{""n"":01}", "不正数値(先頭ゼロ)検知"
    AssertParseFail "{""n"":1e}", "不正数値(指数欠落)検知"
    AssertParseFail "{""n"":foo}", "不正トークン検知"
    AssertParseFail "{""t"":""\uDC00""}", "単独下位サロゲート検知"

    MsgBox "セルフテストが完了しました。", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "セルフテスト失敗: " & Err.Description, vbCritical
End Sub

Private Sub BuildWritePlans(ByVal dataList As Variant, ByVal facilities As Object, ByVal itemAliasMap As Object, _
                            ByVal measurementRows As Object, ByVal timeRows As Object, ByVal targetColumn As Long, _
                            ByRef writePlans As Collection, ByRef skippedEmptyCount As Long, ByRef skippedUnknownFacility As Long, _
                            ByRef skippedUnknownItem As Long, ByRef skippedAmbiguousTime As Long, ByRef skippedInvalidTime As Long)
    If Not IsObject(dataList) Then Exit Sub

    Dim entry As Variant
    For Each entry In dataList
        If Not IsObject(entry) Then GoTo NextEntry

        Dim facilityKey As String
        facilityKey = ResolveFacilityKey(entry, facilities)
        If Len(facilityKey) = 0 Then
            skippedUnknownFacility = skippedUnknownFacility + 1
            GoTo NextEntry
        End If

        Dim measurements As Variant
        measurements = Empty
        If entry.Exists("測定値") Then measurements = entry("測定値")

        If IsObject(measurements) Then
            Dim measureName As Variant
            For Each measureName In measurements.Keys
                Dim rawValue As Variant
                rawValue = measurements(measureName)

                If IsJsonBlank(rawValue) Then
                    skippedEmptyCount = skippedEmptyCount + 1
                Else
                    Dim canonicalItem As String
                    canonicalItem = CanonicalItemName(CStr(measureName), itemAliasMap)

                    Dim rowKey As String
                    rowKey = facilityKey & "|" & NormalizeLabel(canonicalItem)

                    If measurementRows.Exists(rowKey) Then
                        AddWritePlan writePlans, CLng(measurementRows(rowKey)), targetColumn, NormalizeCellValue(rawValue)
                    Else
                        skippedUnknownItem = skippedUnknownItem + 1
                    End If
                End If
            Next measureName
        End If

        If entry.Exists("点検時間") Then
            Dim timeValue As Variant
            timeValue = entry("点検時間")
            If IsJsonBlank(timeValue) Then
                skippedEmptyCount = skippedEmptyCount + 1
            Else
                Dim timeKey As String
                timeKey = facilityKey & "|" & NormalizeLabel("点検時間")

                If timeRows.Exists(timeKey) Then
                    Dim normalizedTime As Variant
                    normalizedTime = NormalizeTimeValue(CStr(timeValue))
                    If IsNull(normalizedTime) Then
                        skippedInvalidTime = skippedInvalidTime + 1
                    Else
                        AddWritePlan writePlans, CLng(timeRows(timeKey)), targetColumn, normalizedTime
                    End If
                Else
                    ' 点検時間欄の場所がシート上で判別できない場合は書き込まない（測定値の反映を優先）。
                    skippedAmbiguousTime = skippedAmbiguousTime + 1
                End If
            End If
        End If
NextEntry:
    Next entry
End Sub

Private Sub ApplyWritePlans(ByVal ws As Worksheet, ByVal writePlans As Collection)
    Dim plan As Variant
    For Each plan In writePlans
        ws.Cells(CLng(plan("row")), CLng(plan("col"))).Value = plan("value")
    Next plan
End Sub

Private Sub AddWritePlan(ByRef writePlans As Collection, ByVal rowNumber As Long, ByVal colNumber As Long, ByVal cellValue As Variant)
    Dim item As Object
    Set item = CreateObject("Scripting.Dictionary")
    item("row") = rowNumber
    item("col") = colNumber
    item("value") = cellValue
    writePlans.Add item
End Sub

Private Function BuildSheetLayoutMap(ByVal ws As Worksheet, ByVal facilities As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim measurementRows As Object
    Set measurementRows = CreateObject("Scripting.Dictionary")
    Dim timeRows As Object
    Set timeRows = CreateObject("Scripting.Dictionary")

    Dim canonicalItems As Variant
    canonicalItems = CanonicalItemsByPriority()

    Dim usedLastRow As Long
    usedLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    Dim currentFacilityKey As String
    Dim r As Long
    For r = 1 To usedLastRow
        Dim rowTokens As Collection
        Set rowTokens = CollectRowLabelTokens(ws, r, LABEL_SCAN_COLUMNS)

        Dim rowText As String
        rowText = NormalizeLabel(JoinCollection(rowTokens, " "))
        If Len(rowText) = 0 Then GoTo ContinueLoop

        Dim facilityInRow As String
        facilityInRow = ResolveFacilityFromRow(rowTokens, rowText, facilities)
        If Len(facilityInRow) > 0 Then currentFacilityKey = facilityInRow

        If Len(currentFacilityKey) = 0 Then GoTo ContinueLoop

        Dim matchedItem As String
        matchedItem = ResolveItemFromRow(rowTokens, rowText, canonicalItems)
        If Len(matchedItem) = 0 Then GoTo ContinueLoop

        Dim key As String
        key = currentFacilityKey & "|" & NormalizeLabel(matchedItem)

        If matchedItem = "点検時間" Then
            If Not timeRows.Exists(key) Then timeRows(key) = r
        Else
            If Not measurementRows.Exists(key) Then measurementRows(key) = r
        End If
ContinueLoop:
    Next r

    result("measurement") = measurementRows
    result("time") = timeRows
    Set BuildSheetLayoutMap = result
End Function

Private Function CollectRowLabelTokens(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal maxLabelCol As Long) As Collection
    Dim parts As Collection
    Set parts = New Collection

    Dim c As Long
    For c = 1 To maxLabelCol
        Dim textValue As String
        textValue = Trim$(GetCellDisplayText(ws.Cells(rowNumber, c)))
        If Len(textValue) > 0 Then parts.Add NormalizeLabel(textValue)
    Next c
    Set CollectRowLabelTokens = parts
End Function

Private Function JoinCollection(ByVal items As Collection, ByVal delimiter As String) As String
    Dim i As Long
    For i = 1 To items.Count
        If i > 1 Then JoinCollection = JoinCollection & delimiter
        JoinCollection = JoinCollection & CStr(items(i))
    Next i
End Function

Private Function TokenizeNormalizedText(ByVal source As String) As Collection
    Dim tokens As New Collection
    tokens.Add NormalizeLabel(source)
    Set TokenizeNormalizedText = tokens
End Function

Private Function GetCellDisplayText(ByVal cell As Range) As String
    Dim target As Range
    If cell.MergeCells Then
        Set target = cell.MergeArea.Cells(1, 1)
    Else
        Set target = cell
    End If
    GetCellDisplayText = CStr(target.Text)
End Function

Private Function ResolveFacilityFromRow(ByVal normalizedRowTokens As Collection, ByVal normalizedRowText As String, ByVal facilities As Object) As String
    Dim bestKey As String
    Dim bestLength As Long

    Dim facilityKey As Variant
    For Each facilityKey In facilities.Keys
        Dim aliases As Collection
        Set aliases = facilities(facilityKey)("aliases")

        Dim aliasValue As Variant
        For Each aliasValue In aliases
            Dim aliasText As String
            aliasText = CStr(aliasValue)
            If Len(aliasText) = 0 Then GoTo ContinueAlias

            If ContainsCollectionValue(normalizedRowTokens, aliasText) Or InStr(1, normalizedRowText, aliasText, vbTextCompare) > 0 Then
                If Len(aliasText) > bestLength Then
                    bestLength = Len(aliasText)
                    bestKey = CStr(facilityKey)
                End If
            End If
ContinueAlias:
        Next aliasValue
    Next facilityKey

    ResolveFacilityFromRow = bestKey
End Function

Private Function ResolveItemFromRow(ByVal normalizedRowTokens As Collection, ByVal normalizedRowText As String, ByVal canonicalItems As Variant) As String
    Dim i As Long
    For i = LBound(canonicalItems) To UBound(canonicalItems)
        Dim itemName As String
        itemName = CStr(canonicalItems(i))
        Dim normalizedItem As String
        normalizedItem = NormalizeLabel(itemName)

        If ContainsCollectionValue(normalizedRowTokens, normalizedItem) Then
            ResolveItemFromRow = itemName
            Exit Function
        End If
    Next i

    For i = LBound(canonicalItems) To UBound(canonicalItems)
        itemName = CStr(canonicalItems(i))
        normalizedItem = NormalizeLabel(itemName)

        If InStr(1, normalizedRowText, normalizedItem, vbTextCompare) > 0 Then
            ResolveItemFromRow = itemName
            Exit Function
        End If
    Next i
End Function

Private Function ContainsCollectionValue(ByVal items As Collection, ByVal expected As String) As Boolean
    Dim item As Variant
    For Each item In items
        If CStr(item) = expected Then
            ContainsCollectionValue = True
            Exit Function
        End If
    Next item
End Function

Private Function CanonicalItemsByPriority() As Variant
    CanonicalItemsByPriority = Array( _
        "地下燃料 移送カウンター値合計", _
        "地下燃料 移送カウンター値", _
        "上水メーター（親）", _
        "上水メーター（子）", _
        "流量（深田）", _
        "流量（酒倉）", _
        "小出槽 油量", _
        "点検時間", _
        "上水メーター", _
        "井水メーター", _
        "電力量", _
        "小出槽", _
        "流量" _
    )
End Function

Private Function BuildFacilityMap() As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    AddFacility map, "ryokuen", "緑苑"
    AddFacility map, "fukada_sakagura", "深田酒倉"
    AddFacility map, "kawabe", "川辺"
    AddFacility map, "yaotsu", "八百津"
    AddFacility map, "wachi", "和知"
    AddFacility map, "kanayama_pump", "兼山ポンプ場"
    AddFacility map, "nakaedo", "中恵土"
    AddFacility map, "d_point", "Ｄ点（中屋川排水放流口）", "D点（中屋川排水放流口）"
    AddFacility map, "a_point", "Ａ点（三井川放流口）", "A点（三井川放流口）"
    AddFacility map, "kawashima_pump", "川島ポンプ場"
    AddFacility map, "kawashima_daini_nonino", "川島第二の二（環境楽園）"
    AddFacility map, "kawashima_daini", "川島第二"
    AddFacility map, "komeno", "米野"
    AddFacility map, "b_point", "Ｂ点（中屋川放流口）", "B点（中屋川放流口）"
    AddFacility map, "nagamori_pump", "長森ポンプ場"
    AddFacility map, "tobu_daini", "東部第二"
    AddFacility map, "c_point", "Ｃ点（中部排水放流口）", "C点（中部排水放流口）"
    AddFacility map, "e_point", "Ｅ点（徳田支線放流口）", "E点（徳田支線放流口）"
    AddFacility map, "shimohaguri", "下羽栗"
    AddFacility map, "ginan_pump", "岐南ポンプ場"
    AddFacility map, "ginan_nishi", "岐南西"
    AddFacility map, "umematsu", "梅松"
    AddFacility map, "tobu_daiichi", "東部第一"
    AddFacility map, "akemi", "芥見"

    Set BuildFacilityMap = map
End Function

Private Sub AddFacility(ByVal map As Object, ByVal facilityKey As String, ByVal facilityName As String, ParamArray aliases() As Variant)
    Dim item As Object
    Set item = CreateObject("Scripting.Dictionary")

    item("name") = facilityName

    Dim aliasCollection As Collection
    Set aliasCollection = New Collection
    aliasCollection.Add NormalizeLabel(facilityName)

    Dim i As Long
    For i = LBound(aliases) To UBound(aliases)
        aliasCollection.Add NormalizeLabel(CStr(aliases(i)))
    Next i

    item("aliases") = aliasCollection
    map(facilityKey) = item
End Sub

Private Function BuildItemAliasMap() As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    AddItemAlias map, "電力量", "電力"
    AddItemAlias map, "流量"
    AddItemAlias map, "流量（深田）", "流量(深田)"
    AddItemAlias map, "流量（酒倉）", "流量(酒倉)"
    AddItemAlias map, "上水メーター", "上水メータ"
    AddItemAlias map, "上水メーター（親）", "上水メーター(親)", "上水メータ（親）", "上水メータ(親)"
    AddItemAlias map, "上水メーター（子）", "上水メーター(子)", "上水メータ（子）", "上水メータ(子)"
    AddItemAlias map, "地下燃料 移送カウンター値", "地下燃料移送カウンター値"
    AddItemAlias map, "地下燃料 移送カウンター値合計", "地下燃料移送カウンター値合計", "地下燃料 移送カウンター値 計"
    AddItemAlias map, "小出槽"
    AddItemAlias map, "小出槽 油量", "小出槽油量"
    AddItemAlias map, "井水メーター", "井水メータ"
    AddItemAlias map, "点検時間"

    Set BuildItemAliasMap = map
End Function

Private Sub AddItemAlias(ByVal map As Object, ByVal canonicalName As String, ParamArray aliases() As Variant)
    Dim canonicalKey As String
    canonicalKey = NormalizeLabel(canonicalName)
    map(canonicalKey) = canonicalName

    Dim i As Long
    For i = LBound(aliases) To UBound(aliases)
        map(NormalizeLabel(CStr(aliases(i)))) = canonicalName
    Next i
End Sub

Private Function CanonicalItemName(ByVal rawName As String, ByVal itemAliasMap As Object) As String
    Dim normalized As String
    normalized = NormalizeLabel(rawName)

    If itemAliasMap.Exists(normalized) Then
        CanonicalItemName = CStr(itemAliasMap(normalized))
    Else
        CanonicalItemName = rawName
    End If
End Function

Private Function ResolveFacilityKey(ByVal entry As Object, ByVal facilities As Object) As String
    If entry.Exists("施設キー") Then
        Dim keyValue As String
        keyValue = Trim$(CStr(entry("施設キー")))
        If facilities.Exists(keyValue) Then
            ResolveFacilityKey = keyValue
            Exit Function
        End If
    End If

    If entry.Exists("施設名") Then
        Dim normalizedName As String
        normalizedName = NormalizeLabel(CStr(entry("施設名")))

        Dim facilityKey As Variant
        For Each facilityKey In facilities.Keys
            Dim aliases As Collection
            Set aliases = facilities(facilityKey)("aliases")

            Dim aliasValue As Variant
            For Each aliasValue In aliases
                If CStr(aliasValue) = normalizedName Then
                    ResolveFacilityKey = CStr(facilityKey)
                    Exit Function
                End If
            Next aliasValue
        Next facilityKey
    End If
End Function

Private Function FindDateColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal targetDateKey As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    Dim startCol As Long
    startCol = LABEL_SCAN_COLUMNS + 1
    If startCol > lastCol Then startCol = 1

    For c = startCol To lastCol
        Dim key As String
        key = NormalizeDateKey(ws.Cells(headerRow, c).Text)
        If Len(key) = 0 Then key = NormalizeDateKey(ws.Cells(headerRow, c).Value)
        If key = targetDateKey Then
            FindDateColumn = c
            Exit Function
        End If
    Next c

    If startCol > 1 Then
        For c = 1 To startCol - 1
            key = NormalizeDateKey(ws.Cells(headerRow, c).Text)
            If Len(key) = 0 Then key = NormalizeDateKey(ws.Cells(headerRow, c).Value)
            If key = targetDateKey Then
                FindDateColumn = c
                Exit Function
            End If
        Next c
    End If
End Function

Private Function NormalizeDateKey(ByVal value As Variant) As String
    On Error GoTo HandleFail

    If IsDate(value) Then
        NormalizeDateKey = Format$(CDate(value), "yyyy-mm-dd")
        Exit Function
    End If

    Dim s As String
    s = Trim$(CStr(value))
    If Len(s) = 0 Then Exit Function

    s = Replace(s, "年", "-")
    s = Replace(s, "月", "-")
    s = Replace(s, "日", "")
    s = Replace(s, "/", "-")
    s = Replace(s, ".", "-")

    If IsDate(s) Then
        NormalizeDateKey = Format$(CDate(s), "yyyy-mm-dd")
        Exit Function
    End If

    If Len(s) = 8 And IsNumeric(s) Then
        Dim ymd As String
        ymd = Left$(s, 4) & "-" & Mid$(s, 5, 2) & "-" & Right$(s, 2)
        If IsDate(ymd) Then NormalizeDateKey = Format$(CDate(ymd), "yyyy-mm-dd")
        Exit Function
    End If
    Exit Function

HandleFail:
    NormalizeDateKey = ""
End Function

Private Function NormalizeLabel(ByVal value As String) As String
    Dim s As String
    s = Trim$(value)
    s = ToNarrowAscii(s)

    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    s = Replace(s, vbTab, "")
    s = Replace(s, "（", "(")
    s = Replace(s, "）", ")")

    NormalizeLabel = LCase$(s)
End Function

Private Function ToNarrowAscii(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    Dim codePoint As Long
    Dim outText As String

    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        codePoint = AscW(ch)

        Select Case codePoint
            Case &HFF10 To &HFF19, &HFF21 To &HFF3A, &HFF41 To &HFF5A
                outText = outText & ChrW$(codePoint - &HFEE0)
            Case Else
                outText = outText & ch
        End Select
    Next i

    ToNarrowAscii = outText
End Function

Private Function ColumnLetter(ByVal ws As Worksheet, ByVal columnNumber As Long) As String
    ColumnLetter = Split(ws.Cells(1, columnNumber).Address(False, False), "1")(0)
End Function

Private Function PickJsonFilePath() As String
    Dim picker As Object
    Set picker = Application.FileDialog(msoFileDialogFilePicker)

    With picker
        .Title = "取込むJSONファイルを選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "JSON", "*.json"
        If .Show <> -1 Then Exit Function
        PickJsonFilePath = .SelectedItems(1)
    End With
End Function

Private Function ReadUtf8TextFile(ByVal filePath As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadUtf8TextFile = stm.ReadText(-1)
    stm.Close
End Function

Private Function IsJsonBlank(ByVal value As Variant) As Boolean
    If IsObject(value) Then
        IsJsonBlank = False
        Exit Function
    End If

    If IsNull(value) Then
        IsJsonBlank = True
    ElseIf VarType(value) = vbString Then
        IsJsonBlank = Len(Trim$(CStr(value))) = 0
    Else
        IsJsonBlank = False
    End If
End Function

Private Function NormalizeCellValue(ByVal rawValue As Variant) As Variant
    If VarType(rawValue) = vbString Then
        NormalizeCellValue = CStr(rawValue)
    Else
        NormalizeCellValue = rawValue
    End If
End Function

Private Function NormalizeTimeValue(ByVal rawTime As String) As Variant
    Dim s As String
    s = Trim$(Replace(rawTime, "　", " "))

    If Len(s) = 0 Then
        NormalizeTimeValue = Null
        Exit Function
    End If

    Dim normalized As String
    normalized = NormalizeTimeOnlyString(s)
    If Len(normalized) > 0 Then
        NormalizeTimeValue = normalized
        Exit Function
    End If

    NormalizeTimeValue = Null
End Function

Private Function NormalizeTimeOnlyString(ByVal rawTime As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^(?:[01]?\d|2[0-3]):[0-5]\d(?:[:][0-5]\d)?$"
    re.Global = False

    If Not re.Test(rawTime) Then Exit Function

    Dim parts() As String
    parts = Split(rawTime, ":")

    Dim hh As Long
    Dim mm As Long
    hh = CLng(parts(0))
    mm = CLng(parts(1))

    NormalizeTimeOnlyString = Format$(TimeSerial(hh, mm, 0), "hh:nn")
End Function

' ===== JSON parser (external reference不要) =====

Private Type JsonState
    Source As String
    Position As Long
    Length As Long
End Type

Private Function ParseJsonObject(ByVal jsonText As String) As Object
    Dim st As JsonState
    st.Source = jsonText
    st.Position = 1
    st.Length = Len(jsonText)

    Dim value As Variant
    value = ParseJsonValue(st)
    SkipJsonWhitespace st

    If Not IsObject(value) Then Err.Raise vbObjectError + 2100, , "JSONルートがオブジェクトではありません。"
    If st.Position <= st.Length Then Err.Raise vbObjectError + 2112, , "JSON末尾に不要な文字があります。"
    Set ParseJsonObject = value
End Function

Private Function ParseJsonValue(ByRef st As JsonState) As Variant
    SkipJsonWhitespace st

    Dim ch As String
    ch = PeekJsonChar(st)

    Select Case ch
        Case "{"
            Set ParseJsonValue = ParseJsonDictionary(st)
        Case "["
            Set ParseJsonValue = ParseJsonArray(st)
        Case """"
            ParseJsonValue = ParseJsonString(st)
        Case "t"
            ExpectJsonLiteral st, "true"
            ParseJsonValue = True
        Case "f"
            ExpectJsonLiteral st, "false"
            ParseJsonValue = False
        Case "n"
            ExpectJsonLiteral st, "null"
            ParseJsonValue = Null
        Case "-", "0" To "9"
            ParseJsonValue = ParseJsonNumber(st)
        Case Else
            Err.Raise vbObjectError + 2116, , "JSON値の先頭文字が不正です: '" & ch & "'"
    End Select

    SkipJsonWhitespace st
End Function

Private Function ParseJsonDictionary(ByRef st As JsonState) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ConsumeJsonChar st, "{"
    SkipJsonWhitespace st

    If PeekJsonChar(st) = "}" Then
        ConsumeJsonChar st, "}"
        Set ParseJsonDictionary = dict
        Exit Function
    End If

    Do
        SkipJsonWhitespace st
        Dim key As String
        key = ParseJsonString(st)

        SkipJsonWhitespace st
        ConsumeJsonChar st, ":"
        SkipJsonWhitespace st

        Dim parsedValue As Variant
        parsedValue = ParseJsonValue(st)
        If IsObject(parsedValue) Then
            Set dict(key) = parsedValue
        Else
            dict(key) = parsedValue
        End If

        SkipJsonWhitespace st
        Dim nextCh As String
        nextCh = PeekJsonChar(st)

        If nextCh = "}" Then
            ConsumeJsonChar st, "}"
            Exit Do
        End If

        ConsumeJsonChar st, ","
    Loop

    Set ParseJsonDictionary = dict
End Function

Private Function ParseJsonArray(ByRef st As JsonState) As Collection
    Dim arr As New Collection

    ConsumeJsonChar st, "["
    SkipJsonWhitespace st

    If PeekJsonChar(st) = "]" Then
        ConsumeJsonChar st, "]"
        Set ParseJsonArray = arr
        Exit Function
    End If

    Do
        Dim parsedItem As Variant
        parsedItem = ParseJsonValue(st)
        If IsObject(parsedItem) Then
            Dim parsedObj As Object
            Set parsedObj = parsedItem
            arr.Add parsedObj
        Else
            arr.Add parsedItem
        End If
        SkipJsonWhitespace st

        Dim nextCh As String
        nextCh = PeekJsonChar(st)

        If nextCh = "]" Then
            ConsumeJsonChar st, "]"
            Exit Do
        End If

        ConsumeJsonChar st, ","
    Loop

    Set ParseJsonArray = arr
End Function

Private Function ParseJsonString(ByRef st As JsonState) As String
    ConsumeJsonChar st, """"

    Dim result As String
    result = ""

    Do While st.Position <= st.Length
        Dim ch As String
        ch = Mid$(st.Source, st.Position, 1)
        st.Position = st.Position + 1

        If ch = """" Then
            ParseJsonString = result
            Exit Function
        ElseIf ch = "\" Then
            If st.Position > st.Length Then Err.Raise vbObjectError + 2101, , "JSON文字列のエスケープが不正です。"

            Dim esc As String
            esc = Mid$(st.Source, st.Position, 1)
            st.Position = st.Position + 1

            Select Case esc
                Case """", "\", "/"
                    result = result & esc
                Case "b"
                    result = result & Chr$(8)
                Case "f"
                    result = result & Chr$(12)
                Case "n"
                    result = result & vbLf
                Case "r"
                    result = result & vbCr
                Case "t"
                    result = result & vbTab
                Case "u"
                    result = result & ParseJsonUnicodeEscape(st)
                Case Else
                    Err.Raise vbObjectError + 2102, , "JSON文字列のエスケープが不正です: \\" & esc
            End Select
        Else
            If AscW(ch) >= 0 And AscW(ch) <= 31 Then
                Err.Raise vbObjectError + 2115, , "JSON文字列に未エスケープ制御文字が含まれています。"
            End If
            result = result & ch
        End If
    Loop

    Err.Raise vbObjectError + 2103, , "JSON文字列が閉じられていません。"
End Function

Private Function ParseJsonUnicodeEscape(ByRef st As JsonState) As String
    If st.Position + 3 > st.Length Then Err.Raise vbObjectError + 2104, , "Unicodeエスケープが不正です。"

    Dim hexCode As String
    hexCode = Mid$(st.Source, st.Position, 4)
    st.Position = st.Position + 4

    If Not IsHex4(hexCode) Then Err.Raise vbObjectError + 2105, , "Unicodeエスケープが不正です: " & hexCode

    Dim highCode As Long
    highCode = CLng("&H" & hexCode)

    If highCode >= &HD800 And highCode <= &HDBFF Then
        If st.Position + 5 <= st.Length And Mid$(st.Source, st.Position, 2) = "\u" Then
            st.Position = st.Position + 2
            Dim lowHex As String
            lowHex = Mid$(st.Source, st.Position, 4)
            st.Position = st.Position + 4

            If Not IsHex4(lowHex) Then Err.Raise vbObjectError + 2109, , "Unicodeサロゲートの下位コードが不正です: " & lowHex

            Dim lowCode As Long
            lowCode = CLng("&H" & lowHex)
            If lowCode < &HDC00 Or lowCode > &HDFFF Then Err.Raise vbObjectError + 2110, , "Unicodeサロゲートペアが不正です。"

            ParseJsonUnicodeEscape = ChrW$(highCode) & ChrW$(lowCode)
            Exit Function
        Else
            Err.Raise vbObjectError + 2111, , "Unicodeサロゲートペアが途中で終わっています。"
        End If
    ElseIf highCode >= &HDC00 And highCode <= &HDFFF Then
        Err.Raise vbObjectError + 2114, , "Unicode下位サロゲートが単独で出現しました。"
    End If

    ParseJsonUnicodeEscape = ChrW$(highCode)
End Function

Private Function IsHex4(ByVal value As String) As Boolean
    Dim i As Long
    If Len(value) <> 4 Then Exit Function

    For i = 1 To 4
        Dim ch As String
        ch = Mid$(value, i, 1)
        If InStr(1, "0123456789abcdefABCDEF", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i

    IsHex4 = True
End Function

Private Function ParseJsonNumber(ByRef st As JsonState) As Variant
    Dim startPos As Long
    startPos = st.Position

    Do While st.Position <= st.Length
        Dim ch As String
        ch = Mid$(st.Source, st.Position, 1)
        Select Case ch
            Case "0" To "9", ".", "e", "E"
                st.Position = st.Position + 1
            Case "-", "+"
                If st.Position = startPos Then
                    If ch = "-" Then
                        st.Position = st.Position + 1
                    Else
                        Exit Do
                    End If
                Else
                    Dim prev As String
                    prev = Mid$(st.Source, st.Position - 1, 1)
                    If prev = "e" Or prev = "E" Then
                        st.Position = st.Position + 1
                    Else
                        Exit Do
                    End If
                End If
            Case Else
                Exit Do
        End Select
    Loop

    Dim token As String
    token = Mid$(st.Source, startPos, st.Position - startPos)

    If Len(token) = 0 Then Err.Raise vbObjectError + 2106, , "JSONの値を解析できません。"
    If Not IsValidJsonNumberToken(token) Then Err.Raise vbObjectError + 2113, , "JSON数値の形式が不正です: " & token

    Dim integerToken As Boolean
    integerToken = (InStr(1, token, ".", vbBinaryCompare) = 0 And InStr(1, token, "e", vbTextCompare) = 0)

    If integerToken Then
        Dim digitToken As String
        digitToken = token
        digitToken = Replace(digitToken, "+", "")
        digitToken = Replace(digitToken, "-", "")

        If Not IsIntegerExactlyRepresentable(digitToken) Then
            ParseJsonNumber = token
        Else
            ParseJsonNumber = CDbl(token)
        End If
    Else
        ' 小数・指数表現は文字列のまま保持し、桁落ちを防ぐ。
        ParseJsonNumber = token
    End If
End Function

Private Function IsIntegerExactlyRepresentable(ByVal unsignedDigits As String) As Boolean
    Dim normalized As String
    normalized = unsignedDigits

    Do While Len(normalized) > 1 And Left$(normalized, 1) = "0"
        normalized = Mid$(normalized, 2)
    Loop

    IsIntegerExactlyRepresentable = (Len(normalized) <= 15)
End Function

Private Function IsValidJsonNumberToken(ByVal token As String) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "^-?(0|[1-9]\d*)(\.\d+)?([eE][+-]?\d+)?$"
        re.Global = False
    End If
    IsValidJsonNumberToken = re.Test(token)
End Function

Private Sub ExpectJsonLiteral(ByRef st As JsonState, ByVal literal As String)
    Dim segment As String
    segment = Mid$(st.Source, st.Position, Len(literal))

    If segment <> literal Then Err.Raise vbObjectError + 2107, , "JSONリテラルが不正です。期待値: " & literal
    st.Position = st.Position + Len(literal)
End Sub

Private Sub ConsumeJsonChar(ByRef st As JsonState, ByVal expectedChar As String)
    Dim ch As String
    ch = PeekJsonChar(st)
    If ch <> expectedChar Then
        Err.Raise vbObjectError + 2108, , "JSON構文エラー。期待値: '" & expectedChar & "' 実値: '" & ch & "'"
    End If
    st.Position = st.Position + 1
End Sub

Private Function PeekJsonChar(ByRef st As JsonState) As String
    If st.Position > st.Length Then
        PeekJsonChar = ""
    Else
        PeekJsonChar = Mid$(st.Source, st.Position, 1)
    End If
End Function

Private Sub SkipJsonWhitespace(ByRef st As JsonState)
    Do While st.Position <= st.Length
        Dim ch As String
        ch = Mid$(st.Source, st.Position, 1)

        Select Case ch
            Case " ", vbTab, vbCr, vbLf
                st.Position = st.Position + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Sub AssertEquals(ByVal expected As String, ByVal actual As String, ByVal testName As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 2200, , testName & " 期待値=[" & expected & "] 実際=[" & actual & "]"
    End If
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal testName As String)
    If Not condition Then
        Err.Raise vbObjectError + 2201, , testName & " が失敗しました。"
    End If
End Sub

Private Sub AssertParseFail(ByVal jsonText As String, ByVal testName As String)
    On Error GoTo ExpectedFail
    Dim obj As Object
    Set obj = ParseJsonObject(jsonText)
    Err.Raise vbObjectError + 2202, , testName & " が失敗しました（本来はエラーになるべき入力を受理しました）。"
ExpectedFail:
    If Err.Number = vbObjectError + 2202 Then
        Err.Raise Err.Number, , Err.Description
    End If
    Err.Clear
End Sub
