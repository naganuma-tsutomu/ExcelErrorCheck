VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckDiffClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private isWbOpen As Boolean
Private thisWb As Workbook
Private thisWs As Worksheet
Public inspectWb As Workbook
Public inspectWs As Worksheet
Private openingSheetNum As String
Private thicknessSheetNum As String
Public searchYearValue As String
Public searchDevValue As String
Private devicesColumn As Long
Private cellArray() As Range

' コンストラクタ
Private Sub Class_Initialize()
    ' ワークブックを指定する
    Set thisWb = ThisWorkbook                    ' 検査記録ブックを変数へ
    ' ワークシートを指定する
    Set thisWs = thisWb.ActiveSheet              ' 検査記録ブックの現在のワークシートを指定する
End Sub

' ワークブックの指定
Public Function initBook(directory As Long, DevCell As String, YearCell As String, devMasterCol As Long) As CheckDiffClass
    Set initBook = Me
    searchDevValue = searchCellsCheck(DevCell)   ' 検索する機器の値を取得
    If searchDevValue <> "" Then
        ' データファイルの名前の設定(計画ファイル)
        Set inspectWb = targetWorkbook(searchDevValue, directory) ' 対象ブックを開く
        If Not inspectWb Is Nothing Then
            ' 検索値を指定する
            searchYearValue = searchCellsCheck(YearCell) ' 検索する値を取得
            devicesColumn = devMasterCol                 ' 機器検索対象とする列を代入
        End If
    End If
End Function

' ワークシートの指定
Public Function initSheet(sheetName As String) As CheckDiffClass
    Set initSheet = Me
    Set inspectWs = Nothing ' Nothingを代入しておく
    On Error GoTo myError
    Set inspectWs = inspectWb.Worksheets(sheetName) ' 検査計画ファイルの最初のワークシートを指定する
    Exit Function
myError:
    OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & sheetName & "シートが周期表ファイル内で見つかりませんでした。" & vbCrLf
End Function

' 検査計画ファイルのワークシート(開放実績)の設定
Public Function initOpeningSheet(sheetNum As String) As CheckDiffClass
    Set initOpeningSheet = Me
    openingSheetNum = sheetNum ' 開放実績のワークシートのインデックスを代入
    Set inspectWs = inspectWb.Worksheets(openingSheetNum) ' 開放実績のワークシートを指定する
End Function

' 検査計画ファイルのワークシート(肉厚測定)の設定
Public Function initThicknessSheet(sheetNum As String) As CheckDiffClass
    Set initThicknessSheet = Me
    thicknessSheetNum = sheetNum ' 肉厚測定のワークシートのインデックスを代入
    Set inspectWs = inspectWb.Worksheets(thicknessSheetNum) ' 検査計画ファイルのワークシートを指定する
End Function

' 検索値の取得
Private Function searchCellsCheck(Cell As String) As String
    searchCellsCheck = thisWs.Range(Cell).Value
    If searchCellsCheck = "" Then
        Dim searchTitle As String: searchTitle = Range(Cell).Offset(0, -1).Value
        MsgBox "入力欄に" & searchTitle & "を入力してください。", vbExclamation
    End If
End Function

' 対象ブックを開く
Private Function targetWorkbook(devName As String, directory As Long) As Workbook
    If isBookOpen(devName & "_高圧ガス周期表.xlsx") = True Then
        isWbOpen = True                          ' 検査計画ファイルを開いているフラグを立てる
        Set targetWorkbook = Workbooks(devName & "_高圧ガス周期表.xlsx") ' 検査計画ファイルを変数へ
    Else
        ' 階層ごとに参照するパスの変更
        Dim pathSub As String
        If directory = 1 Then
            pathSub = "\../"
        ElseIf directory = 2 Then
            pathSub = "\../../"
        Else
            pathSub = "/"
        End If

        ' ファイル階層(修正必要)
        Dim Path As String: Path = ThisWorkbook.Path & pathSub
        Dim wbPath As String: wbPath = FileSearch(Path)

        If wbPath <> "" Then
            Set targetWorkbook = Workbooks.Open(Filename:=wbPath, UpdateLinks:=False, ReadOnly:=True) ' 検査計画ファイルを開く
            'ActiveWindow.Visible = False         ' 検査計画ファイルを非表示
            Exit Function
        End If
    End If
End Function

Function FileSearch(Path As String)
    Dim FSO As Object, Folder As Variant, subFolder As Variant, File As Variant
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each Folder In FSO.GetFolder(Path).SubFolders
        If Folder.Name Like "*周期表*" Then
        Dim wbFolder As String: wbFolder = Mid(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\") + 1)
            For Each subFolder In FSO.GetFolder(Folder.Path).SubFolders
                If subFolder.Name Like wbFolder Then
                    Dim filePath As String
                    For Each File In FSO.GetFolder(subFolder.Path).Files
                        If File.Name Like searchDevValue & "*" Then
                            FileSearch = File.Path
                            Exit Function
                        End If
                    Next File
                    MsgBox "周期表ファイルが見つかりませんでした。機器名称を確認してもう一度お試し下さい。", vbExclamation
                    Exit Function
                End If
            Next subFolder
            MsgBox "周期表フォルダが見つかりませんでした。ファルダ構成を確認してもう一度お試し下さい。", vbExclamation
            Exit Function
        End If
    Next Folder
    MsgBox "周期表フォルダが見つかりませんでした。ファルダ構成を確認してもう一度お試し下さい。", vbExclamation
End Function

' ブックが開かれているか確認
Private Function isBookOpen(bookName) As Boolean
    Dim wb As Workbook
    isBookOpen = False
    For Each wb In Workbooks
        If wb.Name = bookName Then
            isBookOpen = True
            Exit For
        End If
    Next
End Function

'初期時に開いていなければファイルを閉じる
Public Sub CloseBook()
    If Not isWbOpen = True And Not inspectWb Is Nothing Then inspectWb.Close SaveChanges:=False
End Sub

' ------------------------------------------------------------------------------------------

Public Sub CheckInspection(headRow As Long, devGroupCol As String, devNumCol As String, devNameCol As String)
    Dim firstRow As Long
    Dim lastRow As Long
    Dim yearColumn As Long

    ' 検索対象の最初の行の取得
    firstRow = findRows(inspectWs, "first", searchDevValue, devGroupCol)
    ' 検索対象の最後の行の取得
    lastRow = findRows(inspectWs, "last", searchDevValue, devGroupCol)

    If firstRow > 0 And lastRow > 0 Then
        ' 検索年度の列の取得
        yearColumn = findYearColumn(inspectWs, headRow, searchYearValue)
        If yearColumn = 0 Then
            OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & inspectWs.Name & "にて対象年度が見つかりませんでした。" & vbCrLf
            Exit Sub
        End If
    Else
            OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & inspectWs.Name & "にて対象装置が見つかりませんでした。" & vbCrLf
        Exit Sub
    End If
    ' 機器一つ一つを計画と合致するか検索する
    Call FindMatchingDev(devNumCol, devNameCol, yearColumn, firstRow, lastRow)
End Sub

Public Sub CheckInspectionPiping(headRow As Long, devGroupCol As String, devNumCol As String, devNameCol As String)
    Dim firstRow As Long, lastRow As Long, yearColumn As Long, thickYearColumn As Long

    ' 検索年度の列の取得
    yearColumn = findYearColumn(inspectWs, headRow, searchYearValue)
    If yearColumn = 0 Then
            OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & inspectWs.Name & "にて対象年度が見つかりませんでした。" & vbCrLf
        Exit Sub
    End If

    ' 肉厚測定内の機器の取得
    Dim thickSheet As Worksheet: Set thickSheet = inspectWb.Worksheets(thicknessSheet) ' 肉厚測定のシートの取得3

    ' 検索対象の最初の行の取得
    firstRow = findRows(thickSheet, "first", searchDevValue, devGroupCol)
    ' 検索対象の最後の行の取得
    lastRow = findRows(thickSheet, "last", searchDevValue, devGroupCol)

    If firstRow > 0 And lastRow > 0 Then
        ' 検索年度の列の取得
        thickYearColumn = findYearColumn(thickSheet, thicknessHeadRow, searchYearValue)
        If thickYearColumn = 0 Then
            OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & thickSheet.Name & "にて対象年度が見つからないため、配管周期表と肉厚測定周期表のチェックができませんでした。" & vbCrLf
            Exit Sub
        End If
    Else
            OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & thickSheet.Name & "にて対象装置が見つからないため、配管周期表と肉厚測定周期表のチェックができませんでした。" & vbCrLf
        Exit Sub
    End If

    firstRow = headRow + 1
    lastRow = inspectWs.Cells(Rows.count, devNameCol).End(xlUp).Row

    ' 機器一つ一つを計画と合致するか検索する
    Call FindMatchingDev(devNumCol, devNameCol, yearColumn, firstRow, lastRow)
End Sub

' 対象となる最初と最後の行を取得
Private Function findRows(ws As Worksheet, order As String, searchDevValue As String, column As String)
    Dim i As Long
    With ws
        If order = "first" Then
            ' 最初の行番号を検索
            For i = 1 To .Cells(Rows.count, column).End(xlUp).Row
                If .Cells(i, column).Value = searchDevValue Then
                    findRows = i
                    Exit For
                End If
            Next i
        ElseIf order = "last" Then
            ' 最後の行番号を検索
            For i = .Cells(Rows.count, column).End(xlUp).Row To 1 Step -1
                If .Cells(i, column).Value = searchDevValue Then
                    findRows = i
                    Exit For
                End If
            Next i
        End If
    End With

End Function

' 検索対象年度の列の取得
Private Function findYearColumn(ws As Worksheet, searchHeadRow As Long, yearValue As String) As Long
    ' 年度の検索値の成形
    Dim year As String
    If InStr(ws.Name, pipingSheet) = 0 Then '配管シート以外
        year = substrYear(yearValue)
    Else ' 配管シート
        year = substrYearPip(yearValue)
    End If
    Dim i As Long
    findYearColumn = 0
    For i = 1 To 256
        If InStr(ws.Cells(searchHeadRow, i).Value, year) > 0 Then
            findYearColumn = i
            Exit For
        End If
    Next i
End Function

' 年度検索値の成形 配管シートではない場合
Private Function substrYear(yearValue As String) As String
    substrYear = yearValue
    ' 年が含まれていなければ追加
    If Not InStr(substrYear, "年") > 0 Then substrYear = substrYear & "年"
    ' 右から3文字を取得
    substrYear = Right(substrYear, 3)
    Dim i As Long
    ' 数値のみを半角へ変換
    For i = 1 To Len(substrYear)
        If Mid(substrYear, i, 1) Like "[０-９]" Then
            Mid(substrYear, i, 1) = StrConv(Mid(substrYear, i, 1), vbNarrow)
        End If
    Next
End Function

' 年度検索値の成形 配管シートの場合
Private Function substrYearPip(yearValue As String) As String
    substrYearPip = yearValue
    Dim i As Long
    ' 数値のみを半角へ変換
    For i = 1 To Len(substrYearPip)
        If Mid(substrYearPip, i, 1) Like "[０-９]" Then
            Mid(substrYearPip, i, 1) = StrConv(Mid(substrYearPip, i, 1), vbNarrow)
        End If
    Next
End Function

' 計画ファイルからデータを検索し、記録ファイルとの差異をチェックする
Private Sub FindMatchingDev(devNumCol As String, devNameCol As String, yearColumn As Long, firstRow As Long, lastRow As Long)
    Dim firstRecordRow As Long: firstRecordRow = devTopCell ' 検査記録ファイルの基準の最初の行を代入
    Dim lastRecordRow As Long: lastRecordRow = thisWs.Cells(thisWs.Rows.count, devicesColumn).End(xlUp).Row ' 検査記録ファイルの最後の行を取得
    Dim message As String                        ' エラーメッセージ変数宣言
    Dim arrCount As Long                         ' エラーセルのカウント変数
    ' エラーセル配列が存在していれば配列の数を代入 / 配列が存在しなければ0を代入
    If Not Not cellArray Then arrCount = UBound(cellArray) + 1 Else arrCount = 0
    ' 検査記録ファイルの各セルについて、検査計画ファイルと比較し、合致するセルを取得
    With inspectWs
        Dim i As Long
        For i = firstRecordRow To lastRecordRow
            Dim TargetCell As Range: Set TargetCell = thisWs.Cells(i, devicesColumn) ' 検索対象のセルの取得
            Dim NumValue As String: NumValue = TargetCell.Value ' 検索する機器番号の取得
            Dim NameValue As String: NameValue = TargetCell.Offset(0, 1).Value ' 検索する機器名称の取得
            Dim j As Long, devNum As String, devName As String, planMarkerCell As Range
            If Not NumValue <> "" Then
                ' エラーセルの配列に格納
                ReDim Preserve cellArray(arrCount)
                Set cellArray(arrCount) = TargetCell.Offset(0, 1)
                arrCount = arrCount + 1
                message = message & "番号:" & TargetCell.End(xlToLeft).Value & "の機器番号を入力してください。" & vbCrLf
            ElseIf InStr(.Name, pipingSheet) = 0 And NumValue <> pipingDevNum Then ' 配管以外のシートかつ機器番号が配管でない場合
                ' 対象の最初の行から最後の行まで検索をし、機器番号・機器名称をチェックする (配管以外)
                For j = firstRow To lastRow
                    devNum = .Cells(j, devNumCol).Value ' 現在の行の機器番号の取得
                    If NumValue = devNum Then    ' 機器番号が一致した場合
                        devName = .Cells(j, devNameCol).Value ' 現在の行の機器名称の取得
                        If NameValue = devName Then ' 機器名称が一致した場合
                            Set planMarkerCell = .Cells(j, yearColumn) ' 検索年度のマーカーのセルを取得
                            ' 検索年度のマーカーをチェックするプロシージャを呼ぶ
                            Call checkPlanMarker(planMarkerCell, TargetCell, j)
                            Exit For
                        ElseIf NameValue <> devName Then ' 機器名称が一致しない場合
                            ' エラーセルの配列に格納
                            ReDim Preserve cellArray(arrCount)
                            Set cellArray(arrCount) = TargetCell.Offset(0, 1)
                            arrCount = arrCount + 1
                            ' メッセージ変数に追加
                            message = message & "番号:" & TargetCell.End(xlToLeft).Value & " - " & .Name & "にて機器番号: " & NumValue & " の機器名称が一致しません。" & vbCrLf
                            Exit For
                        End If
                    ElseIf j = lastRow Then
                        ' エラーセルの配列に格納
                        ReDim Preserve cellArray(arrCount)
                        Set cellArray(arrCount) = TargetCell
                        arrCount = arrCount + 1
                        ' メッセージ変数に追加
                        message = message & "番号:" & TargetCell.End(xlToLeft).Value & " - " & .Name & "にて機器番号: " & NumValue & " が見つかりませんでした。" & vbCrLf
                        Exit For
                    End If
                Next j
            ElseIf InStr(.Name, pipingSheet) <> 0 And NumValue = pipingDevNum Then ' 配管を含むシートかつ機器番号が配管の場合
                ' 対象の最初の行から最後の行まで検索をし、機器番号・機器名称をチェックする (配管)
                For j = firstRow To lastRow
                    devName = .Cells(j, devNameCol).Value ' 現在の行の機器名称の取得
                    If NameValue = devName Then  ' 機器名称が一致した場合
                        Set planMarkerCell = .Cells(j, yearColumn) ' 検索年度のマーカーのセルを取得
                        ' 検索年度のマーカーをチェックするプロシージャを呼ぶ
                            Dim thickSheet As Worksheet: Set thickSheet = inspectWb.Worksheets(thicknessSheet) ' 肉厚測定のシートの取得
                            Dim thicknessDevSection As String: thicknessDevSection = "C" ' 配管の区分の列記号を代入
                            Dim thickYearColumn As Long: thickYearColumn = findYearColumn(thickSheet, thicknessHeadRow, searchYearValue)

                            ' 検索対象の最初の行の取得
                            Dim thickfirstRow As Long: thickfirstRow = findRows(thickSheet, "first", searchDevValue, thicknessDevGroupCol)
                            ' 検索対象の最後の行の取得
                            Dim thicklastRow As Long: thicklastRow = findRows(thickSheet, "last", searchDevValue, thicknessDevGroupCol)

                            With thickSheet
                                Dim SearchName As String, count As Long, t As Long, thickMarkerCell As Range, thickResult As Long
                                count = 0
                                thickResult = 0
                                For t = thickfirstRow To thicklastRow
                                    SearchName = .Cells(t, thicknessDevSection).Value ' 現在の行の機器番号の取得
                                    If NameValue = SearchName Then    ' 機器番号が一致した場合
                                        Set thickMarkerCell = .Cells(t, thickYearColumn) ' 検索年度のマーカーのセルを取得
                                        ' 検索年度のマーカーをチェックするプロシージャを呼ぶ
                                        thickResult = thickResult + checkPlanMarkerThick(planMarkerCell, thickMarkerCell)
                                        count = count + 1
                                    ElseIf t = thicklastRow And NameValue <> SearchName And count = 0 Then
                                        ' エラーセルの配列に格納
                                        ReDim Preserve cellArray(arrCount)
                                        Set cellArray(arrCount) = TargetCell.Offset(0, 1)
                                        arrCount = arrCount + 1
                                        ' メッセージ変数に追加
                                        message = message & "番号:" & TargetCell.End(xlToLeft).Value & " - " & .Name & "にて配管区分: " & NameValue & " が見つからないため、配管周期表と比較できません。" & vbCrLf
                                    End If
                                Next t
                                If count <> 0 And thickResult = 0 Then
                                    Call checkPlanMarker(planMarkerCell, TargetCell, j)
                                ElseIf count <> 0 And thickResult <> 0 Then
                                        ' エラーセルの配列に格納
                                        ReDim Preserve cellArray(arrCount)
                                        Set cellArray(arrCount) = TargetCell.Offset(0, 1)
                                        arrCount = arrCount + 1
                                        ' メッセージ変数に追加
                                        message = message & "番号:" & TargetCell.End(xlToLeft).Value & " - " & .Name & "にて配管区分: 【" & NameValue & "】 の配管周期表と肉厚測定周期表で一致しない行があります。" & vbCrLf
                                End If
                            End With
                        Exit For
                    ElseIf j = lastRow Then
                        ' エラーセルの配列に格納
                        ReDim Preserve cellArray(arrCount)
                        Set cellArray(arrCount) = TargetCell.Offset(0, 1)
                        arrCount = arrCount + 1
                        ' メッセージ変数に追加
                                        message = message & "番号:" & TargetCell.End(xlToLeft).Value & " - " & .Name & "にて配管区分: " & NameValue & " が見つかりませんでした。" & vbCrLf
                        Exit For
                    End If
                Next j
            End If
            ' i (基準セルの行番号) セルが結合されているため、一度下のセルに移動してから一つ上のセルを取得する
            i = TargetCell.Offset(1).Row - 1
        Next i
    End With
    ' エラーメッセージ
    If message <> "" Then OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text + message
End Sub

' 指定したマーカーが含まれるかをチェック
Private Sub checkPlanMarker(markerCell As Range, TargetCell As Range, index As Long)
    Dim marker As String: marker = markerCell.Value
        If inspectWs.Name = openingSheet Then
            ' 開放実績の場合
            ' 記入欄に"●"が存在すれば、内部の目視検査にチェックがあるか確認する
            If InStr(marker, "●") > 0 Then
                Call checkMarkerChecked(TargetCell, 7, True, index)
            Else
                Call checkMarkerChecked(TargetCell, 3, False, index)
            End If
            ' 記入欄に"▲"または"◆"が存在すれば、肉厚測定以外の非破壊検査にチェックがあるか確認する
            If InStr(marker, "▲") > 0 Or InStr(marker, "◆") > 0 Then
                Call checkMarkerChecked(TargetCell, 25, True, index)
            Else
                Call checkMarkerChecked(TargetCell, 21, False, index)
            End If
            ' 記入欄に"溶"が存在すれば、溶接補修を行った場合の耐圧試験にチェックがあるか確認する
            If InStr(marker, "溶") > 0 Then
                Call checkMarkerChecked(TargetCell, 34, True, index)
            Else
                Call checkMarkerChecked(TargetCell, 30, False, index)
            End If
        ElseIf inspectWs.Name = thicknessSheet Then
            '肉厚測定の場合
            ' 記入欄に"●"が存在すれば、内部の目視検査にチェックがあるか確認する
            If InStr(marker, "●") > 0 Then
                Call checkMarkerChecked(TargetCell, 16, True, index)
            Else
                Call checkMarkerChecked(TargetCell, 12, False, index)
            End If
        ElseIf inspectWs.Name = pipingSheet Or inspectWs.Name = sharePipingSheet Then
            '配管・共有配管の場合
            If InStr(marker, "●") > 0 Then
                Call checkMarkerChecked(TargetCell, 7, True, index)
            Else
                Call checkMarkerChecked(TargetCell, 3, False, index)
            End If
            ' 記入欄に"▲"または"◆"が存在すれば、肉厚測定以外の非破壊検査にチェックがあるか確認する
            If InStr(marker, "▲") > 0 Or InStr(marker, "◆") > 0 Then
                Call checkMarkerChecked(TargetCell, 25, True, index)
            Else
                Call checkMarkerChecked(TargetCell, 21, False, index)
            End If
            ' 記入欄に"溶"が存在すれば、溶接補修を行った場合の耐圧試験にチェックがあるか確認する
            If InStr(marker, "溶") > 0 Then
                Call checkMarkerChecked(TargetCell, 34, True, index)
            Else
                Call checkMarkerChecked(TargetCell, 30, False, index)
            End If
            ' 記入欄に"●"が存在すれば、内部の目視検査にチェックがあるか確認する
            If InStr(markerCell.Offset(1, 0).Value, "■") > 0 Then
                Call checkMarkerChecked(TargetCell, 16, True, index)
            Else
                Call checkMarkerChecked(TargetCell, 12, False, index)
            End If
        End If
End Sub

' 配管シートと肉厚測定シートの比較
Private Function checkPlanMarkerThick(markerCell As Range, markerCellThick As Range) As Long
    checkPlanMarkerThick = 0 ' 関数戻り値の初期化
    Dim marker As String: marker = markerCell.Offset(1, 0).Value ' 非破壊検査マーカーの取得
    Dim markerThick As String: markerThick = markerCellThick.Value ' 肉厚測定のマーカーの取得
    '配管・共有配管の場合
    If InStr(marker, "■") > 0 Then
        If InStr(markerThick, "●") = 0 Then
            checkPlanMarkerThick = checkPlanMarkerThick + 1
        End If
    Else
        If InStr(markerThick, "●") > 0 Then
            checkPlanMarkerThick = checkPlanMarkerThick + 1
        End If
    End If
End Function

' マーカーに対応したチェックボックスがチェックされているか確認する
Private Sub checkMarkerChecked(TargetCell As Range, moveCol As Long, marker As Boolean, i As Long)
    Dim checkCell As Range: Set checkCell = TargetCell.Offset(1, moveCol).Offset(-8, 0)  ' 有のチェックボックス
    Dim bgColor As Long: bgColor = checkCell.Offset(1, 0).Interior.Color
    Dim cellMove As Long
    cellMove = IIf(marker = True, -4, 4)
    If checkCell.Value = ChrW(111) Then ' チェックされていない場合
        Range(checkCell, checkCell.Offset(0, 2)).Interior.Color = 6723891
        Range(checkCell.Offset(0, cellMove), checkCell.Offset(0, cellMove + 2)).Interior.Color = bgColor
        OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & "番号:" & TargetCell.End(xlToLeft).Value & " " & TargetCell.Value & " - " & TargetCell.Offset(0, 1).Value & " でエラーがあります。" & vbCrLf
    Else ' チェックされている場合
        Range(checkCell, checkCell.Offset(0, 2)).Interior.Color = bgColor
        Range(checkCell.Offset(0, cellMove), checkCell.Offset(0, cellMove + 2)).Interior.Color = bgColor
    End If
End Sub


' 配列になったセルを塗る
Public Sub fillDiffCells(Color As Long)
    thisWs.Range(thisWs.Columns(devicesColumn), thisWs.Columns(devicesColumn + 1)).Interior.ColorIndex = 0 ' 全ての背景色を消す
    If Not Not cellArray Then
        Dim i As Long
        ' エラーセルを指定の色に塗りつぶす
        For i = LBound(cellArray) To UBound(cellArray)
            If cellArray(i).Interior.Color <> Color Then
                cellArray(i).Interior.Color = Color
            End If
        Next i
    End If
End Sub


