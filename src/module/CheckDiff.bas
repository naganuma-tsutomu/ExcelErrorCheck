Attribute VB_Name = "CheckDiff"


' 変数データの作成
Sub varDeclaration()
    ' ----------計画ファイル用----------                        ' 対象ファイル名を入力(フル)
    directoryLevel = 2                           ' 対象ファイルがある階層数を入力 ([同階層 :0] [1つ上の階層 :1] [2つ上の階層 :2])
    ' 開放実績パラメータ
    openingSheet = "機器"                             ' シート番号を入力
    openingHeadRow = 22                          ' フィルターの基準となっている行を入力
    openingDevGroupCol = "B"                     ' 装置名の列記号を入力
    openingDevNumCol = "D"                       ' 機器番号の列記号を入力
    openingDevNameCol = "E"                      ' 機器名称の列記号を入力
    ' 肉厚測定パラメータ
    thicknessSheet = "肉厚測定"                           ' 肉厚測定計画表のシート番号を入力
    thicknessHeadRow = 22                        ' フィルターの基準となっている行を入力
    thicknessDevGroupCol = "B"                   ' 装置名の列記号を入力
    thicknessDevNumCol = "D"                     ' 機器番号の列記号を入力
    thicknessDevNameCol = "E"                    ' 機器名称の列記号を入力
    ' 配管パラメータ
    pipingSheet = "配管"                           ' 肉厚測定計画表のシート番号を入力
    pipingHeadRow = 4                        ' フィルターの基準となっている行を入力
    pipingDevNum = "配管"                      ' 配管用の機器番号
    pipingDevNameCol = "B"                    ' 機器名称の列記号を入力
    ' 共有配管パラメータ
    sharePipingSheet = "共有配管"                           ' 肉厚測定計画表のシート番号を入力
    sharePipingHeadRow = 4                        ' フィルターの基準となっている行を入力
    sharePipingDevNum = "配管"                      ' 配管用の機器番号
    sharePipingDevNameCol = "B"                    ' 機器名称の列記号を入力
    ' ----------記録ファイル用----------
    searchDevCell = "BF4"                         ' 装置名の入力セルを入力
    searchYearCell = "BF6"                        ' 検査年度の入力セルを入力
    insPeriodCell = "AN3"                        ' 検査期間の入力セルを入力
    devTopCell = 7
    judgeCell = 40                              ' 判定者の入力セルを入力
    judgeDateCell = 41                       ' 判定年月日の入力セルを入力
    searchDevMasterCol = 2                       ' 検索する機器番号の列番号を入力(基準とする)
End Sub

' 計画ファイルと記録ファイルとの差異をチェックする
Sub CheckDiffFunction()
    If (OutputForm.Visible = True) Then Unload OutputForm
    If openingSheet = "" Then Call varDeclaration
    Dim CheckDiff As New CheckDiffClass    ' クラスをインスタンス化
    ' 対象とするファイル・検索値を取得する
    With CheckDiff
        Call .initBook(directoryLevel, searchDevCell, searchYearCell, searchDevMasterCol)
        ' ファイルが存在すればチェックを行う
        If Not .inspectWb Is Nothing And .searchYearValue <> "" And .searchDevValue <> "" Then
            If InStr(ThisWorkbook.ActiveSheet.Name, "共有") = 0 Then
                ' -------------------------------- 検査対象1 (開放実績)
                ' 対象とするシートを指定する
                Call .initSheet(openingSheet)
                If Not .inspectWs Is Nothing Then ' シートが存在すれば
                    ' 計画ファイルからデータを検索し、記録ファイルとの差異をチェックする
                    Call .CheckInspection(openingHeadRow, openingDevGroupCol, openingDevNumCol, openingDevNameCol)
                Else
                    OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & sharePipingSheet & "シートが周期表ファイル内で見つかりませんでした。" & vbCrLf
                End If
                ' -------------------------------- 検査対象2 (肉厚測定)
                ' 対象とするシートを指定する
                Call .initSheet(thicknessSheet)
                If Not .inspectWs Is Nothing Then' シートが存在すれば
                    ' 計画ファイルからデータを検索し、記録ファイルとの差異をチェックする
                    Call .CheckInspection(thicknessHeadRow, thicknessDevGroupCol, thicknessDevNumCol, thicknessDevNameCol)
                Else
                    OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & sharePipingSheet & "シートが周期表ファイル内で見つかりませんでした。" & vbCrLf
                End If
                ' -------------------------------- 検査対象3 (配管)
                ' 対象とするシートを指定する
                Call .initSheet(pipingSheet)
                If Not .inspectWs Is Nothing Then' シートが存在すれば
                    ' 計画ファイルからデータを検索し、記録ファイルとの差異をチェックする
                    Call .CheckInspectionPiping(pipingHeadRow, thicknessDevGroupCol, pipingDevNum, pipingDevNameCol)
                Else
                    OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & sharePipingSheet & "シートが周期表ファイル内で見つかりませんでした。" & vbCrLf
                End If
            Else
                ' -------------------------------- 検査対象4 (共有配管)
                ' 対象とするシートを指定する
                Call .initSheet(sharePipingSheet)
                If Not .inspectWs Is Nothing Then' シートが存在すれば
                    ' 計画ファイルからデータを検索し、記録ファイルとの差異をチェックする
                    Call .CheckInspectionPiping(sharePipingHeadRow, thicknessDevGroupCol, sharePipingDevNum, sharePipingDevNameCol)
                Else
                    OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & sharePipingSheet & "シートが周期表ファイル内で見つかりませんでした。" & vbCrLf
                End If
            End If
            ' エラー箇所に色付けする
            Call .fillDiffCells(6723891)
            ' アラートの表示
            If OutputForm.ErrorTextBox.Text = "" Then
                 MsgBox "エラーは見つかりませんでした。"
            Else
                OutputForm.Show vbModeless
            End If
        End If
        ' 計画ファイルを開いていなかった場合、再度閉じる
        Call .CloseBook
    End With
End Sub

