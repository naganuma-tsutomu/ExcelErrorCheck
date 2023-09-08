
' ワークブックを開く時のイベント
Private Sub Workbook_Open()
    ' txtに書いてある外部ライブラリを読み込み
    load_from_conf ".\..\..\システム設定\libdef.txt"
End Sub

' -------------------- モジュール読み込みに関する関数 --------------------
' 設定ファイルに書いてある外部ライブラリを読み込み
Sub load_from_conf(conf_path)
    ' 全モジュールを削除
    clear_modules
    ' 絶対パスに変換
    conf_path = abs_path(conf_path)
    If Dir(conf_path) = "" Then
        MsgBox "外部ライブラリ定義" & conf_path & "が存在しません。"
        Exit Sub
    End If

    ' 読み取り
    fp = FreeFile
    Open conf_path For Input As #fp
    Do Until EOF(fp)
        ' １行ずつ
        Line Input #fp, temp_str
        If Not Left(temp_str, 1) = "'" And Len(temp_str) > 0 Then
            module_path = abs_path(temp_str)
            If Dir(module_path) = "" Then
                ' エラー
                MsgBox "モジュール" & module_path & "は存在しません。"
                Exit Do
            Else
                ' モジュールとして取り込み
                include module_path
            End If
        End If
    Loop
    Close #fp
    ThisWorkbook.Save
End Sub

' あるモジュールを外部から読み込み
' パスが.で始まる場合は，相対パスと解釈される
Sub include(file_path)
    ' 絶対パスに変換
    file_path = abs_path(file_path)
    ' モジュールとして登録
    ThisWorkbook.VBProject.VBComponents.Import file_path
End Sub

' 全モジュールを初期化
Private Sub clear_modules()
    For Each component In ThisWorkbook.VBProject.VBComponents
        If component.Type = 1 Or component.Type = 2 Or component.Type = 3 Then
            ' この標準モジュールを削除
            ThisWorkbook.VBProject.VBComponents.Remove component
        End If
    Next component
End Sub

' ファイルパスを絶対パスに変換
Function abs_path(file_path)
    ' 絶対パスに変換
    If Left(file_path, 1) = "." Then
        file_path = ThisWorkbook.Path & Mid(file_path, 2, Len(file_path) - 1)
    End If
    abs_path = file_path
End Function
