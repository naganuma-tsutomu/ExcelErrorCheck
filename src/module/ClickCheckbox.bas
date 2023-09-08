Attribute VB_Name = "ClickCheckbox"
Option Explicit

'図形のクリックでチェックボックスの切り替え

Sub boxOnOff()

    Dim Target As Range

    '図の左上隅が属するセルを取得
    With ActiveSheet.Shapes(Application.Caller)
        Set Target = Range(.TopLeftCell, .TopLeftCell)
    End With

    If Target.MergeCells Then Exit Sub

    Dim LabelCell As Range
    Dim checkBoxCell As Range
    Dim Cancel As Boolean

    If Not Intersect(Target, Range("G:G, K:K, P:P, T:T, Y:Y, AC:AC, AH:AH, AL:AL, AS:AS, AY:AY")) Is Nothing Then
        Cancel = True ' セルを編集状態にしないようにする

        Set checkBoxCell = Target.Offset(0, -2)
        Set LabelCell = Target.Cells
        Call checkedCellValue(checkBoxCell, LabelCell)

    ElseIf Not Intersect(Target, Range("E:E, I:I, N:N, R:R, W:W, AA:AA, AF:AF, AJ:AJ, AQ:AQ, AU:AU, AW:AW")) Is Nothing Then
        Cancel = True ' セルを編集状態にしないようにする

        Set checkBoxCell = Target.Cells
        Set LabelCell = Target.Offset(0, 2)
        Call checkedCellValue(checkBoxCell, LabelCell)

    End If
End Sub

' チェックボックスにチェックする
Private Sub checkedCellValue(CheckBox As Range, Label As Range)
    If CheckBox.Value = ChrW(254) Then
        CheckBox.Value = ChrW(111)
    Else
        CheckBox.Value = ChrW(254)
        Call checkedCellFunc(CheckBox, Label)
    End If
End Sub

' チェックした際にもう一方のチェックボックスのチェックを外す
Private Sub checkedCellFunc(CheckBox As Range, Label As Range)
    Dim checkBoxCell As Range
    ' どちらのチェックボックスか判定
    If Not Intersect(CheckBox, Range("E:E, N:N, W:W, AF:AF")) Is Nothing Then
        ' 隣のセル（右側）を取得
        Set checkBoxCell = CheckBox.Offset(0, 4)
    ElseIf Not Intersect(CheckBox, Range("I:I, R:R, AA:AA, AJ:AJ")) Is Nothing Then
        ' 隣のセル（左側）を取得
        Set checkBoxCell = CheckBox.Offset(0, -4)
    ElseIf Not Intersect(CheckBox, Range("AQ:AQ")) Is Nothing Then
        If InStr(Label.Value, "不") > 0 Then
            Set checkBoxCell = CheckBox.Offset(-3, 0)
        Else
            Set checkBoxCell = CheckBox.Offset(3, 0)
        End If
    ElseIf Not Intersect(CheckBox, Range("AU:AU")) Is Nothing Then
        If InStr(Label.Value, "停止") > 0 Then
            Set checkBoxCell = CheckBox.Offset(-2, 0)
        Else
            Set checkBoxCell = CheckBox.Offset(2, 0)
        End If
    ElseIf Not Intersect(CheckBox, Range("AW:AW")) Is Nothing Then
        If InStr(Label.Value, "無") > 0 Then
            Set checkBoxCell = CheckBox.Offset(-5, 0)
        Else
            Set checkBoxCell = CheckBox.Offset(5, 0)
        End If
    End If
    If Not checkBoxCell Is Nothing Then checkBoxCell.Value = ChrW(111)
End Sub
