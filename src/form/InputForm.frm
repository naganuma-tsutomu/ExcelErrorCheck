VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputForm 
   Caption         =   "検査情報入力フォーム"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "InputForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "InputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' キャンセルボタンでウインドウを閉じる
Private Sub cancelButton_Click()
    Unload Me
End Sub

' フォームを開いたとき
Private Sub UserForm_Initialize()
    Dim InsPeriodArr As Variant
    Dim JudgeDateVal As String
    Dim count As Long
    ' フォームの表示位置を中央にする
    StartUpPosition = 0
    Top = Application.Top + ((Application.Height - Height) / 2)
    Left = Application.Left + ((Application.Width - Width) / 2)

    ' グローバル変数が宣言されない時は再宣言
    If openingSheet = "" Then Call CheckDiff.varDeclaration
    ' シートの内容をユーザーフォームに転記
    With ActiveSheet
        ' 装置名
        Dim devices() As Variant
        Dim Find As String
        Dim Result As Variant
        devices = Array("1RF", "1TP", "1UF", "2UF", "3UF", "1HP", "HDS", "2TP", "4UF", "6EG", "FCC", "FGD", "10DDS", "20HP", "10HP", "20DDS", "2RF", "3PK", "NC")
        Find = .Range(searchDevCell).Value
        Result = Filter(devices, Find, True)
 
    If (UBound(Result) <> -1) Then
        Debug.Print Find & "を含む配列は存在します。"
    Else
        Debug.Print Find & "を含む配列は存在しません。"
    End If

    
        DeviceComboBox.Text = .Range(searchDevCell).Value
        With DeviceComboBox
            For Each device In devices
                .AddItem device
            Next device
        End With
        ' 検査年度
        InsYearTextBox.Text = .Range(searchYearCell).Value
        ' 検査期間
        InsPeriodArr = Split(.Range(insPeriodCell).Value, "～")
        count = UBound(InsPeriodArr) - LBound(InsPeriodArr) + 1
        If count >= 2 Then
            InsPeriodYearOpTextBox.Text = DatePart("yyyy", InsPeriodArr(0))
            InsPeriodMonthOpTextBox.Text = DatePart("m", InsPeriodArr(0))
            InsPeriodDayOpTextBox.Text = DatePart("d", InsPeriodArr(0))
            InsPeriodYearEdTextBox.Text = DatePart("yyyy", InsPeriodArr(1))
            InsPeriodMonthEdTextBox.Text = DatePart("m", InsPeriodArr(1))
            InsPeriodDayEdTextBox.Text = DatePart("d", InsPeriodArr(1))
        End If
    End With
    DeviceComboBox.SetFocus
End Sub

Private Sub inputButton_Click()
    ' ユーザーフォームの内容をExcelに転記
    Dim JudgeDate As String
    Dim lastRow As Long
    With ActiveSheet
        ' 装置名
        .Range(searchDevCell).Value = Me.DeviceComboBox.Text
        ' 検査年度

        If Len(Me.InsYearTextBox.Text) = 4 Then
            .Range(searchYearCell).Value = Me.InsYearTextBox.Text
        Else
            MsgBox "検査年度は4桁で入力して下さい。"
            Exit Sub
        End If

        ' 検査期間
        Dim OpInsPeriodVal As String
        If Len(Me.InsPeriodYearOpTextBox.Text) = 4 And Len(Me.InsPeriodMonthOpTextBox.Text) <= 2 And Len(Me.InsPeriodDayOpTextBox.Text) <= 2 Then
            OpInsPeriodVal = Me.InsPeriodYearOpTextBox.Text + "年" + Me.InsPeriodMonthOpTextBox.Text + "月" + Me.InsPeriodDayOpTextBox.Text + "日"
        End If
        Dim EdInsPeriodVal As String
        If Len(Me.InsPeriodYearEdTextBox.Text) = 4 And Len(Me.InsPeriodMonthEdTextBox.Text) <= 2 And Len(Me.InsPeriodDayEdTextBox.Text) <= 2 Then
            EdInsPeriodVal = Me.InsPeriodYearEdTextBox.Text + "年" + Me.InsPeriodMonthEdTextBox.Text + "月" + Me.InsPeriodDayEdTextBox.Text + "日"
        End If
        If OpInsPeriodVal <> "" And EdInsPeriodVal <> "" Then
            .Range(insPeriodCell).Value = OpInsPeriodVal + "～" + EdInsPeriodVal
        Else
            MsgBox "検査期間を正しく入力して下さい。"
            Exit Sub
        End If
    End With
    Unload Me
End Sub
