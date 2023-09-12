VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputForm
   Caption         =   "検査情報入力フォーム"
   ClientHeight    =   5475
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
    Me.StartUpPosition = 0
    Me.Top = Application.Top + ((Application.Height - Me.Height) / 2)
    Me.Left = Application.Left + ((Application.Width - Me.Width) / 2)
    ' グローバル変数が宣言されない時は再宣言
    If bookName = "" Then Call CheckDiff.varDeclaration
    ' シートの内容をユーザーフォームに転記
    With ActiveSheet
        ' 装置名
        Me.DeviceTextBox.Text = .Range(searchDevCell).Value
        ' 検査年度
        Me.InsYearTextBox.Text = .Range(searchYearCell).Value
        ' 検査期間
        InsPeriodArr = Split(.Range(insPeriodCell).Value, "～")
        count = UBound(InsPeriodArr) - LBound(InsPeriodArr) + 1
        If count >= 2 Then
            Me.InsPeriodYearOpTextBox.Text = DatePart("yyyy", InsPeriodArr(0))
            Me.InsPeriodMonthOpTextBox.Text = DatePart("m", InsPeriodArr(0))
            Me.InsPeriodDayOpTextBox.Text = DatePart("d", InsPeriodArr(0))
            Me.InsPeriodYearEdTextBox.Text = DatePart("yyyy", InsPeriodArr(1))
            Me.InsPeriodMonthEdTextBox.Text = DatePart("m", InsPeriodArr(1))
            Me.InsPeriodDayEdTextBox.Text = DatePart("d", InsPeriodArr(1))
        End If
    End With
    DeviceTextBox.SetFocus
End Sub

Private Sub inputButton_Click()
    ' ユーザーフォームの内容をExcelに転記
    Dim OpInsPeriodVal As String
    Dim EdInsPeriodVal As String
    Dim JudgeDate As String
    Dim lastRow As Long
    With ActiveSheet
        ' 装置名
        .Range(searchDevCell).Value = Me.DeviceTextBox.Text
        ' 検査年度
        .Range(searchYearCell).Value = Me.InsYearTextBox.Text
        ' 検査期間
        OpInsPeriodVal = Me.InsPeriodYearOpTextBox.Text + "年" + Me.InsPeriodMonthOpTextBox.Text + "月" + Me.InsPeriodDayOpTextBox.Text + "日"
        EdInsPeriodVal = Me.InsPeriodYearEdTextBox.Text + "年" + Me.InsPeriodMonthEdTextBox.Text + "月" + Me.InsPeriodDayEdTextBox.Text + "日"
        .Range(insPeriodCell).Value = OpInsPeriodVal + "～" + EdInsPeriodVal
    End With
    Unload Me
End Sub
