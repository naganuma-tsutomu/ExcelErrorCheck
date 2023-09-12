VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputForm
   Caption         =   "���������̓t�H�[��"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "InputForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "InputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' �L�����Z���{�^���ŃE�C���h�E�����
Private Sub cancelButton_Click()
    Unload Me
End Sub

' �t�H�[�����J�����Ƃ�
Private Sub UserForm_Initialize()
    Dim InsPeriodArr As Variant
    Dim JudgeDateVal As String
    Dim count As Long
    ' �t�H�[���̕\���ʒu�𒆉��ɂ���
    Me.StartUpPosition = 0
    Me.Top = Application.Top + ((Application.Height - Me.Height) / 2)
    Me.Left = Application.Left + ((Application.Width - Me.Width) / 2)
    ' �O���[�o���ϐ����錾����Ȃ����͍Đ錾
    If bookName = "" Then Call CheckDiff.varDeclaration
    ' �V�[�g�̓��e�����[�U�[�t�H�[���ɓ]�L
    With ActiveSheet
        ' ���u��
        Me.DeviceTextBox.Text = .Range(searchDevCell).Value
        ' �����N�x
        Me.InsYearTextBox.Text = .Range(searchYearCell).Value
        ' ��������
        InsPeriodArr = Split(.Range(insPeriodCell).Value, "�`")
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
    ' ���[�U�[�t�H�[���̓��e��Excel�ɓ]�L
    Dim OpInsPeriodVal As String
    Dim EdInsPeriodVal As String
    Dim JudgeDate As String
    Dim lastRow As Long
    With ActiveSheet
        ' ���u��
        .Range(searchDevCell).Value = Me.DeviceTextBox.Text
        ' �����N�x
        .Range(searchYearCell).Value = Me.InsYearTextBox.Text
        ' ��������
        OpInsPeriodVal = Me.InsPeriodYearOpTextBox.Text + "�N" + Me.InsPeriodMonthOpTextBox.Text + "��" + Me.InsPeriodDayOpTextBox.Text + "��"
        EdInsPeriodVal = Me.InsPeriodYearEdTextBox.Text + "�N" + Me.InsPeriodMonthEdTextBox.Text + "��" + Me.InsPeriodDayEdTextBox.Text + "��"
        .Range(insPeriodCell).Value = OpInsPeriodVal + "�`" + EdInsPeriodVal
    End With
    Unload Me
End Sub
