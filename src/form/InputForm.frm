VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputForm 
   Caption         =   "���������̓t�H�[��"
   ClientHeight    =   6075
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
    StartUpPosition = 0
    Top = Application.Top + ((Application.Height - Height) / 2)
    Left = Application.Left + ((Application.Width - Width) / 2)

    ' �O���[�o���ϐ����錾����Ȃ����͍Đ錾
    If openingSheet = "" Then Call CheckDiff.varDeclaration
    ' �V�[�g�̓��e�����[�U�[�t�H�[���ɓ]�L
    With ActiveSheet
        ' ���u��
        Dim devices() As Variant
        Dim Find As String
        Dim Result As Variant
        devices = Array("1RF", "1TP", "1UF", "2UF", "3UF", "1HP", "HDS", "2TP", "4UF", "6EG", "FCC", "FGD", "10DDS", "20HP", "10HP", "20DDS", "2RF", "3PK", "NC")
        Find = .Range(searchDevCell).Value
        Result = Filter(devices, Find, True)
 
    If (UBound(Result) <> -1) Then
        Debug.Print Find & "���܂ޔz��͑��݂��܂��B"
    Else
        Debug.Print Find & "���܂ޔz��͑��݂��܂���B"
    End If

    
        DeviceComboBox.Text = .Range(searchDevCell).Value
        With DeviceComboBox
            For Each device In devices
                .AddItem device
            Next device
        End With
        ' �����N�x
        InsYearTextBox.Text = .Range(searchYearCell).Value
        ' ��������
        InsPeriodArr = Split(.Range(insPeriodCell).Value, "�`")
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
    ' ���[�U�[�t�H�[���̓��e��Excel�ɓ]�L
    Dim JudgeDate As String
    Dim lastRow As Long
    With ActiveSheet
        ' ���u��
        .Range(searchDevCell).Value = Me.DeviceComboBox.Text
        ' �����N�x

        If Len(Me.InsYearTextBox.Text) = 4 Then
            .Range(searchYearCell).Value = Me.InsYearTextBox.Text
        Else
            MsgBox "�����N�x��4���œ��͂��ĉ������B"
            Exit Sub
        End If

        ' ��������
        Dim OpInsPeriodVal As String
        If Len(Me.InsPeriodYearOpTextBox.Text) = 4 And Len(Me.InsPeriodMonthOpTextBox.Text) <= 2 And Len(Me.InsPeriodDayOpTextBox.Text) <= 2 Then
            OpInsPeriodVal = Me.InsPeriodYearOpTextBox.Text + "�N" + Me.InsPeriodMonthOpTextBox.Text + "��" + Me.InsPeriodDayOpTextBox.Text + "��"
        End If
        Dim EdInsPeriodVal As String
        If Len(Me.InsPeriodYearEdTextBox.Text) = 4 And Len(Me.InsPeriodMonthEdTextBox.Text) <= 2 And Len(Me.InsPeriodDayEdTextBox.Text) <= 2 Then
            EdInsPeriodVal = Me.InsPeriodYearEdTextBox.Text + "�N" + Me.InsPeriodMonthEdTextBox.Text + "��" + Me.InsPeriodDayEdTextBox.Text + "��"
        End If
        If OpInsPeriodVal <> "" And EdInsPeriodVal <> "" Then
            .Range(insPeriodCell).Value = OpInsPeriodVal + "�`" + EdInsPeriodVal
        Else
            MsgBox "�������Ԃ𐳂������͂��ĉ������B"
            Exit Sub
        End If
    End With
    Unload Me
End Sub
