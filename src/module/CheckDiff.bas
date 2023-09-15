Attribute VB_Name = "CheckDiff"


' �ϐ��f�[�^�̍쐬
Sub varDeclaration()
    ' ----------�v��t�@�C���p----------                        ' �Ώۃt�@�C���������(�t��)
    directoryLevel = 2                           ' �Ώۃt�@�C��������K�w������� ([���K�w :0] [1��̊K�w :1] [2��̊K�w :2])
    ' �J�����уp�����[�^
    openingSheet = "�@��"                             ' �V�[�g�ԍ������
    openingHeadRow = 22                          ' �t�B���^�[�̊�ƂȂ��Ă���s�����
    openingDevGroupCol = "B"                     ' ���u���̗�L�������
    openingDevNumCol = "D"                       ' �@��ԍ��̗�L�������
    openingDevNameCol = "E"                      ' �@�햼�̗̂�L�������
    ' ��������p�����[�^
    thicknessSheet = "��������"                           ' ��������v��\�̃V�[�g�ԍ������
    thicknessHeadRow = 22                        ' �t�B���^�[�̊�ƂȂ��Ă���s�����
    thicknessDevGroupCol = "B"                   ' ���u���̗�L�������
    thicknessDevNumCol = "D"                     ' �@��ԍ��̗�L�������
    thicknessDevNameCol = "E"                    ' �@�햼�̗̂�L�������
    ' �z�ǃp�����[�^
    pipingSheet = "�z��"                           ' ��������v��\�̃V�[�g�ԍ������
    pipingHeadRow = 4                        ' �t�B���^�[�̊�ƂȂ��Ă���s�����
    pipingDevNum = "�z��"                      ' �z�Ǘp�̋@��ԍ�
    pipingDevNameCol = "B"                    ' �@�햼�̗̂�L�������
    ' ���L�z�ǃp�����[�^
    sharePipingSheet = "���L�z��"                           ' ��������v��\�̃V�[�g�ԍ������
    sharePipingHeadRow = 4                        ' �t�B���^�[�̊�ƂȂ��Ă���s�����
    sharePipingDevNum = "�z��"                      ' �z�Ǘp�̋@��ԍ�
    sharePipingDevNameCol = "B"                    ' �@�햼�̗̂�L�������
    ' ----------�L�^�t�@�C���p----------
    searchDevCell = "BF4"                         ' ���u���̓��̓Z�������
    searchYearCell = "BF6"                        ' �����N�x�̓��̓Z�������
    insPeriodCell = "AN3"                        ' �������Ԃ̓��̓Z�������
    devTopCell = 7
    judgeCell = 40                              ' ����҂̓��̓Z�������
    judgeDateCell = 41                       ' ����N�����̓��̓Z�������
    searchDevMasterCol = 2                       ' ��������@��ԍ��̗�ԍ������(��Ƃ���)
End Sub

' �v��t�@�C���ƋL�^�t�@�C���Ƃ̍��ق��`�F�b�N����
Sub CheckDiffFunction()
    If (OutputForm.Visible = True) Then Unload OutputForm
    If openingSheet = "" Then Call varDeclaration
    Dim CheckDiff As New CheckDiffClass    ' �N���X���C���X�^���X��
    ' �ΏۂƂ���t�@�C���E�����l���擾����
    With CheckDiff
        Call .initBook(directoryLevel, searchDevCell, searchYearCell, searchDevMasterCol)
        ' �t�@�C�������݂���΃`�F�b�N���s��
        If Not .inspectWb Is Nothing And .searchYearValue <> "" And .searchDevValue <> "" Then
            If InStr(ThisWorkbook.ActiveSheet.Name, "���L") = 0 Then
                ' -------------------------------- �����Ώ�1 (�J������)
                ' �ΏۂƂ���V�[�g���w�肷��
                Call .initSheet(openingSheet)
                If Not .inspectWs Is Nothing Then ' �V�[�g�����݂����
                    ' �v��t�@�C������f�[�^���������A�L�^�t�@�C���Ƃ̍��ق��`�F�b�N����
                    Call .CheckInspection(openingHeadRow, openingDevGroupCol, openingDevNumCol, openingDevNameCol)
                Else
                    OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & sharePipingSheet & "�V�[�g�������\�t�@�C�����Ō�����܂���ł����B" & vbCrLf
                End If
                ' -------------------------------- �����Ώ�2 (��������)
                ' �ΏۂƂ���V�[�g���w�肷��
                Call .initSheet(thicknessSheet)
                If Not .inspectWs Is Nothing Then' �V�[�g�����݂����
                    ' �v��t�@�C������f�[�^���������A�L�^�t�@�C���Ƃ̍��ق��`�F�b�N����
                    Call .CheckInspection(thicknessHeadRow, thicknessDevGroupCol, thicknessDevNumCol, thicknessDevNameCol)
                Else
                    OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & sharePipingSheet & "�V�[�g�������\�t�@�C�����Ō�����܂���ł����B" & vbCrLf
                End If
                ' -------------------------------- �����Ώ�3 (�z��)
                ' �ΏۂƂ���V�[�g���w�肷��
                Call .initSheet(pipingSheet)
                If Not .inspectWs Is Nothing Then' �V�[�g�����݂����
                    ' �v��t�@�C������f�[�^���������A�L�^�t�@�C���Ƃ̍��ق��`�F�b�N����
                    Call .CheckInspectionPiping(pipingHeadRow, thicknessDevGroupCol, pipingDevNum, pipingDevNameCol)
                Else
                    OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & sharePipingSheet & "�V�[�g�������\�t�@�C�����Ō�����܂���ł����B" & vbCrLf
                End If
            Else
                ' -------------------------------- �����Ώ�4 (���L�z��)
                ' �ΏۂƂ���V�[�g���w�肷��
                Call .initSheet(sharePipingSheet)
                If Not .inspectWs Is Nothing Then' �V�[�g�����݂����
                    ' �v��t�@�C������f�[�^���������A�L�^�t�@�C���Ƃ̍��ق��`�F�b�N����
                    Call .CheckInspectionPiping(sharePipingHeadRow, thicknessDevGroupCol, sharePipingDevNum, sharePipingDevNameCol)
                Else
                    OutputForm.ErrorTextBox.Text = OutputForm.ErrorTextBox.Text & sharePipingSheet & "�V�[�g�������\�t�@�C�����Ō�����܂���ł����B" & vbCrLf
                End If
            End If
            ' �G���[�ӏ��ɐF�t������
            Call .fillDiffCells(6723891)
            ' �A���[�g�̕\��
            If OutputForm.ErrorTextBox.Text = "" Then
                 MsgBox "�G���[�͌�����܂���ł����B"
            Else
                OutputForm.Show vbModeless
            End If
        End If
        ' �v��t�@�C�����J���Ă��Ȃ������ꍇ�A�ēx����
        Call .CloseBook
    End With
End Sub

