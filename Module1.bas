Attribute VB_Name = "Module1"
Option Explicit

'/**
' * UiPath����R�[������邱�Ƃ�O��Ƃ���Excel VBA�t���[�����[�N�B
' * �{�����͕K���W�����W���[���ɑ΂��ăC���|�[�g���Ďg���ĉ������B
' *
' * �K�v�ɉ����Ĉ�����ǉ�����B
' * ���������͖̂{�֐��ł͂Ȃ��AMain() �ɋL�ڂ��邱�ƁB
' *
' * ���Q�l�FUiPath�i���b�W�x�[�X�FVBA Macro�̎������Ή����̃G���[�n���h�����O�iOnErrorGoto�j
' * https://www.uipath.com/ja/resources/knowledge-base/vba-macro-onerrorgoto
' *
' * ���Q�l�FQiita�F�^�C�g��
' * URL
' *
' * @param {�^} �ϐ��� - 1�߂̈���
' * @param {�^} �ϐ��� - 2�߂̈���
' * @return {String} - ���펞�͋�, �ُ펞�̓G���[���e��Ԃ�
' */
Public Function UiPathVBAFramework() As String
    ' �G���[�g���b�v
    On Error GoTo ERR_CATCH_

    Dim Ret As String
    Ret = ""    ' �ԋp�l��������

    Application.ScreenUpdating = False  ' ��ʍX�VOFF
    Application.DisplayAlerts = False   ' �x���\��OFF

    ' ���C�������� Main() �ɋL��
    Call Main

' �I������
EXIT_:
    Application.ScreenUpdating = False  ' ��ʍX�VON
    Application.DisplayAlerts = False   ' �x���\��ON

    UiPathVBAFramework = Ret

    On Error GoTo 0
    Exit Function

' --------------------------
' ----- �G���[�L���b�` -----
' --------------------------
ERR_CATCH_:
    ' �G���[���e��ԋp�l�ɕێ�
    Ret = "Excel�}�N�����ŃG���[���������܂����B" & vbCrLf & _
          "�G���[�ԍ��F" & Err.Number & vbCrLf & _
          "�G���[���e�F" & Err.Description
    Err.Clear

    ' �I��������
    Resume EXIT_
End Function

'/**
' * ���������L�ڂ���֐��B
' * �K�v�ɉ����Ĉ�����ǉ�����B
' *
' * @param {�^} �ϐ��� - 1�߂̈���
' * @param {�^} �ϐ��� - 2�߂̈���
' */
Private Sub Main()
    ' ���������L��
End Sub
