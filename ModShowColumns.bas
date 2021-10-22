Attribute VB_Name = "ModShowColumns"
Option Explicit

'ShowColumns       �E�E�E���ꏊ�FFukamiAddins3.ModCell 
'CheckArray1D      �E�E�E���ꏊ�FFukamiAddins3.ModArray
'CheckArray1DStart1�E�E�E���ꏊ�FFukamiAddins3.ModArray



Public Sub ShowColumns(ColumnABCList1D, TargetSheet As Worksheet, Optional ByVal MaxColABC As String, Optional InputShow As Boolean = True)
'�w���̂ݕ\���ɂ���
'20210917

'����
'ColumnABCList�E�E�E��\���Ώۂ̗񖼂�1�����z�� ��) ("A","B","C")
'TargetSheet  �E�E�E�Ώۂ̃V�[�g
'MaxColABC    �E�E�E��\���֑ؑΏۂ̗�͈͂̍ő��
'InputShow    �E�E�E�w�ߗ��\���Ȃ�True,��\���Ȃ�False�B�f�t�H���g��True
                                                                 
    '�����`�F�b�N
    Call CheckArray1D(ColumnABCList1D, "ColumnABCList1D")
    Call CheckArray1DStart1(ColumnABCList1D, "ColumnABCList1D")
    
    If MaxColABC = "" Then '��\���֑ؑΏۂ̗�͈͂̍ő�񂪎w�肳��Ă��Ȃ��ꍇ�̓V�[�g�̍ŏI��
        MaxColABC = Split(Cells(1, Columns.Count).Address(True, False), "$")(0) '�ŏI��ԍ��̃A���t�@�x�b�g�擾
    End If
    
    Dim I          As Long
    Dim N          As Long
    Dim ColumnName As String    '�\���Ώۂ̗񖼂��܂Ƃ߂�����
    N = UBound(ColumnABCList1D) '�Ώۂ̗�̌�
    ColumnName = ""             '�񖼂܂Ƃ߂̏�����
    For I = 1 To N
        ColumnName = ColumnName & ColumnABCList1D(I) & ":" & ColumnABCList1D(I)
        If I < N Then '�񖼂̍Ōゾ��","�����Ȃ�
            ColumnName = ColumnName & ","
        End If
    Next I
    
    Dim TargetCell As Range                        '�Ώ۔͈͂̃Z���I�u�W�F�N�g
    Set TargetCell = TargetSheet.Range(ColumnName) '�Ώ۔͈͂��Z���I�u�W�F�N�g�Ŏ擾
                                                                                    
    Application.ScreenUpdating = False             '��ʍX�V���������č�����
    
    If InputShow = True Then                                 '�\���ɐ؂�ւ��邩�A��\���ɐ؂�ւ��邩
        TargetSheet.Columns("A:" & MaxColABC).Hidden = True  '�S�̂��\��
        TargetCell.EntireColumn.Hidden = False               '�w�ߗ�̂ݕ\������
    Else
        TargetSheet.Columns("A:" & MaxColABC).Hidden = False '�S�̂��\��
        TargetCell.EntireColumn.Hidden = True                '�w�ߗ�̂ݕ\������
    End If
    
    ActiveWindow.ScrollColumn = 1     '��ԍ��̗�ɃX�N���[�����ĕ\������
    Application.ScreenUpdating = True '��ʍX�V�����̉���
    
End Sub

Private Sub CheckArray1D(InputArray, Optional HairetuName As String = "�z��")
'���͔z��1�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy As Integer
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "��1�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName As String = "�z��")
'����1�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub


