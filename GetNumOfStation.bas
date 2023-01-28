Attribute VB_Name = "GetNumOfStation"
Option Explicit
Sub GetNumOfStation()
    
    '�t�@�C�����ɑ��Z��or�v�Z�ǂƂ����������܂܂��ꍇ�A�s���ɂ����̐����𑫂��Ă����ǐ����Z�o����}�N��
    
    Dim wbDst As Workbook       '���ʏo�͐�u�b�N�̃��[�N�u�b�N
    Dim wsDst As Worksheet      '���ʏo�͐�u�b�N�̃��[�N�V�[�g
    Dim wbOrg As Workbook       '���o���u�b�N�̃��[�N�u�b�N
    Dim wsOrg As Worksheet      '���o���u�b�N�̃��[�N�V�[�g
    Dim NameOrg As Variant      '�����l�͈́i�t�@�C�����j
    Dim StartRow As Long        '�t�B���^�[��̈�ԏ�̍s
    Dim EndRow As Long          '�t�B���^�[��̈�ԉ��̍s
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim Num As Long             '���Z�ǁA�v���ǂ̊ۂɓ��鐔�l
    Dim Sum As Long             '�ǐ����v
    Dim CellNum As Long         '�f�[�^��
    
    Set wbDst = Workbooks("�f�[�^���͂܂Ƃ�.xlsm")
    Set wbOrg = Workbooks(wbDst.Worksheets("�}�N��").Range("D3").Value)
    Set wbOrg = wbOrg.Worksheets("CHK�ō� DB")
    
    With wsOrg
    
        '�t�B���^�[�K�p��̈�ԏ�̍s�ԍ��擾
        StartRow = 0
    
        For i = 5 To .Cells(.Rows.Count, 5).End(xlUp).Row
            If .Cells(i, 5).EntireRow.Hidden = False Then
                If StartRow = 0 Then
                    StartRow = .Cells(i, 5).Row
                    Exit For
                End If
            End If
        Next
        
        '�f�[�^���擾
        For j = 5 To .Cells(.Rows.Count, 5).End(xlUp).Row
            If .Cells(j, 5).EntireRow.Hidden = False Then
                CellNum = CellNum + 1
            End If
        Next
        
        '�ǐ��擾
        NameOrg = .Range(.Cells(StartRow, 5), .Cells(.Cells(.Rows.Count, 5).End(xlUp).Row, 5))
        
        Sum = 0
        
        For k = LBound(NameOrg, 1) To UBound(NameOrg, 1)                                                  '�t�B���^�[��\������Ă��邷�ׂĂ̍s�����Ă���
            If .Cells(StartRow + k - 1, 1).EntireRow.Hidden = False Then
                If InStr(NameOrg(k, 1), "��") > 0 Then                                                    '�t�@�C������"��"�Ƃ����������܂�ł�����
                    If IsNumeric(Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "��") + 1, 1)) Then              '"��"�̎��̕����������Ȃ�
                        If IsNumeric(Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "��") + 1, 2)) Then          '"��"�̎��̎��̕����������Ȃ炻���̓񌅂̐��������v�ɑ���
                            Num = Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "��") + 1, 2)
                            Sum = Sum + Num
                        Else                                                                              '"��"�̎��̎��̕����������łȂ����"��"�̎��̈ꌅ�̐��������v�ɑ���
                            Num = Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "��") + 1, 1)
                            Sum = Sum + Num
                            Sum = Sum + Num
                        End If
                    End If
                ElseIf InStr(NameOrg(k, 1), "�v") > O Then                                                '�t�@�C������"�v"�Ƃ����������܂�ł�����
                    If IsNumeric(Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "�v") + 1, 1)) Then              '"�v"�̎��̕����������Ȃ�
                        If IsNumeric(Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "�v") + 1, 2)) Then          '"�v"�̎��̎��̕����������Ȃ炻���̓񌅂̐���-1�����v�ɑ���
                            Num = Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "�v") + 1, 2)
                            Sum = Sum + Num - 1
                        Else                                                                              '"�v"�̎��̎��̕����������łȂ����"�v"�̎��̈ꌅ�̐���-1�����v�ɑ���
                            Num = Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "�v") + 1, 1)
                            Sum = Sum + Num - 1
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    wsDst.Range("M8") = CellNum + Sum
    
    MsgBox "�����������܂����B"
        
    
End Sub
