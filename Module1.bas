Attribute VB_Name = "Module1"
Option Explicit
Sub CSV�Ǎ��N���A()
    With Sheets("�Ǎ�CSV�W�J")
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).ClearContents
    End With
    With Sheets("MENU")
        .Range("�Ǎ��ŉ��s").ClearContents
        .Range("�Ǎ��ŉE��").ClearContents
    End With
    MsgBox "�u�Ǎ�CSV�W�J�v�V�[�g�̓��e���N���A���܂���"
End Sub
Sub �ꗗ�����N���A()
    With Sheets("�ꗗ����")
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).ClearContents
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).Borders.LineStyle = False
    End With
    MsgBox "�u�ꗗ�����v�V�[�g�̓��e���N���A���܂���"
End Sub
Sub CSV�Ǎ�()
    Dim �n�� As Date, �I�� As Date
    Dim �t�@�C����
    Dim �S�� As String, �����R�[�h As String, ��ؕ��� As String, �A���Ǎ����[�h As String
    Dim ���o�s�� As Long, �Ǎ��s�� As Long, �Ǎ��� As Long, �I�s As Long
    Dim ����()
    �t�@�C���� = Application.GetOpenFilename(FileFilter:="CSV�t�@�C���i*.csv�j,*.csv", Title:="CSV�t�@�C���̑I��")
    If �t�@�C���� = False Then Exit Sub
    �n�� = Timer
    ���s��.Show vbModeless
    ���s��.Repaint
    With Sheets("MENU")
        �����R�[�h = .Range("�����R�[�h")
        ��ؕ��� = .Range("��ؕ���")
        ���o�s�� = .Range("�Ǎ����o�s��")
        �A���Ǎ����[�h = .Range("�A���Ǎ����[�h")
    End With
    With CreateObject("ADODB.Stream")
        .Charset = �����R�[�h
        .Open
        .LoadFromFile �t�@�C����
        �S�� = .ReadText
        .Close
    End With
        
    Call CSV���(�S��, ��ؕ���, ���o�s��, �Ǎ��s��, �Ǎ���, ����) '�O3�������3�Ԃ��̃C���[�W
        
    With Sheets("�Ǎ�CSV�W�J")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        If �I�s = 1 And .Cells(1, 1) = "" Then
            Range(.Cells(1, 1), .Cells(�Ǎ��s��, �Ǎ���)) = ����
            Else:
                Range(.Cells(�I�s + 1, 1), .Cells(�I�s + �Ǎ��s��, �Ǎ���)) = ����
                �Ǎ��s�� = �Ǎ��s�� + �I�s
        End If
    End With
    With Sheets("MENU")
        .Range("�Ǎ��ŉ��s") = �Ǎ��s��
        .Range("�Ǎ��ŉE��") = �Ǎ���
    End With
    �I�� = Timer
    If �A���Ǎ����[�h = "ON" Then
        If MsgBox("�t�@�C���̓ǂݍ��݂��������܂����B" & vbCrLf & vbCrLf & "�������ԁF" & �I�� - �n�� & vbCrLf & vbCrLf & "�ǉ��Ńt�@�C����ǂݍ��݂܂����H", vbYesNo) = vbYes Then
            Call CSV�Ǎ�
        End If
    End If
    Unload ���s��
End Sub
Sub CSV���(�S�� As String, ��ؕ��� As String, ���o�s�� As Long, �Ǎ��s�� As Long, �Ǎ��� As Long, ���� As Variant)
    Dim ���s�R�[�h As String
    Dim ���v�f�� As Long, �Y�� As Long, �� As Long, �J�n As Long, �I�� As Long, �s As Long, �� As Long
    Dim �͎��J�E���g As Long '�_�u���N�H�[�e�[�V�����̐������or���s���_�ŋ����Ȃ玟�̍��ֈڂ��ėǂ�
    Dim �͔��� As Long '���Y�������ꕶ���ň͂܂�Ă��邩�ǂ���(0 or 1)
    
    ���s�R�[�h = vbLf
    �S�� = Replace(�S��, vbCr, "") '���s�R�[�h��vbCrLf�������ꍇ�p�̕␳
    For �� = 1 To Len(�S��)
        Select Case Mid(�S��, ��, 1)
            Case ��ؕ���, ���s�R�[�h: ���v�f�� = ���v�f�� + 1 '���ۂ̍�����葽���Ȃ�ꍇ�����邽�߁u���v
        End Select
    Next
    
    ReDim �I�n(1 To ���v�f��, 1 To 3) '1��F���̍ŏ��̎��̔ԍ��A2��F�����̕�������Mid(�S��,1��,2��)�ō��f�[�^��o
    �Y�� = 1
    For �� = 1 To Len(�S��) '1��������́��I�n�z��֋L�^
        Select Case ��
            Case 1 '��������
                Select Case Mid(�S��, ��, 1) '�͔���
                    Case """"
                        �͔��� = 1
                        �͎��J�E���g = �͎��J�E���g + 1
                    Case Else: �͔��� = 0
                End Select
                �I�n(�Y��, 1) = �� + �͔��� '�n���ʒu�L�^
            Case Else '���������F��ؕ����܂��͉��s�R�[�h�P�ʂőO�����߁������J�n����
                Select Case Mid(�S��, ��, 1)
                    Case ��ؕ���, ���s�R�[�h
                        If �͔��� = 0 Or �͔��� = 1 And �͎��J�E���g Mod 2 = 0 Then '����؂�ƒf��\������
                            �I�n(�Y��, 2) = �� - �I�n(�Y��, 1) - �͔��� '�������L�^
                            If Mid(�S��, ��, 1) = ���s�R�[�h Then '���s�ʒu�L�^
                                �I�n(�Y��, 3) = 1
                                If �Ǎ��� = 0 Then �Ǎ��� = �Y��
                            End If
                            Select Case Mid(�S��, �� + 1, 1) '�����͔̈���
                                Case """": �͔��� = 1
                                Case Else: �͔��� = 0
                            End Select
                            If �� < Len(�S��) Then
                                �Y�� = �Y�� + 1 '�����ֈړ�
                                �I�n(�Y��, 1) = �� + 1 + �͔��� '�����̎n���ʒu�L�^
                            End If
                        End If
                    Case """": �͎��J�E���g = �͎��J�E���g + 1
                End Select
        End Select
    Next
    
    �Ǎ��s�� = �Y�� / �Ǎ���
    �I�� = �Ǎ��s�� * �Ǎ���
    ReDim ����(1 To �Ǎ��s��, 1 To �Ǎ���) 'CSV�����&�����������ʂ��L�^
    �J�n = 1
    If ���o�s�� > 0 Then
        If MsgBox("���o�s���܂߂ēǂݍ��݂܂����H", vbYesNo) = vbNo Then
            �J�n = �Ǎ��� * ���o�s�� + 1
            �Ǎ��s�� = �Ǎ��s�� - ���o�s��
        End If
    End If
    �s = 1
    �� = 1
    For �Y�� = �J�n To �I��
        ����(�s, ��) = Trim(Replace(Mid(�S��, �I�n(�Y��, 1), �I�n(�Y��, 2)), """""", """")) '�A���_�u���N�H�[�e�[�V�������������E�[�̋󔒍폜���L�^
        Select Case �I�n(�Y��, 3)
            Case 1
                �s = �s + 1
                �� = 1
            Case Else
                �� = �� + 1
        End Select
    Next
End Sub
Sub �ꗗ����()
    Dim �n�� As Date, �I�� As Date
    Dim �Ǎ��ŉ��s As Long, �Ǎ��ŉE�� As Long, ���o�s As Long, �ݒ�I�s As Long, �s As Long, �� As Long, �s�� As Long, �� As Long, �Y�� As Long
    Dim �ݒ�(), �f�[�^()
    With Sheets("MENU")
        �Ǎ��ŉ��s = .Range("�Ǎ��ŉ��s")
        �Ǎ��ŉE�� = .Range("�Ǎ��ŉE��")
        ���o�s = .Range("�Ǎ����o�s��")
        �ݒ�I�s = .Cells(Rows.Count, 7).End(xlUp).Row
        Select Case True
            Case �Ǎ��ŉ��s < 1, �Ǎ��ŉE�� < 1
                MsgBox "�Ǎ��f�[�^������܂���"
                Exit Sub
            Case �ݒ�I�s < 3
                MsgBox "�ꗗ�����ݒ�����Ă�������"
                Exit Sub
        End Select
        �ݒ� = Range(.Cells(3, 6), .Cells(�ݒ�I�s, 14))
        For �s = 3 To �ݒ�I�s
            If �� < .Cells(�s, 6) Then �� = .Cells(�s, 6)
        Next
    End With
    
    �n�� = Timer
    ���s��.Show vbModeless
    ���s��.Repaint
    With Sheets("�Ǎ�CSV�W�J")
        �f�[�^ = Range(.Cells(���o�s + 1, 1), .Cells(�Ǎ��ŉ��s, �Ǎ��ŉE��))
        �s�� = �Ǎ��ŉ��s - ���o�s
    End With
    With Sheets("�ꗗ����")
        ReDim ����(1 To �s�� + 1, 1 To ��)
        For �Y�� = 1 To �ݒ�I�s - 2
            ����(1, �ݒ�(�Y��, 1)) = �ݒ�(�Y��, 2)
            For �s = 1 To �s��
                If �ݒ�(�Y��, 3) <> "" Then ����(�s + 1, �ݒ�(�Y��, 1)) = �f�[�^(�s, �ݒ�(�Y��, 3))
                If �ݒ�(�Y��, 5) <> "" Then ����(�s + 1, �ݒ�(�Y��, 1)) = Trim(����(�s + 1, �ݒ�(�Y��, 1)) & �ݒ�(�Y��, 4) & �f�[�^(�s, �ݒ�(�Y��, 5)))
                If �ݒ�(�Y��, 7) <> "" Then ����(�s + 1, �ݒ�(�Y��, 1)) = Trim(����(�s + 1, �ݒ�(�Y��, 1)) & �ݒ�(�Y��, 6) & �f�[�^(�s, �ݒ�(�Y��, 7)))
                If �ݒ�(�Y��, 8) <> "" Then ����(�s + 1, �ݒ�(�Y��, 1)) = Format(����(�s + 1, �ݒ�(�Y��, 1)), �ݒ�(�Y��, 8))
            Next
            Select Case �ݒ�(�Y��, 9)
                Case "": .Columns(�ݒ�(�Y��, 1)).NumberFormatLocal = "G/�W��"
                Case Else: .Columns(�ݒ�(�Y��, 1)).NumberFormatLocal = �ݒ�(�Y��, 9)
            End Select
        Next
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).ClearContents
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).Borders.LineStyle = False
        Range(.Cells(1, 1), .Cells(�s�� + 1, ��)) = ����
        Range(.Cells(1, 1), .Cells(�s�� + 1, ��)).Borders.LineStyle = True
        For �� = 1 To ��
            .Columns(��).AutoFit
        Next
        .Activate
    End With
    �I�� = Timer
    MsgBox "�ꗗ�������������܂���" & vbCrLf & vbCrLf & "�������ԁF" & �I�� - �n��
    Unload ���s��
End Sub
