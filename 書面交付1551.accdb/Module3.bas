Attribute VB_Name = "Module3"
Option Compare Database
Option Base 1


'�Ώ��������ɐݒu���𔲂��o����
'�Ώ�������������ꍇ��
'�E�H���\��������ׂē������H�N�G��
'    �[���ԍ��ɓ������̂����邩�iYES�F�������̂ŁA�����������ݒu���BNO�F�Ώ����͌��ݒu���j
'
'
'�ǂ��炩�ɍH�����i�ݒu���j�������Ă���ꍇ�A�Ώۏ��i���̂�fld_a="��"�ǉ�
'[2022/06]
'
'�E�ݒu�������i�V�K�j
'�E�T�[�r�X�J�n���i�ڍs�Ȃǁj
'
'�R�[�X�ύX�`�F�b�N
'�E�T�[�r�X�J�n���ɕύX�����邩
'
'�T�[�r�X�I�����`�F�b�N
'�E����i�ΏۂƂȂ�T�[�r�X�̓L�����y�[�����ƂȂ�j
'�E�Ȃ�
'
'�_�񐬗���
'�E���i���̖̂������f���͏Ȃ�
'�Efld_a="��"�̌_��\�������Q��
'
'�\���ԍ�





Sub test()
'�Ώ�������������l���T�[�`
'�ڋq�ԍ��ŃO���[�v���A�Ԃ�U��iDICTIONARY�֐����g�p�j
'�H���\������������邩�i�S�ē����Ȃ�Ώ����͈�ԑ傫�������j�Ⴄ�Ȃ�ELSE
'
Dim db As Database
Dim rs As Recordset
Dim rs_count As Integer

Set db = CurrentDb()
Set rs = db.OpenRecordset("Q_TA_H_K")

rs.MoveLast
rs.MoveFirst
rs_count = rs.RecordCount

Dim Myarray
Myarray = rs.GetRows(rs_count) '���R�[�h��2�����z��

'�s�����ւ�
ReDim Myarray2(UBound(Myarray, 2) + 1, UBound(Myarray, 1) + 2)
For i = 1 To UBound(Myarray, 1) '���񂩁H
    For J = 1 To UBound(Myarray, 2)
        Myarray2(J, i) = Myarray(i - 1, J)
    Next
Next
    

    
'Dictionary�̃I�u�W�F�N�g���쐬
Dim a
Set a = CreateObject("Scripting.Dictionary")

'�u�d���v��ۑ�����z��
Dim Data
ReDim Data(1 To UBound(Myarray2, 1), 1 To 1)
'ReDim Preserve Myarray(LBound(Myarray, 1), UBound(Myarray) + 1)
'�u���i�v�̗�����[�v
For i = 1 To UBound(Myarray2, 1)

'�܂��o�^����Ă��Ȃ��ꍇ

    If a.exists(Myarray2(i, 1)) = False Then
        a.Add Myarray2(i, 1), 1 '�o�^����i1�͂Ȃ�ł������j
        s = 1
        Debug.Print Myarray2(i, 1) & ":" & s
        Myarray2(i, 32) = s
    '���ɓo�^����Ă���ꍇ
    Else '(26,i)(27,i)�g�p
        s = s + 1
        Debug.Print Myarray2(i, 1) & ":" & s
        Myarray2(i, 32) = s
        
    End If
            
Next



Debug.Print �񎟌��z��t�B���^�[�֐�(Myarray2, 1, 105592501)

For i = 0 To UBound(Myarray2, 1) - 1
    Do While Myarray2(i, 32) < Myarray2(i + 1, 32)
        Debug.Print Myarray2(i, 1)
        i = i + 1
    Loop
Next

Dim Taisyo
Dim ss As Boolean
ss = True 'ON
For i = 0 To rs_count - 1
    '27��1�ŏ������J�n����
    If Myarray(27, i) = 1 Then '�Ώ���CK
        '�z��Ɋi�[
        
        
        ss = False 'OFF
    Else
        
    End If
Next i

End Sub

Sub tesst()

Dim test(3)

test(0) = "a"
test(1) = "a"
test(2) = "b"
test(3) = "a"

Debug.Print DeleteSameValue1(test)

End Sub

Function DeleteSameValue1(ar())

Dim dic    '// �d�����������l���i�[����Dictionary
Set dic = CreateObject("Scripting.Dictionary")
    
    Dim i                       '// ���[�v�J�E���^�P
    Dim ii                      '// ���[�v�J�E���^�Q
    Dim iLen                    '// �z��v�f��
    Dim arEdit()                '// �ҏW��̔z��
    
    ReDim arEdit(0)
    iLen = UBound(ar)
    
    '// �z�񃋁[�v
    For i = 0 To iLen
        '// �z��ɖ��o�^�̒l�̏ꍇ
        If (dic.exists(ar(i)) = False) Then
            '// Dictionary�ɒǉ�
            Call dic.Add(ar(i), ar(i))
            
            '// �d�����Ȃ��l�݂̂�ҏW��z��Ɋi�[����
            arEdit(UBound(arEdit)) = ar(i)
            ReDim Preserve arEdit(UBound(arEdit) + 1)
        End If
    Next
    
    '// �z��Ɋi�[�ς݂̏ꍇ
    If (IsEmpty(arEdit(0)) = False) Then
        '// �]���ȗ̈���폜
        ReDim Preserve arEdit(UBound(arEdit) - 1)
    End If
    
    '// �����ɕҏW��z���ݒ�
    ar = arEdit
End Function

Function ex1(str0, str2, str4, str5, str6, str7, str10, str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
        '�e���v�����R�s�[
        wb_t.Worksheets("temp_1").Copy Before:=wb.Worksheets("Sheet1")
        
    With wb
        .ActiveSheet.Name = str12
    End With
    
    cc = cc + 1
    With wb.Sheets(str12)
        .Cells(1, 2).Value = "�@" & str2                    'm_������X�֔ԍ�
        .Cells(2, 2).Value = "�@" & str4 & str5 & str6      'm_�Z��
        .Cells(3, 2).Value = "�@" & str7                    'm_������
        .Cells(5, 2).Value = "�@" & str0 & "�@�l"             'm_�����於�i���O�j
        .Cells(7, 2).Value = "�@���q�l�ԍ��F" & str57       's_�ڋq�ԍ�
        .Cells(8, 2).Value = "�@�`�[�ԍ��F" & str12         's_�`�[�ԍ�
        .Cells(12, 2).Value = "�d�C�������̂��m�点�@�@" & Date_ss(str14)               's_�d�C�����̂��m�点
        .Cells(15, 4).Value = " " & nk1(str39) & "�~"            's_�����(�ō�)
        .Cells(15, 5).Value = "�i��������ő����z�@�@�@�@�@" & nk1(str40) & "�~�j"      's_����œ������z
        .Cells(18, 4).Value = " " & str57                   's_�ڋq�ԍ�
        .Cells(19, 4).Value = " " & str10                   's_����
        .Cells(20, 4).Value = " " & str19                   's_���p�ꏊ
        .Cells(22, 4).Value = str20 & " �` " & str21        's_�g�p�J�n�� s_�g�p�I����
        .Cells(23, 4).Value = Date_k(str22)                 's_���j��
        .Cells(24, 4).Value = str26                         's_�_����
        .Cells(25, 4).Value = str13                         's_�����n�_����ԍ�
        .Cells(22, 7).Value = str23                         's_����
        .Cells(23, 7).Value = Date_k(str28)                 's_���񌟐j��
        .Cells(24, 7).Value = str27                         's_�_��e��
        .Cells(25, 7).Value = "�����U��"                    '�x�������@
        .Cells(26, 7).Value = Date_sa(str22)                '��������
        .Cells(29, 2).Value = str15                         's_����
        .Cells(29, 5).Value = str16                         's_�P��
        .Cells(29, 7).Value = int_0(nk(str17))              's_�_��d��/�d�͗�)
        .Cells(29, 9).Value = nk(str18) & "�~"              's_����z(�ō�)
        .Cells(41, 9).Value = "����            " & nk(touge2) & "�~"                    '�R������P��_����
        .Cells(42, 9).Value = "����            " & nk(yokuge2) & "�~"                     '�R������P��_����
        .Cells(43, 9).Value = "�����͓����Ɣ��" & pm(Val(yokuge2) - Val(touge2)) & "�~"   '�O���ƍ����̔�r�l
    End With
   
End Function

Function ex2(str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
            
    With wb.Sheets(str12)
        .Cells(28 + C, 2).Value = str15                     's_����
        .Cells(28 + C, 5).Value = nk(str16)                     's_�P��
        .Cells(28 + C, 7).Value = str17                     's_�_��d��/�d�͗�
        .Cells(28 + C, 9).Value = str18 & "�~"              's_����z(�ō�)
    End With
    
End Function

Sub excell_create()
    Dim db As Database
    Dim rs As Recordset

    touge2 = Form_TOP.tougetu
    yokuge2 = Form_TOP.yokugetu
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Q_sales_member")         '�捞�݌��N�G��
    Dim rs_CT As Long
    Dim C As Long
    excel_file = "C:\Users\OCN060605102\Desktop\��ƃt�H���_\���ʌ�t\excel\���ʌ�t���_����e��.xls"                            '�G�N�G���t�@�C��
    temp_excel_file = "C:\Users\OCN060605102\Desktop\��ƃt�H���_\���ʌ�t\excel\temp\���ʌ�t���_����e��_temp.xls"    '�G�N�Z���e���v���[�g
    Set exAPP = CreateObject("Excel.Application")       '�G�N�Z���Z�b�g
    'exAPP.Visible = True                               '��\��
    exAPP.DisplayAlerts = False                         '�x������
    
    Set wb = exAPP.Workbooks.Add                            '�G�N�Z���t�@�C���쐬
    wb.SaveAs (excel_file)                                  '�G�N�Z���t�@�C���w��
    Set wb_t = exAPP.Workbooks.Open(temp_excel_file)        '�G�N�Z���e���v���[�g
    
    
    'ExAPP.Workbooks.Add (excel_file)                      '�G�N�Z���t�@�C���쐬
    'Set wb = ExAPP.Workbooks.Open(excel_file)
    'wb.Save
    'Set wb = ExAPP.wb.Open(excel_file)
    'Set wb_t = ExAPP.Workbooks.Open(temp_excel_file)        '�G�N�Z���e���v���[�g
    
    cc = 0
    rs.MoveFirst
    rs.MoveLast
    rs.MoveFirst
    
    rs_CT = rs.RecordCount
    Myarray = rs.GetRows(rs_CT)     '�Q�����z��Z�b�g
    


    
    
    '���[�v�Ŏ��̃��R�[�h��������ŃZ����������
    For i = 0 To (rs_CT - 1)

        str0 = Myarray(0, i) 'm_�����於
        str1 = Myarray(1, i) 'm_�����於(�J�i)
        str2 = Myarray(2, i) 'm_������X�֔ԍ�
        str3 = Myarray(3, i) 'm_������s���{��
        str4 = Myarray(4, i) 'm_������s�撬��
        str5 = Myarray(5, i) 'm_������Z��
        str6 = Myarray(6, i) 'm_������Ԓn
        str7 = Myarray(7, i) 'm_�����挚����
        str8 = Myarray(8, i) 's_���ID
        str9 = Myarray(9, i) 's_������ID
        str10 = Myarray(10, i) 's_����
        str11 = Myarray(11, i) 's_����敪
        str12 = Myarray(12, i) 's_�`�[�ԍ��i�d���m�F���ځj
        str13 = Myarray(13, i) 's_�����n�_����ԍ�
        str14 = Myarray(14, i) 's_�����N��
        str15 = Myarray(15, i) 's_����
        str16 = Myarray(16, i) 's_�P��
        str17 = Myarray(17, i) 's_�_��d��/�d�͗�
        str18 = Myarray(18, i) 's_����z(�ō�)
        str19 = Myarray(19, i) 's_���p�ꏊ
        str20 = Myarray(20, i) 's_�g�p�J�n��
        str21 = Myarray(21, i) 's_�g�p�I����
        str22 = Myarray(22, i) 's_���j��
        str23 = Myarray(23, i) 's_����
        str24 = Myarray(24, i) 's_�g�p��
        str25 = Myarray(25, i) 's_�v����ID
        str26 = Myarray(26, i) 's_�_����
        str27 = Myarray(27, i) 's_�_��e��
        str28 = Myarray(28, i) 's_���񌟐j��
        str29 = Myarray(29, i) 's_�m��g�p�ʎ捞��
        str30 = Myarray(30, i) 's_�����Z���
        str31 = Myarray(31, i) 's_�����ԍ�
        str32 = Myarray(32, i) 's_������
        str33 = Myarray(33, i) 's_�����G���[
        str34 = Myarray(34, i) 's_�x�����@
        str35 = Myarray(35, i) 's_����(�U��)�w���
        str36 = Myarray(36, i) 's_�R���r�j�x������
        str37 = Myarray(37, i) 's_�����˗��A�g��
        str38 = Myarray(38, i) 's_���J��
        str39 = Myarray(39, i) 's_�����(�ō�)
        str40 = Myarray(40, i) 's_����œ������z
        str41 = Myarray(41, i) 's_�܂Ƃߐ����z
        str42 = Myarray(42, i) 's_�c��
        str43 = Myarray(43, i) 's_������
        str44 = Myarray(44, i) 's_�����z
        str45 = Myarray(45, i) 's_�������@
        str46 = Myarray(46, i) 's_�ŏI������
        str47 = Myarray(47, i) 's_����
        str48 = Myarray(48, i) 's_���X�e�[�^�X
        str49 = Myarray(49, i) 's_�x������
        str50 = Myarray(50, i) 's_���ؓ���
        str51 = Myarray(51, i) 's_�v���
        str52 = Myarray(52, i) 's_�o�^��
        str53 = Myarray(53, i) 's_�ŏI�X�V��
        str54 = Myarray(54, i) 's_�������X��
        str55 = Myarray(55, i) 's_���[�����M
        str56 = Myarray(56, i) 's_�㗝�X
        str57 = Myarray(57, i) 's_�O���L�[

        Dim tc As Long
                
        If i = 0 Then '�ŏ�
        
                '�@�����R�[�h�����Ȃ�J�E���g�{�P
                If str12 = Myarray(12, i + 1) Then
                    C = C + 1
                    Debug.Print "�@"
                    Call ex1(str0, str2, str4, str5, str6, str7, str10, str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    
                    Debug.Print str12
                    Debug.Print str15 & C
                    
                '�A���̃��R�[�h���Ⴄ
                    Else
                    C = 1
                    Debug.Print "�A"
                    Call ex1(str0, str2, str4, str5, str6, str7, str10, str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str12
                    Debug.Print str15 & C
                    C = 0
                    tc = tc + 1
                End If
                
        ElseIf i = rs_CT - 1 Then '�Ō�
        
                '�B�O���R�[�h�����Ȃ�J�E���g�{�P
                If str12 = Myarray(12, i - 1) Then
                    C = C + 1
                    Debug.Print "�B"
                    Call ex2(str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str15 & C
                    
                '�C�O���R�[�h�ƈႤ���Ō�P�J�E���g
                Else
                    C = 1
                    Debug.Print "�C"
                    Call ex1(str0, str2, str4, str5, str6, str7, str10, str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str12
                    Debug.Print str15 & C
                    tc = tc + 1
                End If
                
        Else
                
                '�D�����R�[�h������& C = 0�Ȃ�J�E���g�{�P
                If str12 = Myarray(12, i + 1) And C = 0 Then
                    C = C + 1
                    Debug.Print "�D"
                    Call ex1(str0, str2, str4, str5, str6, str7, str10, str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str12
                    Debug.Print str15 & C
                    
                '�E�����R�[�h�����Ȃ�J�E���g�{�P
                ElseIf str12 = Myarray(12, i + 1) Then
                    C = C + 1
                    Debug.Print "�E"
                    Call ex2(str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str15 & C
                End If
                
                '�F�����R�[�h�Ⴄ���O���R�[�h�Ɠ���
                If str12 <> Myarray(12, i + 1) And str12 = Myarray(1, i - 1) Then
                    C = C + 1
                    Debug.Print "�F"
                    Call ex2(str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str15 & C
                    C = 0
                    tc = tc + 1
                
                '�G�����R�[�h��Ƃ��Ȃ�J�E���g�O
                ElseIf str12 <> Myarray(12, i + 1) Then
                    
                    If str12 = Myarray(12, 0) Then '�ŏ��̎捞�ݎ҂��m�F
                    C = C + 1
                    Debug.Print "�H"
                    Call ex2(str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str15 & C
                    Else
                    C = C + 1
                    Debug.Print "�I" '���̐l��
                    Call ex2(str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)

                    'Call ex1(str0, str2, str4, str5, str6, str7, str10, str12, str13, _
            'str14, str15, str16, str17, str18, str19, str20, str21, _
            'str22, str23, str26, str27, str28, str39, str40, str57, c)
                    Debug.Print str12
                    Debug.Print str15 & C
                    End If
                    

                    C = 0
                    tc = tc + 1
                End If
            
        End If
       
    Next i
        
    '�G�N�Z���I������
    With wb
        .Worksheets("Sheet1").Delete
        .SaveAs (excel_file)
    End With
    
    'exAPP.DisplayAlerts = True

    
    Set wb = Nothing            '���[�N�V�[�g���
    Set wb_t = Nothing          '���[�N�V�[�g���


    exAPP.Quit
    Set exAPP = Nothing         '�G�N�Z�����

'Public exAPP As Object     '�I�u�W�F�N�g�G�N�Z���錾
'Public wb As Object        '�I�u�W�F�N�g���[�N�u�b�N�錾
'Public wb_t As Object      '�I�u�W�F�N�g���[�N�u�b�N�錾


MsgBox "�`�[�ԍ��̐��F" & cc
End Sub

'************************************************************************************************************
'**2�����z����t�B���^�[����i����̂P��ɑ΂��ĂЂƂ̒l�Ɉ�v������̂𒊏o
'**�������Fdata => �t�B���^�[������2�����z�� �o���A���g�^
'**�������Fcol_num => �t�B���^�[�����̔ԍ��i������1���琔����j�����^
'**��O�����Fkey_array => ���o�������L�[���[�h
'**��data�Ƀw�b�_�[���܂܂��O��A�w�b�_�[���܂܂ꂽ�z���Ԃ�
'**��Option Base 1 ��K���g�p������ŗ��p���邱��
'*************************************************************************************************************
Function �񎟌��z��t�B���^�[�֐�(ByVal Data As Variant, ByVal col_num As Integer, ByVal key As String) As Variant
    Dim cnt As Long '�t�B���^�[��̔z��̍s��
    Dim n_col As Long '�z��̗�
    Dim dic As Object
    Dim r_array As Variant '�����ɍ�����̊e����ꎞ�I�Ɋi�[���邽�߂̔z��
    Dim data_fil As Variant '�t�B���^�[���2�����z��
    '���̔z��(data)�̗񐔁i=�t�B���^�[��̗�
    n_col = UBound(Data, 2)
    '��L�̗񐔁i�ϐ��j��p����r_array���Ē�`
    ReDim r_array(n_col) As String
    '�����^�z����`
    Set dic = CreateObject("Scripting.Dictionary")
    '�t�B���^�[��̔z��̍s�����J�E���g���A�Y�������s��z��Ɋi�[����
    cnt = 1 '�w�b�_�[���̗񐔂����炩���ߍl��
    For i = 2 To UBound(Data)
        If CStr(Data(i, col_num)) = key Then
            '���o�������s�̗��z��Ɋi�[����
            For J = 1 To n_col
                r_array(J) = Data(i, J)
            Next J
            cnt = cnt + 1
            dic.Add cnt, r_array
        End If
    Next i
    '�w�b�_�[�����������^�z��Ɋi�[
    For J = 1 To n_col
        r_array(J) = Data(1, J)
    Next J
    dic.Add 1, r_array
    '�t�B���^�[��̔z��̍Ē�`
    ReDim data_fil(cnt, n_col) As String
    '�t�B���^�[��̔z��Ƀw�b�_�[���i�[
    For J = 1 To n_col
        data_fil(1, J) = dic.Item(1)(J)
    Next J
    '�w�b�_�[�ȊO�̒l���t�B���^�[��̔z��Ɋi�[
    For i = 2 To cnt
        For J = 1 To n_col
            data_fil(i, J) = dic.Item(i)(J)
        Next J
    Next i
    '�֐��ɑ��
    �񎟌��z��t�B���^�[�֐� = data_fil
End Function


