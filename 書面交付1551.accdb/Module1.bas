Attribute VB_Name = "Module1"
Option Compare Database
Sub Importtest()
Dim d, s
'test
's = "Path = ""C:\Users\OCN060605102\Desktop\��ƃt�H���_\���ʌ�t\csv\bqrcsv\���ʌ�tALL20221026.csv"""
'd = "Path = ""C:\Users\OCN060605102\Desktop\��ƃt�H���_\���ʌ�t\csv\bqrcsv\���ʌ�tALL.csv"""

Dim Prj As CodeProject
Dim Obj As ImportExportSpecification
Dim ImpName As String, strXML As String

'�p�X�����m�F����C���|�[�g����̖��O���w��
ImpName = "�C���|�[�g-���ʌ�tALL"

'XML�v���p�e�B�̒l���C�~�f�B�G�C�g �E�B���h�E�ɕ\��
Set Prj = CurrentProject
Set Obj = Prj.ImportExportSpecifications(ImpName)
strXML = Obj.XML
Debug.Print strXML

'path�擾���ύX-----------------------
Debug.Print InStr(strXML, "Path") + 8
'Debug.Print InStr(strXML, "xmlns")
path_c = (InStr(strXML, "xmlns") - 2) - (InStr(strXML, "Path") + 8)
path_cc = Mid(strXML, InStr(strXML, "Path") + 8, path_c)

'�ύX�O
s = "Path = """ & path_cc & """"
'�ύX��
d = "Path = ""C:\Users\OCN060605102\Desktop\��ƃt�H���_\���ʌ�t\csv\bqrcsv\���ʌ�tALL20221025.csv"""
strXML = Replace(strXML, s, d)
Debug.Print strXML
Obj.XML = strXML
'path�擾���ύX-----------------------

'�捞
DoCmd.RunSavedImportExport "�C���|�[�g-���ʌ�tALL"

'�e�[�u�����ύX(���ʌ�tALL�����ʌ�tALLYYYYMMDD)
DoCmd.Rename "���ʌ�tALL20221025", acTable, "���ʌ�tALL"


End Sub
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

Sub table_create()
    'table���݊m�F�A����΃N���A
    'table�쐬����(�I�[�g�i���o�[�̃C���f�b�N�X����)
    
    Dim TABLE_NAME As String
    TABLE_NAME = "SYOMEN"
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    Set db = CurrentDb
    
    For Each tdf In db.TableDefs
        'Debug.Print tdf.Name
            If tdf.Name = TABLE_NAME Then
                DoCmd.DeleteObject acTable, TABLE_NAME
            End If
    Next
    
    
    Set tb = db.CreateTableDef(TABLE_NAME) '�e�[�u���쐬
    tb.Fields.Append tb.CreateField("ID", dbLong) 'ID
    'PrimaryKey��ݒ�
    Set idx = tb.CreateIndex("PrimaryKey")
    idx.Primary = True
    '�C���f�b�N�X���\������uID�v�t�B�[���h���쐬����
    idx.Fields.Append idx.CreateField("ID")
    tb.Indexes.Append idx '�쐬�����C���f�b�N�X���R���N�V�����ɒǉ�����
   
    tb.Fields.Append tb.CreateField("C_id", dbText, 9) '�ڋq�ԍ�
    tb.Fields.Append tb.CreateField("Name", dbText, 50) '����
    tb.Fields.Append tb.CreateField("Tanmatu", dbText, 50) '�[���ԍ�
    tb.Fields.Append tb.CreateField("Syouhin", dbText, 50) '���i����
    tb.Fields.Append tb.CreateField("Hojyo", dbText, 50) '�⏕�Ȗږ���
    tb.Fields.Append tb.CreateField("Nebiki", dbText, 50) '�l�����z
    tb.Fields.Append tb.CreateField("Seikyuu", dbText, 50) '�������z
    tb.Fields.Append tb.CreateField("Teijyou", dbText, 50) '���l���z
    tb.Fields.Append tb.CreateField("Keiyakubi_u", dbText, 50) '�_��(��t)
    tb.Fields.Append tb.CreateField("Keiyakubi_m", dbText, 50) '�_��i�\���j
    tb.Fields.Append tb.CreateField("Kaisibi", dbText, 50) '�T�[�r�X�J�n��
    tb.Fields.Append tb.CreateField("Shinki", dbText, 50) '�V�Kck
    tb.Fields.Append tb.CreateField("setwari", dbText, 50) '�Z�b�g��
    tb.Fields.Append tb.CreateField("Sv", dbText, 50) '�T�[�r�X
    tb.Fields.Append tb.CreateField("CP_END", dbText, 50) 'CP�I��
     tb.Fields.Append tb.CreateField("F", dbBoolean, 10) '���ʗL��
    
    'ID�t�B�[���h���I�[�g�i���o�[�ɂ���
    tb.Fields("ID").Attributes = dbAutoIncrField
    '�e�[�u����ǉ�
    db.TableDefs.Append tb
    db.Close
    
    
End Sub
Sub ImportALLDAY()
    DoCmd.RunSavedImportExport "�C���|�[�g-���ʌ�tALL"
    'DoCmd.RunSavedImportExport "�C���|�[�g-���ʌ�tDAY"
End Sub

Private Sub TextImport()

'�ϐ��錾
Dim FilePath As String '�t�@�C���p�X
Dim XDAY As String '�捞��

XDAY = "20221101"
'CSV�t�@�C����


'csv�C���|�[�g���̃t�@�C���p�X
FilePath = "C:\Users\OCN060605102\Desktop\��ƃt�H���_\���ʌ�t\csv\bqrcsv\���ʌ�tALL" & XDAY & ".csv"

'csv�t�@�C���̃C���|�[�g
DoCmd.TransferText acImportDelim, "���ʌ�tDAY �C���|�[�g��`", "���ʌ�tALL" & XDAY, FilePath

'csv�t�@�C�����C���|�[�g�����|��ʒm����B
MsgBox "csv�t�@�C�����C���|�[�g���܂����B"

End Sub

Sub taisyo1()
'Call ImportALLDAY
Call table_create
'�Ώ����̎捞
'�捞��Ώیڋq�ԍ��̑Ώ����Ń��[�v�捞����IF��stb�����^��������A���΂̊�{�f�W�^���������ꍇ�΂ƂȂ�s��}��
'���ʌ�t_�d���J�b�g�����[�v
    Dim StringKey As String
    Dim NiraiCk As String
    Dim db As Database
    Dim rs As Recordset
    Dim rs_count As Long
    Dim rs_all_count As Long
    Dim sinki As String
    
    Set db = CurrentDb()
    '========================================
    Dim kaisibi As String
    Dim kaisibi1 As String
    Dim kaisibi2 As String '����
    Dim kaisibi3 As String '�O��
    Dim kaisibi4 As String '�O�X��
    Dim kaisibi5 As String
    Dim F As Boolean
    Dim KS As String
    kaisibi = "2022/10/24"
    kaisibi1 = Replace(kaisibi, "/", "")
    kaisibi2 = "���ʌ�tALL" & kaisibi1
    kaisibi3 = "���ʌ�tALL" & zenjitu1(kaisibi)
    kaisibi4 = "���ʌ�tALL" & zenjitu2(kaisibi) '�O�X��
    kaisibi5 = zenjitu3(kaisibi) '�O��YYYY/MM/DD
    'Debug.Print kaisibi5
    '========================================
    '���ʌ�t�Ώۂ̓��i�J�n���j
    SQL1 = "SELECT " & kaisibi2 & ".* FROM " & kaisibi2 & " WHERE " & kaisibi2 & ".�T�[�r�X�J�n�� = '" & kaisibi & "' "
    'Debug.Print SQL1
    'Set rs = db.OpenRecordset("���ʌ�tDAY")
    Set rs = CurrentDb.OpenRecordset(SQL1)
    
    '���ʌ�t�Ώۂ̓��i�ڋq�ԍ��j
    SQL = "SELECT " & kaisibi2 & ".* FROM " & kaisibi2 & " WHERE " & kaisibi2 & ".�ڋqCD = '" & kokyaku & "' "
    
    Set RS_ALL = CurrentDb.OpenRecordset(SQL)
    

    
    '�������ݗp
    Set rs_syomen = db.OpenRecordset("SYOMEN", dbOpenDynaset)
    
    rs.MoveLast
    rs.MoveFirst
    rs_count = rs.RecordCount
    
    Myarray = rs.GetRows(rs_count) '���R�[�h��2�����z��
    
    Dim Myarray2
    '�s�����ւ�(�Y�������O���P�ɕύX�j
    ReDim Myarray2(1 To UBound(Myarray, 2) + 1, 1 To UBound(Myarray, 1) + 1)
    For i = 1 To UBound(Myarray, 1) + 1
        For J = 1 To UBound(Myarray, 2) + 1
            Myarray2(J, i) = Myarray(i - 1, J - 1)
        Next
    Next

'---------------------------------


'---------------------------------


Dim Myarray6
'Dictionary�̃I�u�W�F�N�g���쐬
Dim a
Set a = CreateObject("Scripting.Dictionary")
    
    
    '�ϐ��Ƀt�B���^�������ڋq�ԍ��ŕ���
    'Data1 = FilterArray2D(Myarray2, StringKey, 2)

'�u�d���v��ۑ�����z��
Dim Data
ReDim Data(UBound(Myarray2, 1), 1 To 1)
'Dim StringKey As String
'ReDim Preserve Myarray(LBound(Myarray, 1), UBound(Myarray) + 1)

For i = 1 To UBound(Myarray2, 1) '�ڋq�ԍ��ꗗ���擾
    If a.exists(Myarray2(i, 2)) = False Then
    a.Add Myarray2(i, 2), 2 '�o�^����i1�͂Ȃ�ł������j
    Else '�g�p����Ă���ꍇ
    End If
Next
Keys = a.Keys
items = a.items

Dim myRegEx As Object '���K���ϐ�
Set myRegEx = CreateObject("VBScript.RegExp") 'RegEx�I�u�W�F�N�g������
myRegEx.pattern = "^(3Z0|3Z2)\w{2}"
Dim regCheck As Boolean
Dim iii As Long
For ii = 0 To a.Count - 1 '�d���J�b�g�̌ڋq�ԍ��̃��[�v�J�n
 
        'Debug.Print Keys(i)
        StringKey = Keys(ii)
        
        '�ΏۂƂȂ�Ώ�����ϐ��֊i�[
        
        
        Data1 = FilterArray2D(Myarray2, StringKey, 2) '�Ώۂ̌ڋq�ԍ��̂݃��[�v����,�z��,�L�[,�z��̃J��������
        
        '����---------------------------
        
        
    '���ʌ�tDAY(�Ώیڋq�̑S�Ă𒊏o)
    SQL1 = "SELECT " & kaisibi2 & ".* FROM " & kaisibi2 & " WHERE " & kaisibi2 & ".�ڋqCD = '" & StringKey & "'"
            Debug.Print SQL1
    'Set rs = db.OpenRecordset("���ʌ�tDAY")
    'SQL1 = "SELECT ���ʌ�tALL.* FROM ���ʌ�tALL WHERE ���ʌ�tALL.�T�[�r�X�J�n�� = '" & kaisibi & "' "
    Set rs_a = CurrentDb.OpenRecordset(SQL1)
    rs_a.MoveLast
    rs_a.MoveFirst
    rs_count_a = rs.RecordCount
    
    Myarray33 = rs_a.GetRows(rs_count_a) '���R�[�h��2�����z��
    
    Dim Myarray3
    '�s�����ւ�(�Y�������O���P�ɕύX�j
    ReDim Myarray3(1 To UBound(Myarray33, 2) + 1, 1 To UBound(Myarray33, 1) + 1)
    For i = 1 To UBound(Myarray33, 1) + 1
        For J = 1 To UBound(Myarray33, 2) + 1
            Myarray3(J, i) = Myarray33(i - 1, J - 1)
        Next
    Next
    
    Data2 = FilterArray2D(Myarray3, StringKey, 2) '�Ώۂ̌ڋq�ԍ��̑S��
        
    '--------------------------
        
        
    '���ʌ�t�Ώۓ��̑O�X���@CP�I����r�p�i�ڋq�ԍ��j
    'SQL4 = "SELECT " & kaisibi4 & ".* FROM " & kaisibi4 & " WHERE " & kaisibi4 & ".�ڋqCD = '" & StringKey & "' & " And kaisibi4 & ".�T�[�r�X�I���� = '" & kaisibi5 & "' "
    SQL4 = "SELECT " & kaisibi4 & ".* FROM " & kaisibi4 & " " & _
            "WHERE " & kaisibi4 & ".�ڋqCD = '" & StringKey & "' " & _
            "And " & kaisibi4 & ".�T�[�r�X�I���� = '" & kaisibi5 & "' "
    'Debug.Print SQL4
    Set rs_CPCK = CurrentDb.OpenRecordset(SQL4)
        
'        rs_CPCK.MoveLast
'        rs_CPCK.MoveFirst
'        rs_CPCK_count = rs_CPCK.RecordCount
            If rs_CPCK.EOF Then
            Debug.Print "�ʏ�H��"
            Else
'CP�`�F�b�N
    SQL_CP = "UPDATE SYOMEN " & _
            "SET SYOMEN.CP_END = '" & kaisibi5 & "' " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "'"
            cpck = 1
            Debug.Print SQL_CP
            'db.Execute SQL_CP(���s�͍Ō��)
            
'                Debug.Print "�O��CP�I��-" & "�ڋq�ԍ��F" & StringKey
'                            rs_syomen.AddNew
'                            rs_syomen!C_id = StringKey
'                            rs_syomen!Name = "-CP�I��-"
'                            rs_syomen!CP_END = kaisibi5
'                            rs_syomen.Update
                            
            End If
            '���ʌ�t�Ώۓ��@�T�[�r�X�J�n�����_���t�����r�B"��"�Ō_�񏑂Ȃ��B�R�[�X�ύX�B�Ȃ̂ŏ��ʌ�t����(�v�`�F�b�N)
'    SQL5 = "SELECT " & kaisibi2 & ".* FROM " & kaisibi2 & " " & _
'            "WHERE " & kaisibi2 & ".�ڋqCD = '" & StringKey & "' " & _
'            "And " & kaisibi2 & ".�_���t�� = '" & kaisibi & "' "

    SQL5 = "SELECT DISTINCT " & kaisibi2 & ".���Ƌ敪_C FROM " & kaisibi2 & " " & _
            "WHERE " & kaisibi2 & ".�ڋqCD = '" & StringKey & "' " & _
            "And " & kaisibi2 & ".���Ƌ敪_C <> '�đ�' " & _
            "And " & kaisibi2 & ".���Ƌ敪_C <> 'OTHERS' " & _
            "And " & kaisibi2 & ".���Ƌ敪_C NOT IN " & _
            "(SELECT " & kaisibi3 & ".���Ƌ敪_C FROM " & kaisibi3 & " " & _
            "WHERE " & kaisibi3 & ".�ڋqCD = '" & StringKey & "')"
            
'UPDATE

    Debug.Print SQL5
'    Dim kaisibi3 As String '�O�X�� �T�u�N�G������(�O���̌_��Ɣ�r���V�K�̂ݒ��o)
        Set rs_kouji_CK = CurrentDb.OpenRecordset(SQL5)
        
            
            If rs_kouji_CK.EOF And Not cpck = 1 Then
            
            Debug.Print "�V�K����"
'                rs_syomen.AddNew
'                rs_syomen!C_id = StringKey
'                rs_syomen!Name = "-����-"
'                rs_syomen.Update
                F = True
                sinki = "�p��"
            Else
            
            Debug.Print "�V�K"
'                rs_syomen.AddNew
'                rs_syomen!C_id = StringKey
'                rs_syomen!Name = "-NO����-"
'                rs_syomen.Update


                F = False
                sinki = "�V�K"
            End If
            
        For iii = 1 To UBound(Data1)
            '�O�X���ɃT�[�r�X�I�������Z�b�g����Ă��邩�B����Ă���Ȃ�b�o�I���Ȃ̂ŃX���[
'            If rs_CPCK.EOF Then
'            Debug.Print "����"
'            End If
            '�Ώ������C�R�[�����m�F
            If Data1(iii, 51) = kaisibi Then
                        If Data1(iii, 31) <> "�n��g�E�a�r" And Data1(iii, 31) <> "BS-NHK" And Data1(iii, 31) <> "BS-����" Then '�n��g�E�a�r�͕K�v�Ȃ������Ȃ̂ŃJ�b�g�i���ɂ�����΂����ɒǉ��j
                            '���ڊm��F
                            Debug.Print Data1(iii, 2) & " " & Data1(iii, 3) & " " & Data1(iii, 31) & " " & Data1(iii, 51)
                            NiraiCk = NiraiCk + Data1(iii, 31)
                            'table�֒ǉ�
                            rs_syomen.AddNew
                            rs_syomen!C_id = Data1(iii, 2)
                            rs_syomen!Name = Data1(iii, 3)
                            rs_syomen!Tanmatu = Data1(iii, 28)
                            rs_syomen!Syouhin = Data1(iii, 31)
                            rs_syomen!Keiyakubi_u = Data1(iii, 49) '�_���t��
                            rs_syomen!Keiyakubi_m = Data1(iii, 50) '�_��\����
                            rs_syomen!kaisibi = Data1(iii, 51)
                            rs_syomen!Hojyo = Data1(iii, 45)
                            rs_syomen!Nebiki = Data1(iii, 46)
                            rs_syomen!Seikyuu = Data1(iii, 47)
                            rs_syomen!Teijyou = Data1(iii, 48)
                            rs_syomen!Sv = Data1(iii, 91)
                            rs_syomen.Update
                        End If
                        regCheck = myRegEx.test(Data1(iii, 29))
                        
                If regCheck = True Then '�r�s�a�����^�����m�F
        '            MsgBox "STB�����^������"
                    z = Data1(iii, 28) '�[���ԍ�
                    y = Data1(iii, 51) '�T�[�r�X�J�n��
                    '�΂̊�{�f�W�^���`�F�b�N
                    For iiii = 1 To UBound(Data2)
                        'If iiii = iii Then '���g�̓X�L�b�v
                        '    iiii = iiii + 1
                        'Else
                            If Data2(iiii, 28) = z And (Data2(iiii, 45) = "��{�f�W�^��" Or Data2(iiii, 45) = "�ԑg�K�C�h") Then '�����[���ԍ��Ŋ�{�f�W�^���������͔ԑg�K�C�h�ł���
        '                        MsgBox "�����[���ԍ��Ŋ�{�f�W�^���ł���"
                                    If Data2(iiii, 51) <> y Then '�T�[�r�X�J�n�����A�΂̊�{�f�W�^���ƈႤ
                                        Debug.Print "STOP"
                                        Debug.Print Data2(iiii, 2) & " " & Data2(iiii, 3) & " " & Data2(iiii, 31) & " " & Data2(iiii, 51)
                                        
                                        '�e�[�u���ɒǉ�
                                        rs_syomen.AddNew
                                        rs_syomen!C_id = Data2(iiii, 2)
                                        rs_syomen!Name = Data2(iiii, 3)
                                        rs_syomen!Tanmatu = Data2(iiii, 28)
                                        rs_syomen!Syouhin = Data2(iiii, 31) & "��"
                                        rs_syomen!Keiyakubi_u = Data2(iiii, 49) '�_���t��
                                        rs_syomen!Keiyakubi_u = Data2(iiii, 50) '�_��\����
                                        rs_syomen!kaisibi = Data2(iiii, 51)
                                        rs_syomen!Hojyo = Data2(iiii, 45)
                                        rs_syomen!Nebiki = Data2(iiii, 46)
                                        rs_syomen!Seikyuu = Data2(iiii, 47)
                                        rs_syomen!Teijyou = Data2(iiii, 48)
                                        rs_syomen!Sv = Data2(iiii, 91)
                                        rs_syomen.Update
                                        NiraiCk = NiraiCk + Data2(iiii, 31) & "��"
                                    End If
                                    
                                    
                                
                            End If
                        'End If
                    Next
                End If
            Else
            End If
        Next


    


'�ɂ炢�p�b�Nor�ɂ炢�v���X�Ń����^�����Ȃ������ꍇ�R�[�X��ǉ�����
If (NiraiCk Like "*�ɂ炢�p�b�N*" Or NiraiCk Like "*�ɂ炢�v���X*") And Not NiraiCk Like "*�����^��*" Then
Debug.Print "��������"
'ALL����Ώۂ̌ڋq�ԍ��̂ݒ��o�����̒��̃T�[�r�X�̂݉�����
'kokyaku = "100711901"
SQL = "SELECT * FROM " & kaisibi2 & " WHERE �ڋqCD = '" & StringKey & "' "
Set RS_ALL = CurrentDb.OpenRecordset(SQL)

'count ck
RS_ALL.MoveLast
RS_ALL.MoveFirst
rs_all_count = RS_ALL.RecordCount

For i = 1 To rs_all_count
    
    If RS_ALL.���i���� = "�|�s�����[" Or RS_ALL.���i���� = "�v���C��" Or RS_ALL.���i���� = "�|�s�����[2��ڈȍ~" Or RS_ALL.���i���� = "�v���C���Q��ڈȍ~" Or RS_ALL.���i_C = "3Z026" Then  '�T�[�r�X���Ȃ�ǉ�����
                                                '�e�[�u���ɒǉ�
        rs_syomen.AddNew
        rs_syomen!C_id = RS_ALL!�ڋqCD '�ڋq�ԍ�
        rs_syomen!Name = RS_ALL!�ڋq����  '�T�[�r�X��
        rs_syomen!Tanmatu = RS_ALL!�[���ԍ� '�[���ԍ�
        rs_syomen!Syouhin = RS_ALL!���i���� & "�@��" '���i����
        rs_syomen!Keiyakubi_u = RS_ALL!�_���t�� '������t��
        rs_syomen!Keiyakubi_m = RS_ALL!�_��\���� '�����\����
        rs_syomen!kaisibi = RS_ALL!�T�[�r�X�J�n�� '�T�[�r�X�J�n��
        rs_syomen!Hojyo = RS_ALL!�⏕�Ȗږ��� '�⏕
        rs_syomen!Nebiki = RS_ALL!�l�����z '�l��
        rs_syomen!Seikyuu = RS_ALL!�������z '����
        rs_syomen!Teijyou = RS_ALL!���l���z '���
        rs_syomen!Sv = RS_ALL!���Ƌ敪_C '�T�[�r�X
        rs_syomen.Update
        NiraiCk = NiraiCk + RS_ALL!���i���� & "�@��"
    End If

    RS_ALL.MoveNext
Next

End If



    '����ck
    Dim wari As String
    wari = ""
        'TV
        Select Case True
            Case NiraiCk Like "*STB*" Or NiraiCk Like "*�ɂ炢�p�b�N*" Or NiraiCk Like "*�ɂ炢�v���X*"
            wari = wari + "TV"
        End Select
        'NET
        Select Case True
            Case NiraiCk Like "*�p�[�\�i��*" Or NiraiCk Like "*�G�R�m�~�[*" Or NiraiCk Like "*�v���~�A��*" Or NiraiCk Like "*�v���`�i�P�Q�O*" Or NiraiCk Like "*�q�J���ɂ炢*"
            wari = wari + "NET"
        End Select
        'TEL
        Select Case True
            Case NiraiCk Like "*�P�[�u���v���X*"
            wari = wari + "TEL"
        End Select
    
    'Debug.Print wari
Select Case wari
    Case "TVNET"
'    rs_syomen.AddNew
'    rs_syomen!C_id = StringKey
'    rs_syomen!Tanmatu = "00"
'        rs_syomen!Syouhin = "����"
'    rs_syomen!setwari = "�_�u��"
'    rs_syomen.Update
    
End Select

Select Case wari
    Case "TVTEL"
'    rs_syomen.AddNew
'    rs_syomen!C_id = StringKey
'    rs_syomen!Tanmatu = "00"
'    rs_syomen!Syouhin = "����"
'    rs_syomen!setwari = "�_�u��"
'    rs_syomen.Update
End Select

Select Case wari
    Case "NETTEL"
'    rs_syomen.AddNew
'    rs_syomen!C_id = StringKey
'    rs_syomen!Tanmatu = "00"
'    rs_syomen!Syouhin = "����"
'    rs_syomen!setwari = "�_�u��"
'    rs_syomen.Update
End Select

Select Case wari
    Case "TVNETTEL"
'    rs_syomen.AddNew
'    rs_syomen!C_id = StringKey
'    rs_syomen!Tanmatu = "00"
'    rs_syomen!Syouhin = "����"
'    rs_syomen!setwari = "�g���v��"
'    rs_syomen.Update
End Select

If wari = "" Then
wari = "�Z�b�g������"

End If

NiraiCk = ""
NiraiSinki = ""
NiraiSinki1 = ""
'�V�K�`�F�b�N�J�n

'�H����
SQL2 = "SELECT " & kaisibi2 & ".���Ƌ敪_C FROM " & kaisibi2 & " " & _
        "GROUP BY " & kaisibi2 & ".�ڋqCD, " & kaisibi2 & ".���Ƌ敪_C " & _
        "HAVING " & kaisibi2 & ".���Ƌ敪_C <> 'OTHERS' AND " & kaisibi2 & ".���Ƌ敪_C <> '�đ�' " & _
        "AND " & kaisibi2 & ".�ڋqCD = '" & StringKey & "'"
Set RS_SINKI = CurrentDb.OpenRecordset(SQL2)


'�H���O��
SQL3 = "SELECT " & kaisibi3 & ".���Ƌ敪_C FROM " & kaisibi3 & " GROUP BY " & kaisibi3 & ".�ڋqCD, " & kaisibi3 & ".���Ƌ敪_C HAVING " & kaisibi3 & ".���Ƌ敪_C <>'OTHERS' AND " & kaisibi3 & ".���Ƌ敪_C <>'�đ�' AND " & kaisibi3 & ".�ڋqCD = '" & StringKey & "'ORDER BY " & kaisibi3 & ".���Ƌ敪_C;"
'Debug.Print SQL3
Set RS_SINKI1 = CurrentDb.OpenRecordset(SQL3)

    '--------------------------
    '�V�K�`�F�b�N(�O�����R�[�h)
    If RS_SINKI1.EOF Then
        '���S�Ȃ�V�K�_��
        Debug.Print "���S�V�K"
'            rs_syomen.AddNew
'            rs_syomen!C_id = StringKey
'            rs_syomen!Shinki = "�V�K"
'            rs_syomen.Update
            KS = "����"
        Else
    End If
    '----------------------------






'RECODESET���݃`�F�b�N
    If RS_SINKI.EOF Then
    
    Else
        RS_SINKI.MoveLast
        RS_SINKI.MoveFirst
        rs_all_sinki_count = RS_SINKI.RecordCount
        
        For ss = 1 To rs_all_sinki_count
            'Debug.Print RS_SINKI.���Ƌ敪_C & "�y�����z"
            NiraiSinki = NiraiSinki + RS_SINKI.���Ƌ敪_C
            RS_SINKI.MoveNext
        Next
        
    End If
   
    

'RECODESET���݃`�F�b�N
    If RS_SINKI1.EOF Then
    
    Else
        RS_SINKI1.MoveLast
        RS_SINKI1.MoveFirst
        rs_all_sinki_count1 = RS_SINKI1.RecordCount
        
        For ss = 1 To rs_all_sinki_count1
            'Debug.Print RS_SINKI1.���Ƌ敪_C & "�y�O���z"
            NiraiSinki1 = NiraiSinki1 + RS_SINKI1.���Ƌ敪_C
            RS_SINKI1.MoveNext
        Next
        RS_SINKI1.MoveFirst
    End If

'-------------------�O������͈����Z
If RS_SINKI1.EOF Then
Else
        'RS_SINKI1.MoveFirst
        RS_SINKI.MoveFirst
        For ss = 1 To rs_all_sinki_count1
            NiraiSinki = Replace(NiraiSinki, RS_SINKI1.���Ƌ敪_C, "")
            RS_SINKI1.MoveNext
        Next ss
        
        If NiraiSinki = "" Then
            Debug.Print "�V�K�Ȃ�"
        Else
            Debug.Print "�V�K����:" & NiraiSinki
        End If
'-------------------
End If
    
    '���R�[�h�X�V����
    Select Case NiraiSinki
    Case "NETPTEL"
    
    Case ""
    Case ""
    Case ""
    End Select
    
    Select Case wari
    Case "NETTEL"
    wari = "�_�u��"
    Case "TVTEL"
    wari = "�_�u��"
    Case "TVNET"
    wari = "�_�u��"
    Case "TV"
    wari = "�Z�b�g������"
    Case "NET"
    wari = "�Z�b�g������"
    Case "TEL"
    wari = "�Z�b�g������"
    End Select

    
    SQL6 = "UPDATE SYOMEN " & _
            "SET SYOMEN.F = " & F & " " & _
            ", SYOMEN.setwari = '" & wari & "' " & _
            ", SYOMEN.Shinki = '" & sinki & "' " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "'"
            Debug.Print SQL6
'            db.Execute SQL6
'Set RS_SINKI0 = CurrentDb.OpenRecordset(SQL6)




'�I�v�V����CH��ǉ�
SQL10 = "(SELECT DISTINCT SYOMEN.Tanmatu FROM SYOMEN " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "' )"
SQL10_2 = "(SELECT DISTINCT SYOMEN.Syouhin FROM SYOMEN " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "' )"
SQL10_3 = "('BS-NHK', 'BS-����', 'BS-WOWOW')"
SQL10_1 = "INSERT INTO SYOMEN ( C_id, Name, Tanmatu, Syouhin, Hojyo, Nebiki, Seikyuu, Teijyou, Keiyakubi_u, Keiyakubi_m, Kaisibi, Sv ) " & _
            "SELECT " & kaisibi2 & ".�ڋqCD, " & kaisibi2 & ".�ڋq����, " & _
            "" & kaisibi2 & ".�[���ԍ�, " & kaisibi2 & ".���i����, " & _
            "" & kaisibi2 & ".�⏕�Ȗږ���, " & kaisibi2 & ".�l�����z, " & _
            "" & kaisibi2 & ".�������z, " & kaisibi2 & ".���l���z, " & _
            "" & kaisibi2 & ".�_���t��, " & kaisibi2 & ".�_��\����, " & _
            "" & kaisibi2 & ".�T�[�r�X�J�n��, " & kaisibi2 & ".���Ƌ敪_C " & _
            "FROM " & kaisibi2 & " " & _
            "WHERE (((" & kaisibi2 & ".�ڋqCD)= '" & StringKey & "') AND (" & kaisibi2 & ".�[���ԍ�) IN " & SQL10 & _
            "AND  (" & kaisibi2 & ".���i����) NOT IN " & SQL10_2 & _
            "AND  (" & kaisibi2 & ".���i����) NOT IN " & SQL10_3 & _
            "AND (" & kaisibi2 & ".�⏕�Ȗږ���) = '�L���`�����l��');"
'Debug.Print SQL10_1
db.Execute SQL10_1


'�V�K�`�F�b�N�P�i�ǉ��V�K�����m�j
    SQL5_1 = "UPDATE SYOMEN " & _
            "SET SYOMEN.Shinki = '�V�K' " & _
            ",SYOMEN.F = TRUE " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "'" & _
            "AND SYOMEN.Sv IN (" & SQL5 & ")"
            'Debug.Print SQL5_1
            db.Execute SQL5_1
            
    SQL5_2 = "UPDATE SYOMEN " & _
            "SET SYOMEN.Shinki = '�p��' " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "'" & _
            "AND SYOMEN.Sv NOT IN (" & SQL5 & ")"
            'Debug.Print SQL5_2
            db.Execute SQL5_2

'�V�K�`�F�b�N�Q�i���S�V�K�����m�j
    SQL8 = "SELECT " & kaisibi3 & ".���Ƌ敪_C FROM " & kaisibi3 & " " & _
            "WHERE " & kaisibi3 & ".�ڋqCD = '" & StringKey & "' "

    Debug.Print SQL8
'    Dim kaisibi3 As String '�O�X�� �T�u�N�G������(�O���̌_��Ɣ�r���V�K�̂ݒ��o)
        Set rs_SS_CK = CurrentDb.OpenRecordset(SQL8)
            
            If rs_SS_CK.EOF = True Then
                SQL8_1 = "UPDATE SYOMEN " & _
            "SET SYOMEN.Shinki = '���S�V�K' " & _
            ",SYOMEN.F = TRUE " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "'"
            db.Execute SQL8_1
            End If
SQL9 = "SELECT " & kaisibi2 & ".�[���^�^�C�v FROM " & kaisibi2 & " " & _
            "WHERE " & kaisibi2 & ".�ڋqCD = '" & StringKey & "' " & _
            "AND " & kaisibi2 & ".�[���^�^�C�v = 'V1';"
Set RS_SQL9 = CurrentDb.OpenRecordset(SQL9)

SQL9_1 = "SELECT " & kaisibi3 & ".�[���^�^�C�v FROM " & kaisibi3 & " " & _
            "WHERE " & kaisibi3 & ".�ڋqCD = '" & StringKey & "' " & _
            "AND " & kaisibi3 & ".�[���^�^�C�v = 'V1';"
Set RS_SQL9_1 = CurrentDb.OpenRecordset(SQL9_1)
If RS_SQL9.EOF = False And RS_SQL9_1.EOF = True Then

'Debug.Print "FTTF�ֈڍs"
SQL9_2 = "UPDATE SYOMEN " & _
            "SET SYOMEN.F = TRUE " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "' "
            db.Execute SQL9_2
End If

'�L�����y�[���I���`�F�b�N
If cpck = 1 Then
db.Execute SQL_CP
End If

    'Debug.Print NiraiSinki
    sinki = ""
    NiraiSinki = ""
    NiraiSinki1 = ""
    cpck = 0
    KS = ""
     
Next

End Sub
Sub taisyo2()
'�Ώ�������������l���T�[�`
'�ڋq�ԍ��ŃO���[�v���A�Ԃ�U��iDICTIONARY�֐����g�p�j
'�H���\������������邩�i�S�ē����Ȃ�Ώ����͈�ԑ傫�������j�Ⴄ�Ȃ�ELSE

Dim db As Database
Dim rs As Recordset
Dim rs_count As Integer

'�Q�Ɨp
Set db = CurrentDb()
Set rs = db.OpenRecordset("Q_TA_H_K")

'�������ݗp
Set rs_syomen = db.OpenRecordset("SYOMEN", dbOpenDynaset)

rs.MoveLast
rs.MoveFirst
rs_count = rs.RecordCount
Dim Myarray
Myarray = rs.GetRows(rs_count) '���R�[�h��2�����z��
    
'�s�����ւ�(�Y�������O���P�ɕύX�j
ReDim Myarray2(1 To UBound(Myarray, 2) + 1, 1 To UBound(Myarray, 1) + 3)
For i = 1 To UBound(Myarray, 1) + 1
    For J = 1 To UBound(Myarray, 2) + 1
        Myarray2(J, i) = Myarray(i - 1, J - 1)
    Next
Next
    
'Dictionary�̃I�u�W�F�N�g���쐬
Dim a
Set a = CreateObject("Scripting.Dictionary")

'�u�d���v��ۑ�����z��
Dim Data
ReDim Data(UBound(Myarray2, 1), 1 To 1)
'ReDim Preserve Myarray(LBound(Myarray, 1), UBound(Myarray) + 1)
'�u���i�v�̗�����[�v
For i = 1 To UBound(Myarray2, 1)

'�܂��o�^����Ă��Ȃ��ꍇ
s = 0
    If a.exists(Myarray2(i, 2)) = False Then
        a.Add Myarray2(i, 2), 1 '�o�^�i1�͓K���j
        s = 1
        'Debug.Print Myarray2(i, 1) & ":" & s
        'Myarray2(i, 32) = s
    '���ɓo�^����Ă���ꍇ
    Else '(26,i)(27,i)�g�p
        s = s + 1
        'Debug.Print Myarray2(i, 1) & ":" & s
        'Myarray2(i, 32) = s
    End If

Next


Dim StringKey As String
Dim MaxDay As Variant
'Dictionary�̃I�u�W�F�N�g���쐬
Dim b
Set b = CreateObject("Scripting.Dictionary")
'Dicsionary�̕ϐ������[�v
Keys = a.Keys

    '���[�v�J�n
    For i = 0 To a.Count - 1
        
        'Dic(Keys1����ɂ���)
        Keys1 = Null
        b.RemoveAll
        
        '���t���𕶎���ɕϊ�
        StringKey = Keys(i)
        
        '�ϐ��Ƀt�B���^�������ڋq�ԍ��ŕ���
        Data1 = FilterArray2D(Myarray2, StringKey, 2)
            
            '���������[�v�J�n
            For ii = LBound(Data1) To UBound(Data1)
                
                '�d���`�F�b�N
                If b.exists(Data1(ii, 9)) = False Then
                    
                    'Null�`�F�b�N
                    If IsNull(Data1(ii, 9)) Then
                        
                        'Null����Ȃ���Βǉ�
                        Else
                        b.Add (Data1(ii, 9)), 1
                            
                            If MaxDay = Empty Then 'MaxDay��Null�̏ꍇ�ǉ�
                            MaxDay = CDate(Nz(Data1(ii, 9), 0))
                            Else
                                    '�傫�����t���ϐ��֊i�[
                                    If MaxDay < CDate(Nz(Data1(ii, 9), 0)) Then '�G���[�o���ꍇ��0����t�ɕύX�\��
                                    MaxDay = CDate(Nz(Data1(ii, 9), 0))
                                    End If
                            End If
                    End If
                    
                Else '�d�����Ă��Ȃ��ꍇ�A��r���傫�����Ȃ�Ίi�[
                    
                    If MaxDay < CDate(Nz(Data1(ii, 9), 0)) Then '�G���[�o���ꍇ��0����t�ɕύX�\��
                    '�㏑��
                    MaxDay = CDate(Nz(Data1(ii, 9), 0))
                    End If
                    
                End If
            Next ii
            
            '�������ݏ���
            Keys1 = b.Keys
            For z = 1 To UBound(Myarray2)
                
                '�t�B���^���ꂽDic�ɍő���t������������(�e�[�u���ɏ������ޗl�ɕύX)
                If Myarray2(z, 2) = StringKey Then
                Myarray2(z, 33) = Format(MaxDay, "yyyy/mm/dd")
                End If
            Next
            MaxDay = Empty '�N���A������Dic������
    
    Next
    
End Sub

'FilterArray2D     �E�E�E���ꏊ�FFukamiAddins3.ModArray
'CheckArray2D      �E�E�E���ꏊ�FFukamiAddins3.ModArray
'CheckArray2DStart1�E�E�E���ꏊ�FFukamiAddins3.ModArray

Public Function FilterArray2D(Array2D, FilterStr As String, TargetCol As Long)
'�񎟌��z����w���Ńt�B���^�[�����z����o�͂���B
'20210929

'����
'Array2D  �E�E�E�񎟌��z��
'FilterStr�E�E�E�t�B���^�[���镶���iString�^�j
'TargetCol�E�E�E�t�B���^�[�����iLong�^�j
    
    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Myarray2")
    Call CheckArray2DStart1(Array2D, "Myarray2")
    
    '�t�B���^�[�����v�Z
    Dim i           As Long
    Dim J           As Long
    Dim K           As Long
    Dim M           As Long
    Dim N           As Long
    Dim FilterCount As Long
    Dim TargetStr   As String
    N = UBound(Array2D, 1)
    M = UBound(Array2D, 2)
    K = 0
    For i = 1 To N
        TargetStr = Array2D(i, TargetCol)
        If TargetStr = FilterStr Then
            K = K + 1
        End If
    Next i
    
    FilterCount = K
    
    If K = 0 Then
        '�t�B���^�[�ŉ���������Ȃ������ꍇ��Empty��Ԃ�
        FilterArray2D = Empty
        Exit Function
    End If
    
    '�o�͂���z��̍쐬
    Dim Output
    ReDim Output(1 To FilterCount, 1 To M)
    
    K = 0
    For i = 1 To N
        TargetStr = Array2D(i, TargetCol)
        If TargetStr = FilterStr Then
            K = K + 1
            For J = 1 To M
                Output(K, J) = Array2D(i, J)
            Next J
        End If
    Next i
    
    '�o��
    FilterArray2D = Output
    
End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName As String = "�z��")
'���͔z��2�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy2 As Integer
    Dim Dummy3 As Integer
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "��2�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName As String = "�z��")
'����2�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Sub csvtorikomi()
    
    Dim CsvName As String
    Dim CsvPath As String
    
    CsvName = InputBox("�H������������͂��Ă��������BYYYYMMDD")
    '�ΏۂƂȂ�f�B���N�g������t���{�g���qCSV�Ń��[�v�������N�e�[�u�����쐬����
    'CsvPath = GetFileName("\\192.168.10.1\catv\���[�֌W\�ɂ炢�ł�-�Ɩ��Ǘ�\�d�C�����̂��m�点", link_table)
    CsvPath_ALL = ("C:\Users\OCN060605102\Desktop\��ƃt�H���_\���ʌ�t\csv\bqrcsv" & "\���ʌ�tALL" & CsvName & ".csv")
    CsvPath_DAY = ("C:\Users\OCN060605102\Desktop\��ƃt�H���_\���ʌ�t\csv\bqrcsv" & "\���ʌ�tDAY" & CsvName & ".csv")
    Debug.Print CsvPath_ALL
    Debug.Print CsvPath_DAY
    
    '�t�@�C���̑��݊m�F
    
    '�����N�e�[�u���쐬
    DoCmd.TransferText acLinkDelim, link_table_teigi, link_table_name, CsvPath, True
    DoCmd.TransferText acLinkDelim, link_table_teigi, link_table_name, CsvPath, True
    
    
    Dim dbs As Database
    Dim dtf As TableDef
    
    '�C���X�^���X����
    Set dbs = CurrentDb
    Set dtf = dbs.TableDefs("���l")
    
    '�ڑ������m�F�i�f�o�b�O�v�����g�j
    Debug.Print dtf.TableName
    Debug.Print dtf.Connect
    
    '��n��
    Set dtf = Nothing
    Set dbs = Nothing

    
End Sub

Function zenjitu(dt) As String '�e�L�X�g���t����t�f�[�^�֕ϊ���String�̑O����result
    da = CDate(dt)
    dt = DateAdd("d", -1, da)
    zenjitu = dt
    'Debug.Print dt
End Function

Function zenjitu1(ByVal dt As String) As String '�e�L�X�g���t����t�f�[�^�֕ϊ���String�̑O����result
    da = CDate(dt)
    dt = DateAdd("d", -1, da)
    dt = Format(dt, "YYYYMMDD")
    zenjitu1 = dt
    'Debug.Print dt
End Function
Function zenjitu2(ByVal dt As String) As String '�e�L�X�g���t����t�f�[�^�֕ϊ���String�̑O�X����result
    da = CDate(dt)
    dt = DateAdd("d", -2, da)
    dt = Format(dt, "YYYYMMDD")
    zenjitu2 = dt
    'Debug.Print dt
End Function

Function zenjitu3(ByVal dt As String) As String '�e�L�X�g���t����t�f�[�^�֕ϊ���String�̑O����result
    da = CDate(dt)
    dt = DateAdd("d", -1, da)
    dt = Format(dt, "YYYY/MM/DD")
    zenjitu3 = dt
    'Debug.Print dt
End Function
