Attribute VB_Name = "Module2"
Option Compare Database

Sub cd_cut()

'###########################################################
'�S�ڋq�ԍ�","��؂�Ȃ�
'�e�L�X�g�t�@�C���ɋL�����J��
'###########################################################

Dim db As Database
Dim rs As Recordset
Dim i As String
Dim ii As String
Open "C:\���ʌ�t\BQR.txt" For Output As #1

Set db = CurrentDb()
Set rs = db.OpenRecordset("Q_�d���J�b�g")
Set WSH = CreateObject("Wscript.Shell")

'�擪���R�[�h�Ɉړ�����
rs.MoveFirst
i = rs!�ڋqCD

Do Until rs.EOF
    i = rs!�ڋqCD
    Print #1, i & ",";
    rs.MoveNext
Loop

'Print #1, Right(i, Len(i) - 1)
Close #1
WSH.Run "C:\���ʌ�t\BQR.txt", 3 '�e�L�X�g�ɂĊJ��
End Sub

Function bikou_cut(Serch_str As String) As String

If IsNull(Serch_str) Then
bikou_cut = "null"
Exit Function
End If

'Serch_str = "�q�J���ɂ炢�W���H����0�~����߰݁i2024/3�����܂ŋx�~�E���s�j4KSTB2�N����_��i2022/4���H���j"
b = InStr(1, Serch_str, "�q�J���ɂ炢�W���H����0�~����߰�", 1)
'Debug.Print B
If b = 0 Then
bikou_cut = "0"
Else
    bikou_cut = Mid(Serch_str, b, 39)
End If
'Debug.Print bikou_cut
End Function
