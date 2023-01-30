Attribute VB_Name = "Module2"
Option Compare Database

Sub cd_cut()

'###########################################################
'全顧客番号","区切りつなげ
'テキストファイルに記入し開く
'###########################################################

Dim db As Database
Dim rs As Recordset
Dim i As String
Dim ii As String
Open "C:\書面交付\BQR.txt" For Output As #1

Set db = CurrentDb()
Set rs = db.OpenRecordset("Q_重複カット")
Set WSH = CreateObject("Wscript.Shell")

'先頭レコードに移動する
rs.MoveFirst
i = rs!顧客CD

Do Until rs.EOF
    i = rs!顧客CD
    Print #1, i & ",";
    rs.MoveNext
Loop

'Print #1, Right(i, Len(i) - 1)
Close #1
WSH.Run "C:\書面交付\BQR.txt", 3 'テキストにて開く
End Sub

Function bikou_cut(Serch_str As String) As String

If IsNull(Serch_str) Then
bikou_cut = "null"
Exit Function
End If

'Serch_str = "ヒカリにらい標準工事費0円ｷｬﾝﾍﾟｰﾝ（2024/3月末まで休止・解約不可）4KSTB2年定期契約（2022/4月工事）"
b = InStr(1, Serch_str, "ヒカリにらい標準工事費0円ｷｬﾝﾍﾟｰﾝ", 1)
'Debug.Print B
If b = 0 Then
bikou_cut = "0"
Else
    bikou_cut = Mid(Serch_str, b, 39)
End If
'Debug.Print bikou_cut
End Function
