Attribute VB_Name = "Module1"
Option Compare Database
Sub Importtest()
Dim d, s
'testt
's = "Path = ""C:\Users\OCN060605102\Desktop\作業フォルダ\書面交付\csv\bqrcsv\書面交付ALL20221026.csv"""
'd = "Path = ""C:\Users\OCN060605102\Desktop\作業フォルダ\書面交付\csv\bqrcsv\書面交付ALL.csv"""

Dim Prj As CodeProject
Dim Obj As ImportExportSpecification
Dim ImpName As String, strXML As String

'パス名を確認するインポート操作の名前を指定
ImpName = "インポート-書面交付ALL"

'XMLプロパティの値をイミディエイト ウィンドウに表示
Set Prj = CurrentProject
Set Obj = Prj.ImportExportSpecifications(ImpName)
strXML = Obj.XML
Debug.Print strXML

'path取得し変更-----------------------
Debug.Print InStr(strXML, "Path") + 8
'Debug.Print InStr(strXML, "xmlns")
path_c = (InStr(strXML, "xmlns") - 2) - (InStr(strXML, "Path") + 8)
path_cc = Mid(strXML, InStr(strXML, "Path") + 8, path_c)

'変更前
s = "Path = """ & path_cc & """"
'変更後
d = "Path = ""C:\Users\OCN060605102\Desktop\作業フォルダ\書面交付\csv\bqrcsv\書面交付ALL20221025.csv"""
strXML = Replace(strXML, s, d)
Debug.Print strXML
Obj.XML = strXML
'path取得し変更-----------------------

'取込
DoCmd.RunSavedImportExport "インポート-書面交付ALL"

'テーブル名変更(書面交付ALL→書面交付ALLYYYYMMDD)
DoCmd.Rename "書面交付ALL20221025", acTable, "書面交付ALL"


End Sub
'対処日を元に設置日を抜き出す○
'対処日が複数ある場合○
'・工事予定日がすべて同じか？クエリ
'    端末番号に同じものがあるか（YES：同じもので、交換日が後を設置日。NO：対処日は後を設置日）
'
'
'どちらかに工事日（設置日）が入っている場合、対象商品名称にfld_a="○"追加
'[2022/06]
'
'・設置完了日（新規）
'・サービス開始日（移行など）
'
'コース変更チェック
'・サービス開始日に変更があるか
'
'サービス終了日チェック
'・あり（対象となるサービスはキャンペーン中となる）
'・なし
'
'契約成立日
'・商品名称の無線モデムは省く
'・fld_a="○"の契約申込日を参照
'
'申込番号

Sub table_create()
    'table存在確認、あればクリア
    'table作成準備(オートナンバーのインデックスあり)
    
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
    
    
    Set tb = db.CreateTableDef(TABLE_NAME) 'テーブル作成
    tb.Fields.Append tb.CreateField("ID", dbLong) 'ID
    'PrimaryKeyを設定
    Set idx = tb.CreateIndex("PrimaryKey")
    idx.Primary = True
    'インデックスを構成する「ID」フィールドを作成する
    idx.Fields.Append idx.CreateField("ID")
    tb.Indexes.Append idx '作成したインデックスをコレクションに追加する
   
    tb.Fields.Append tb.CreateField("C_id", dbText, 9) '顧客番号
    tb.Fields.Append tb.CreateField("Name", dbText, 50) '氏名
    tb.Fields.Append tb.CreateField("Tanmatu", dbText, 50) '端末番号
    tb.Fields.Append tb.CreateField("Syouhin", dbText, 50) '商品名称
    tb.Fields.Append tb.CreateField("Hojyo", dbText, 50) '補助科目名称
    tb.Fields.Append tb.CreateField("Nebiki", dbText, 50) '値引金額
    tb.Fields.Append tb.CreateField("Seikyuu", dbText, 50) '請求金額
    tb.Fields.Append tb.CreateField("Teijyou", dbText, 50) '定常値引額
    tb.Fields.Append tb.CreateField("Keiyakubi_u", dbText, 50) '契約(受付)
    tb.Fields.Append tb.CreateField("Keiyakubi_m", dbText, 50) '契約（申込）
    tb.Fields.Append tb.CreateField("Kaisibi", dbText, 50) 'サービス開始日
    tb.Fields.Append tb.CreateField("Shinki", dbText, 50) '新規ck
    tb.Fields.Append tb.CreateField("setwari", dbText, 50) 'セット割
    tb.Fields.Append tb.CreateField("Sv", dbText, 50) 'サービス
    tb.Fields.Append tb.CreateField("CP_END", dbText, 50) 'CP終了
     tb.Fields.Append tb.CreateField("F", dbBoolean, 10) '書面有無
    
    'IDフィールドをオートナンバーにする
    tb.Fields("ID").Attributes = dbAutoIncrField
    'テーブルを追加
    db.TableDefs.Append tb
    db.Close
    
    
End Sub
Sub ImportALLDAY()
    DoCmd.RunSavedImportExport "インポート-書面交付ALL"
    'DoCmd.RunSavedImportExport "インポート-書面交付DAY"
End Sub

Private Sub TextImport()

'変数宣言
Dim FilePath As String 'ファイルパス
Dim XDAY As String '取込日

XDAY = "20221101"
'CSVファイル名


'csvインポート元のファイルパス
FilePath = "C:\Users\OCN060605102\Desktop\作業フォルダ\書面交付\csv\bqrcsv\書面交付ALL" & XDAY & ".csv"

'csvファイルのインポート
DoCmd.TransferText acImportDelim, "書面交付DAY インポート定義", "書面交付ALL" & XDAY, FilePath

'csvファイルをインポートした旨を通知する。
MsgBox "csvファイルをインポートしました。"

End Sub

Sub taisyo1()
'Call ImportALLDAY
Call table_create
'対処日の取込
'取込後対象顧客番号の対処日でループ取込時にIFでstbレンタルがあり、かつ対の基本デジタルが無い場合対となる行を挿入
'書面交付_重複カットをループ
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
    Dim kaisibi2 As String '当日
    Dim kaisibi3 As String '前日
    Dim kaisibi4 As String '前々日
    Dim kaisibi5 As String
    Dim F As Boolean
    Dim KS As String
    kaisibi = "2022/10/24"
    kaisibi1 = Replace(kaisibi, "/", "")
    kaisibi2 = "書面交付ALL" & kaisibi1
    kaisibi3 = "書面交付ALL" & zenjitu1(kaisibi)
    kaisibi4 = "書面交付ALL" & zenjitu2(kaisibi) '前々日
    kaisibi5 = zenjitu3(kaisibi) '前日YYYY/MM/DD
    'Debug.Print kaisibi5
    '========================================
    '書面交付対象の日（開始日）
    SQL1 = "SELECT " & kaisibi2 & ".* FROM " & kaisibi2 & " WHERE " & kaisibi2 & ".サービス開始日 = '" & kaisibi & "' "
    'Debug.Print SQL1
    'Set rs = db.OpenRecordset("書面交付DAY")
    Set rs = CurrentDb.OpenRecordset(SQL1)
    
    '書面交付対象の日（顧客番号）
    SQL = "SELECT " & kaisibi2 & ".* FROM " & kaisibi2 & " WHERE " & kaisibi2 & ".顧客CD = '" & kokyaku & "' "
    
    Set RS_ALL = CurrentDb.OpenRecordset(SQL)
    

    
    '書き込み用
    Set rs_syomen = db.OpenRecordset("SYOMEN", dbOpenDynaset)
    
    rs.MoveLast
    rs.MoveFirst
    rs_count = rs.RecordCount
    
    Myarray = rs.GetRows(rs_count) 'レコードを2次元配列
    
    Dim Myarray2
    '行列入れ替え(添え字を０→１に変更）
    ReDim Myarray2(1 To UBound(Myarray, 2) + 1, 1 To UBound(Myarray, 1) + 1)
    For i = 1 To UBound(Myarray, 1) + 1
        For J = 1 To UBound(Myarray, 2) + 1
            Myarray2(J, i) = Myarray(i - 1, J - 1)
        Next
    Next

'---------------------------------


'---------------------------------


Dim Myarray6
'Dictionaryのオブジェクトを作成
Dim a
Set a = CreateObject("Scripting.Dictionary")
    
    
    '変数にフィルタ処理し顧客番号で分解
    'Data1 = FilterArray2D(Myarray2, StringKey, 2)

'「重複」を保存する配列
Dim Data
ReDim Data(UBound(Myarray2, 1), 1 To 1)
'Dim StringKey As String
'ReDim Preserve Myarray(LBound(Myarray, 1), UBound(Myarray) + 1)

For i = 1 To UBound(Myarray2, 1) '顧客番号一覧を取得
    If a.exists(Myarray2(i, 2)) = False Then
    a.Add Myarray2(i, 2), 2 '登録する（1はなんでもいい）
    Else '使用されている場合
    End If
Next
Keys = a.Keys
items = a.items

Dim myRegEx As Object '正規化変数
Set myRegEx = CreateObject("VBScript.RegExp") 'RegExオブジェクトを準備
myRegEx.pattern = "^(3Z0|3Z2)\w{2}"
Dim regCheck As Boolean
Dim iii As Long
For ii = 0 To a.Count - 1 '重複カットの顧客番号のループ開始
 
        'Debug.Print Keys(i)
        StringKey = Keys(ii)
        
        '対象となる対処日を変数へ格納
        
        
        Data1 = FilterArray2D(Myarray2, StringKey, 2) '対象の顧客番号のみループ処理,配列,キー,配列のカラム順番
        
        '準備---------------------------
        
        
    '書面交付DAY(対象顧客の全てを抽出)
    SQL1 = "SELECT " & kaisibi2 & ".* FROM " & kaisibi2 & " WHERE " & kaisibi2 & ".顧客CD = '" & StringKey & "'"
            Debug.Print SQL1
    'Set rs = db.OpenRecordset("書面交付DAY")
    'SQL1 = "SELECT 書面交付ALL.* FROM 書面交付ALL WHERE 書面交付ALL.サービス開始日 = '" & kaisibi & "' "
    Set rs_a = CurrentDb.OpenRecordset(SQL1)
    rs_a.MoveLast
    rs_a.MoveFirst
    rs_count_a = rs.RecordCount
    
    Myarray33 = rs_a.GetRows(rs_count_a) 'レコードを2次元配列
    
    Dim Myarray3
    '行列入れ替え(添え字を０→１に変更）
    ReDim Myarray3(1 To UBound(Myarray33, 2) + 1, 1 To UBound(Myarray33, 1) + 1)
    For i = 1 To UBound(Myarray33, 1) + 1
        For J = 1 To UBound(Myarray33, 2) + 1
            Myarray3(J, i) = Myarray33(i - 1, J - 1)
        Next
    Next
    
    Data2 = FilterArray2D(Myarray3, StringKey, 2) '対象の顧客番号の全て
        
    '--------------------------
        
        
    '書面交付対象日の前々日　CP終了比較用（顧客番号）
    'SQL4 = "SELECT " & kaisibi4 & ".* FROM " & kaisibi4 & " WHERE " & kaisibi4 & ".顧客CD = '" & StringKey & "' & " And kaisibi4 & ".サービス終了日 = '" & kaisibi5 & "' "
    SQL4 = "SELECT " & kaisibi4 & ".* FROM " & kaisibi4 & " " & _
            "WHERE " & kaisibi4 & ".顧客CD = '" & StringKey & "' " & _
            "And " & kaisibi4 & ".サービス終了日 = '" & kaisibi5 & "' "
    'Debug.Print SQL4
    Set rs_CPCK = CurrentDb.OpenRecordset(SQL4)
        
'        rs_CPCK.MoveLast
'        rs_CPCK.MoveFirst
'        rs_CPCK_count = rs_CPCK.RecordCount
            If rs_CPCK.EOF Then
            Debug.Print "通常工事"
            Else
'CPチェック
    SQL_CP = "UPDATE SYOMEN " & _
            "SET SYOMEN.CP_END = '" & kaisibi5 & "' " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "'"
            cpck = 1
            Debug.Print SQL_CP
            'db.Execute SQL_CP(実行は最後に)
            
'                Debug.Print "前日CP終了-" & "顧客番号：" & StringKey
'                            rs_syomen.AddNew
'                            rs_syomen!C_id = StringKey
'                            rs_syomen!Name = "-CP終了-"
'                            rs_syomen!CP_END = kaisibi5
'                            rs_syomen.Update
                            
            End If
            '書面交付対象日　サービス開始日＝契約受付日を比較。"＝"で契約書なし。コース変更。なので書面交付無し(要チェック)
'    SQL5 = "SELECT " & kaisibi2 & ".* FROM " & kaisibi2 & " " & _
'            "WHERE " & kaisibi2 & ".顧客CD = '" & StringKey & "' " & _
'            "And " & kaisibi2 & ".契約受付日 = '" & kaisibi & "' "

    SQL5 = "SELECT DISTINCT " & kaisibi2 & ".事業区分_C FROM " & kaisibi2 & " " & _
            "WHERE " & kaisibi2 & ".顧客CD = '" & StringKey & "' " & _
            "And " & kaisibi2 & ".事業区分_C <> '再送' " & _
            "And " & kaisibi2 & ".事業区分_C <> 'OTHERS' " & _
            "And " & kaisibi2 & ".事業区分_C NOT IN " & _
            "(SELECT " & kaisibi3 & ".事業区分_C FROM " & kaisibi3 & " " & _
            "WHERE " & kaisibi3 & ".顧客CD = '" & StringKey & "')"
            
'UPDATE

    Debug.Print SQL5
'    Dim kaisibi3 As String '前々日 サブクエリ条件(前日の契約と比較し新規のみ抽出)
        Set rs_kouji_CK = CurrentDb.OpenRecordset(SQL5)
        
            
            If rs_kouji_CK.EOF And Not cpck = 1 Then
            
            Debug.Print "新規無し"
'                rs_syomen.AddNew
'                rs_syomen!C_id = StringKey
'                rs_syomen!Name = "-書面-"
'                rs_syomen.Update
                F = True
                sinki = "継続"
            Else
            
            Debug.Print "新規"
'                rs_syomen.AddNew
'                rs_syomen!C_id = StringKey
'                rs_syomen!Name = "-NO書面-"
'                rs_syomen.Update


                F = False
                sinki = "新規"
            End If
            
        For iii = 1 To UBound(Data1)
            '前々日にサービス終了日がセットされているか。されているならＣＰ終了なのでスルー
'            If rs_CPCK.EOF Then
'            Debug.Print "あり"
'            End If
            '対処日がイコールか確認
            If Data1(iii, 51) = kaisibi Then
                        If Data1(iii, 31) <> "地上波・ＢＳ" And Data1(iii, 31) <> "BS-NHK" And Data1(iii, 31) <> "BS-民放" Then '地上波・ＢＳは必要なさそうなのでカット（他にもあればここに追加）
                            '項目確定F
                            Debug.Print Data1(iii, 2) & " " & Data1(iii, 3) & " " & Data1(iii, 31) & " " & Data1(iii, 51)
                            NiraiCk = NiraiCk + Data1(iii, 31)
                            'tableへ追加
                            rs_syomen.AddNew
                            rs_syomen!C_id = Data1(iii, 2)
                            rs_syomen!Name = Data1(iii, 3)
                            rs_syomen!Tanmatu = Data1(iii, 28)
                            rs_syomen!Syouhin = Data1(iii, 31)
                            rs_syomen!Keiyakubi_u = Data1(iii, 49) '契約受付日
                            rs_syomen!Keiyakubi_m = Data1(iii, 50) '契約申込日
                            rs_syomen!kaisibi = Data1(iii, 51)
                            rs_syomen!Hojyo = Data1(iii, 45)
                            rs_syomen!Nebiki = Data1(iii, 46)
                            rs_syomen!Seikyuu = Data1(iii, 47)
                            rs_syomen!Teijyou = Data1(iii, 48)
                            rs_syomen!Sv = Data1(iii, 91)
                            rs_syomen.Update
                        End If
                        regCheck = myRegEx.test(Data1(iii, 29))
                        
                If regCheck = True Then 'ＳＴＢレンタルを確認
        '            MsgBox "STBレンタルあり"
                    z = Data1(iii, 28) '端末番号
                    y = Data1(iii, 51) 'サービス開始日
                    '対の基本デジタルチェック
                    For iiii = 1 To UBound(Data2)
                        'If iiii = iii Then '自身はスキップ
                        '    iiii = iiii + 1
                        'Else
                            If Data2(iiii, 28) = z And (Data2(iiii, 45) = "基本デジタル" Or Data2(iiii, 45) = "番組ガイド") Then '同じ端末番号で基本デジタルもしくは番組ガイドである
        '                        MsgBox "同じ端末番号で基本デジタルである"
                                    If Data2(iiii, 51) <> y Then 'サービス開始日が、対の基本デジタルと違う
                                        Debug.Print "STOP"
                                        Debug.Print Data2(iiii, 2) & " " & Data2(iiii, 3) & " " & Data2(iiii, 31) & " " & Data2(iiii, 51)
                                        
                                        'テーブルに追加
                                        rs_syomen.AddNew
                                        rs_syomen!C_id = Data2(iiii, 2)
                                        rs_syomen!Name = Data2(iiii, 3)
                                        rs_syomen!Tanmatu = Data2(iiii, 28)
                                        rs_syomen!Syouhin = Data2(iiii, 31) & "○"
                                        rs_syomen!Keiyakubi_u = Data2(iiii, 49) '契約受付日
                                        rs_syomen!Keiyakubi_u = Data2(iiii, 50) '契約申込日
                                        rs_syomen!kaisibi = Data2(iiii, 51)
                                        rs_syomen!Hojyo = Data2(iiii, 45)
                                        rs_syomen!Nebiki = Data2(iiii, 46)
                                        rs_syomen!Seikyuu = Data2(iiii, 47)
                                        rs_syomen!Teijyou = Data2(iiii, 48)
                                        rs_syomen!Sv = Data2(iiii, 91)
                                        rs_syomen.Update
                                        NiraiCk = NiraiCk + Data2(iiii, 31) & "○"
                                    End If
                                    
                                    
                                
                            End If
                        'End If
                    Next
                End If
            Else
            End If
        Next


    


'にらいパックorにらいプラスでレンタルがなかった場合コースを追加する
If (NiraiCk Like "*にらいパック*" Or NiraiCk Like "*にらいプラス*") And Not NiraiCk Like "*レンタル*" Then
Debug.Print "処理する"
'ALLから対象の顧客番号のみ抽出しその中のサービスのみ加える
'kokyaku = "100711901"
SQL = "SELECT * FROM " & kaisibi2 & " WHERE 顧客CD = '" & StringKey & "' "
Set RS_ALL = CurrentDb.OpenRecordset(SQL)

'count ck
RS_ALL.MoveLast
RS_ALL.MoveFirst
rs_all_count = RS_ALL.RecordCount

For i = 1 To rs_all_count
    
    If RS_ALL.商品名称 = "ポピュラー" Or RS_ALL.商品名称 = "プライム" Or RS_ALL.商品名称 = "ポピュラー2台目以降" Or RS_ALL.商品名称 = "プライム２台目以降" Or RS_ALL.商品_C = "3Z026" Then  'サービス名なら追加する
                                                'テーブルに追加
        rs_syomen.AddNew
        rs_syomen!C_id = RS_ALL!顧客CD '顧客番号
        rs_syomen!Name = RS_ALL!顧客氏名  'サービス名
        rs_syomen!Tanmatu = RS_ALL!端末番号 '端末番号
        rs_syomen!Syouhin = RS_ALL!商品名称 & "　◇" '商品名称
        rs_syomen!Keiyakubi_u = RS_ALL!契約受付日 '加入受付日
        rs_syomen!Keiyakubi_m = RS_ALL!契約申込日 '加入申込日
        rs_syomen!kaisibi = RS_ALL!サービス開始日 'サービス開始日
        rs_syomen!Hojyo = RS_ALL!補助科目名称 '補助
        rs_syomen!Nebiki = RS_ALL!値引金額 '値引
        rs_syomen!Seikyuu = RS_ALL!請求金額 '請求
        rs_syomen!Teijyou = RS_ALL!定常値引額 '定常
        rs_syomen!Sv = RS_ALL!事業区分_C 'サービス
        rs_syomen.Update
        NiraiCk = NiraiCk + RS_ALL!商品名称 & "　◇"
    End If

    RS_ALL.MoveNext
Next

End If



    '割引ck
    Dim wari As String
    wari = ""
        'TV
        Select Case True
            Case NiraiCk Like "*STB*" Or NiraiCk Like "*にらいパック*" Or NiraiCk Like "*にらいプラス*"
            wari = wari + "TV"
        End Select
        'NET
        Select Case True
            Case NiraiCk Like "*パーソナル*" Or NiraiCk Like "*エコノミー*" Or NiraiCk Like "*プレミアム*" Or NiraiCk Like "*プラチナ１２０*" Or NiraiCk Like "*ヒカリにらい*"
            wari = wari + "NET"
        End Select
        'TEL
        Select Case True
            Case NiraiCk Like "*ケーブルプラス*"
            wari = wari + "TEL"
        End Select
    
    'Debug.Print wari
Select Case wari
    Case "TVNET"
'    rs_syomen.AddNew
'    rs_syomen!C_id = StringKey
'    rs_syomen!Tanmatu = "00"
'        rs_syomen!Syouhin = "割引"
'    rs_syomen!setwari = "ダブル"
'    rs_syomen.Update
    
End Select

Select Case wari
    Case "TVTEL"
'    rs_syomen.AddNew
'    rs_syomen!C_id = StringKey
'    rs_syomen!Tanmatu = "00"
'    rs_syomen!Syouhin = "割引"
'    rs_syomen!setwari = "ダブル"
'    rs_syomen.Update
End Select

Select Case wari
    Case "NETTEL"
'    rs_syomen.AddNew
'    rs_syomen!C_id = StringKey
'    rs_syomen!Tanmatu = "00"
'    rs_syomen!Syouhin = "割引"
'    rs_syomen!setwari = "ダブル"
'    rs_syomen.Update
End Select

Select Case wari
    Case "TVNETTEL"
'    rs_syomen.AddNew
'    rs_syomen!C_id = StringKey
'    rs_syomen!Tanmatu = "00"
'    rs_syomen!Syouhin = "割引"
'    rs_syomen!setwari = "トリプル"
'    rs_syomen.Update
End Select

If wari = "" Then
wari = "セット割無し"

End If

NiraiCk = ""
NiraiSinki = ""
NiraiSinki1 = ""
'新規チェック開始

'工事日
SQL2 = "SELECT " & kaisibi2 & ".事業区分_C FROM " & kaisibi2 & " " & _
        "GROUP BY " & kaisibi2 & ".顧客CD, " & kaisibi2 & ".事業区分_C " & _
        "HAVING " & kaisibi2 & ".事業区分_C <> 'OTHERS' AND " & kaisibi2 & ".事業区分_C <> '再送' " & _
        "AND " & kaisibi2 & ".顧客CD = '" & StringKey & "'"
Set RS_SINKI = CurrentDb.OpenRecordset(SQL2)


'工事前日
SQL3 = "SELECT " & kaisibi3 & ".事業区分_C FROM " & kaisibi3 & " GROUP BY " & kaisibi3 & ".顧客CD, " & kaisibi3 & ".事業区分_C HAVING " & kaisibi3 & ".事業区分_C <>'OTHERS' AND " & kaisibi3 & ".事業区分_C <>'再送' AND " & kaisibi3 & ".顧客CD = '" & StringKey & "'ORDER BY " & kaisibi3 & ".事業区分_C;"
'Debug.Print SQL3
Set RS_SINKI1 = CurrentDb.OpenRecordset(SQL3)

    '--------------------------
    '新規チェック(前日レコード)
    If RS_SINKI1.EOF Then
        '完全なる新規契約★
        Debug.Print "完全新規"
'            rs_syomen.AddNew
'            rs_syomen!C_id = StringKey
'            rs_syomen!Shinki = "新規"
'            rs_syomen.Update
            KS = "あり"
        Else
    End If
    '----------------------------






'RECODESET存在チェック
    If RS_SINKI.EOF Then
    
    Else
        RS_SINKI.MoveLast
        RS_SINKI.MoveFirst
        rs_all_sinki_count = RS_SINKI.RecordCount
        
        For ss = 1 To rs_all_sinki_count
            'Debug.Print RS_SINKI.事業区分_C & "【当日】"
            NiraiSinki = NiraiSinki + RS_SINKI.事業区分_C
            RS_SINKI.MoveNext
        Next
        
    End If
   
    

'RECODESET存在チェック
    If RS_SINKI1.EOF Then
    
    Else
        RS_SINKI1.MoveLast
        RS_SINKI1.MoveFirst
        rs_all_sinki_count1 = RS_SINKI1.RecordCount
        
        For ss = 1 To rs_all_sinki_count1
            'Debug.Print RS_SINKI1.事業区分_C & "【前日】"
            NiraiSinki1 = NiraiSinki1 + RS_SINKI1.事業区分_C
            RS_SINKI1.MoveNext
        Next
        RS_SINKI1.MoveFirst
    End If

'-------------------前日ありは引き算
If RS_SINKI1.EOF Then
Else
        'RS_SINKI1.MoveFirst
        RS_SINKI.MoveFirst
        For ss = 1 To rs_all_sinki_count1
            NiraiSinki = Replace(NiraiSinki, RS_SINKI1.事業区分_C, "")
            RS_SINKI1.MoveNext
        Next ss
        
        If NiraiSinki = "" Then
            Debug.Print "新規なし"
        Else
            Debug.Print "新規あり:" & NiraiSinki
        End If
'-------------------
End If
    
    'レコード更新処理
    Select Case NiraiSinki
    Case "NETPTEL"
    
    Case ""
    Case ""
    Case ""
    End Select
    
    Select Case wari
    Case "NETTEL"
    wari = "ダブル"
    Case "TVTEL"
    wari = "ダブル"
    Case "TVNET"
    wari = "ダブル"
    Case "TV"
    wari = "セット割無し"
    Case "NET"
    wari = "セット割無し"
    Case "TEL"
    wari = "セット割無し"
    End Select

    
    SQL6 = "UPDATE SYOMEN " & _
            "SET SYOMEN.F = " & F & " " & _
            ", SYOMEN.setwari = '" & wari & "' " & _
            ", SYOMEN.Shinki = '" & sinki & "' " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "'"
            Debug.Print SQL6
'            db.Execute SQL6
'Set RS_SINKI0 = CurrentDb.OpenRecordset(SQL6)




'オプションCHを追加
SQL10 = "(SELECT DISTINCT SYOMEN.Tanmatu FROM SYOMEN " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "' )"
SQL10_2 = "(SELECT DISTINCT SYOMEN.Syouhin FROM SYOMEN " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "' )"
SQL10_3 = "('BS-NHK', 'BS-民放', 'BS-WOWOW')"
SQL10_1 = "INSERT INTO SYOMEN ( C_id, Name, Tanmatu, Syouhin, Hojyo, Nebiki, Seikyuu, Teijyou, Keiyakubi_u, Keiyakubi_m, Kaisibi, Sv ) " & _
            "SELECT " & kaisibi2 & ".顧客CD, " & kaisibi2 & ".顧客氏名, " & _
            "" & kaisibi2 & ".端末番号, " & kaisibi2 & ".商品名称, " & _
            "" & kaisibi2 & ".補助科目名称, " & kaisibi2 & ".値引金額, " & _
            "" & kaisibi2 & ".請求金額, " & kaisibi2 & ".定常値引額, " & _
            "" & kaisibi2 & ".契約受付日, " & kaisibi2 & ".契約申込日, " & _
            "" & kaisibi2 & ".サービス開始日, " & kaisibi2 & ".事業区分_C " & _
            "FROM " & kaisibi2 & " " & _
            "WHERE (((" & kaisibi2 & ".顧客CD)= '" & StringKey & "') AND (" & kaisibi2 & ".端末番号) IN " & SQL10 & _
            "AND  (" & kaisibi2 & ".商品名称) NOT IN " & SQL10_2 & _
            "AND  (" & kaisibi2 & ".商品名称) NOT IN " & SQL10_3 & _
            "AND (" & kaisibi2 & ".補助科目名称) = '有料チャンネル');"
'Debug.Print SQL10_1
db.Execute SQL10_1


'新規チェック１（追加新規を検知）
    SQL5_1 = "UPDATE SYOMEN " & _
            "SET SYOMEN.Shinki = '新規' " & _
            ",SYOMEN.F = TRUE " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "'" & _
            "AND SYOMEN.Sv IN (" & SQL5 & ")"
            'Debug.Print SQL5_1
            db.Execute SQL5_1
            
    SQL5_2 = "UPDATE SYOMEN " & _
            "SET SYOMEN.Shinki = '継続' " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "'" & _
            "AND SYOMEN.Sv NOT IN (" & SQL5 & ")"
            'Debug.Print SQL5_2
            db.Execute SQL5_2

'新規チェック２（完全新規を検知）
    SQL8 = "SELECT " & kaisibi3 & ".事業区分_C FROM " & kaisibi3 & " " & _
            "WHERE " & kaisibi3 & ".顧客CD = '" & StringKey & "' "

    Debug.Print SQL8
'    Dim kaisibi3 As String '前々日 サブクエリ条件(前日の契約と比較し新規のみ抽出)
        Set rs_SS_CK = CurrentDb.OpenRecordset(SQL8)
            
            If rs_SS_CK.EOF = True Then
                SQL8_1 = "UPDATE SYOMEN " & _
            "SET SYOMEN.Shinki = '完全新規' " & _
            ",SYOMEN.F = TRUE " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "'"
            db.Execute SQL8_1
            End If
SQL9 = "SELECT " & kaisibi2 & ".端末型タイプ FROM " & kaisibi2 & " " & _
            "WHERE " & kaisibi2 & ".顧客CD = '" & StringKey & "' " & _
            "AND " & kaisibi2 & ".端末型タイプ = 'V1';"
Set RS_SQL9 = CurrentDb.OpenRecordset(SQL9)

SQL9_1 = "SELECT " & kaisibi3 & ".端末型タイプ FROM " & kaisibi3 & " " & _
            "WHERE " & kaisibi3 & ".顧客CD = '" & StringKey & "' " & _
            "AND " & kaisibi3 & ".端末型タイプ = 'V1';"
Set RS_SQL9_1 = CurrentDb.OpenRecordset(SQL9_1)
If RS_SQL9.EOF = False And RS_SQL9_1.EOF = True Then

'Debug.Print "FTTFへ移行"
SQL9_2 = "UPDATE SYOMEN " & _
            "SET SYOMEN.F = TRUE " & _
            "WHERE SYOMEN.C_id = '" & StringKey & "' "
            db.Execute SQL9_2
End If

'キャンペーン終了チェック
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
'対処日が複数ある人をサーチ
'顧客番号でグループし連番を振る（DICTIONARY関数を使用）
'工事予定日が複数あるか（全て同じなら対処日は一番大きい数字）違うならELSE

Dim db As Database
Dim rs As Recordset
Dim rs_count As Integer

'参照用
Set db = CurrentDb()
Set rs = db.OpenRecordset("Q_TA_H_K")

'書き込み用
Set rs_syomen = db.OpenRecordset("SYOMEN", dbOpenDynaset)

rs.MoveLast
rs.MoveFirst
rs_count = rs.RecordCount
Dim Myarray
Myarray = rs.GetRows(rs_count) 'レコードを2次元配列
    
'行列入れ替え(添え字を０→１に変更）
ReDim Myarray2(1 To UBound(Myarray, 2) + 1, 1 To UBound(Myarray, 1) + 3)
For i = 1 To UBound(Myarray, 1) + 1
    For J = 1 To UBound(Myarray, 2) + 1
        Myarray2(J, i) = Myarray(i - 1, J - 1)
    Next
Next
    
'Dictionaryのオブジェクトを作成
Dim a
Set a = CreateObject("Scripting.Dictionary")

'「重複」を保存する配列
Dim Data
ReDim Data(UBound(Myarray2, 1), 1 To 1)
'ReDim Preserve Myarray(LBound(Myarray, 1), UBound(Myarray) + 1)
'「商品」の列をループ
For i = 1 To UBound(Myarray2, 1)

'まだ登録されていない場合
s = 0
    If a.exists(Myarray2(i, 2)) = False Then
        a.Add Myarray2(i, 2), 1 '登録（1は適当）
        s = 1
        'Debug.Print Myarray2(i, 1) & ":" & s
        'Myarray2(i, 32) = s
    '既に登録されている場合
    Else '(26,i)(27,i)使用
        s = s + 1
        'Debug.Print Myarray2(i, 1) & ":" & s
        'Myarray2(i, 32) = s
    End If

Next


Dim StringKey As String
Dim MaxDay As Variant
'Dictionaryのオブジェクトを作成
Dim b
Set b = CreateObject("Scripting.Dictionary")
'Dicsionaryの変数をループ
Keys = a.Keys

    'ループ開始
    For i = 0 To a.Count - 1
        
        'Dic(Keys1を空にする)
        Keys1 = Null
        b.RemoveAll
        
        '日付けを文字列に変換
        StringKey = Keys(i)
        
        '変数にフィルタ処理し顧客番号で分解
        Data1 = FilterArray2D(Myarray2, StringKey, 2)
            
            '分解をループ開始
            For ii = LBound(Data1) To UBound(Data1)
                
                '重複チェック
                If b.exists(Data1(ii, 9)) = False Then
                    
                    'Nullチェック
                    If IsNull(Data1(ii, 9)) Then
                        
                        'Nullじゃなければ追加
                        Else
                        b.Add (Data1(ii, 9)), 1
                            
                            If MaxDay = Empty Then 'MaxDayがNullの場合追加
                            MaxDay = CDate(Nz(Data1(ii, 9), 0))
                            Else
                                    '大きい日付け変数へ格納
                                    If MaxDay < CDate(Nz(Data1(ii, 9), 0)) Then 'エラー出た場合は0を日付に変更予定
                                    MaxDay = CDate(Nz(Data1(ii, 9), 0))
                                    End If
                            End If
                    End If
                    
                Else '重複していない場合、比較し大きい日ならば格納
                    
                    If MaxDay < CDate(Nz(Data1(ii, 9), 0)) Then 'エラー出た場合は0を日付に変更予定
                    '上書き
                    MaxDay = CDate(Nz(Data1(ii, 9), 0))
                    End If
                    
                End If
            Next ii
            
            '書き込み処理
            Keys1 = b.Keys
            For z = 1 To UBound(Myarray2)
                
                'フィルタされたDicに最大日付けを書き込む(テーブルに書き込む様に変更)
                If Myarray2(z, 2) = StringKey Then
                Myarray2(z, 33) = Format(MaxDay, "yyyy/mm/dd")
                End If
            Next
            MaxDay = Empty 'クリアし次のDicを検索
    
    Next
    
End Sub

'FilterArray2D     ・・・元場所：FukamiAddins3.ModArray
'CheckArray2D      ・・・元場所：FukamiAddins3.ModArray
'CheckArray2DStart1・・・元場所：FukamiAddins3.ModArray

Public Function FilterArray2D(Array2D, FilterStr As String, TargetCol As Long)
'二次元配列を指定列でフィルターした配列を出力する。
'20210929

'引数
'Array2D  ・・・二次元配列
'FilterStr・・・フィルターする文字（String型）
'TargetCol・・・フィルターする列（Long型）
    
    '引数チェック
    Call CheckArray2D(Array2D, "Myarray2")
    Call CheckArray2DStart1(Array2D, "Myarray2")
    
    'フィルター件数計算
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
        'フィルターで何もかからなかった場合はEmptyを返す
        FilterArray2D = Empty
        Exit Function
    End If
    
    '出力する配列の作成
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
    
    '出力
    FilterArray2D = Output
    
End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName As String = "配列")
'入力配列が2次元配列かどうかチェックする
'20210804

    Dim Dummy2 As Integer
    Dim Dummy3 As Integer
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "は2次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName As String = "配列")
'入力2次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Sub csvtorikomi()
    
    Dim CsvName As String
    Dim CsvPath As String
    
    CsvName = InputBox("工事完了日を入力してください。YYYYMMDD")
    '対象となるディレクトリを日付け＋拡張子CSVでループしリンクテーブルを作成する
    'CsvPath = GetFileName("\\192.168.10.1\catv\帳票関係\にらいでんき-業務管理\電気料金のお知らせ", link_table)
    CsvPath_ALL = ("C:\Users\OCN060605102\Desktop\作業フォルダ\書面交付\csv\bqrcsv" & "\書面交付ALL" & CsvName & ".csv")
    CsvPath_DAY = ("C:\Users\OCN060605102\Desktop\作業フォルダ\書面交付\csv\bqrcsv" & "\書面交付DAY" & CsvName & ".csv")
    Debug.Print CsvPath_ALL
    Debug.Print CsvPath_DAY
    
    'ファイルの存在確認
    
    'リンクテーブル作成
    DoCmd.TransferText acLinkDelim, link_table_teigi, link_table_name, CsvPath, True
    DoCmd.TransferText acLinkDelim, link_table_teigi, link_table_name, CsvPath, True
    
    
    Dim dbs As Database
    Dim dtf As TableDef
    
    'インスタンス生成
    Set dbs = CurrentDb
    Set dtf = dbs.TableDefs("備考")
    
    '接続情報を確認（デバッグプリント）
    Debug.Print dtf.TableName
    Debug.Print dtf.Connect
    
    '後始末
    Set dtf = Nothing
    Set dbs = Nothing

    
End Sub

Function zenjitu(dt) As String 'テキスト日付を日付データへ変換しStringの前日をresult
    da = CDate(dt)
    dt = DateAdd("d", -1, da)
    zenjitu = dt
    'Debug.Print dt
End Function

Function zenjitu1(ByVal dt As String) As String 'テキスト日付を日付データへ変換しStringの前日をresult
    da = CDate(dt)
    dt = DateAdd("d", -1, da)
    dt = Format(dt, "YYYYMMDD")
    zenjitu1 = dt
    'Debug.Print dt
End Function
Function zenjitu2(ByVal dt As String) As String 'テキスト日付を日付データへ変換しStringの前々日をresult
    da = CDate(dt)
    dt = DateAdd("d", -2, da)
    dt = Format(dt, "YYYYMMDD")
    zenjitu2 = dt
    'Debug.Print dt
End Function

Function zenjitu3(ByVal dt As String) As String 'テキスト日付を日付データへ変換しStringの前日をresult
    da = CDate(dt)
    dt = DateAdd("d", -1, da)
    dt = Format(dt, "YYYY/MM/DD")
    zenjitu3 = dt
    'Debug.Print dt
End Function
