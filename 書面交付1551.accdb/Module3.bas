Attribute VB_Name = "Module3"
Option Compare Database
Option Base 1


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





Sub test()
'対処日が複数ある人をサーチ
'顧客番号でグループし連番を振る（DICTIONARY関数を使用）
'工事予定日が複数あるか（全て同じなら対処日は一番大きい数字）違うならELSE
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
Myarray = rs.GetRows(rs_count) 'レコードを2次元配列

'行列入れ替え
ReDim Myarray2(UBound(Myarray, 2) + 1, UBound(Myarray, 1) + 2)
For i = 1 To UBound(Myarray, 1) '何回か？
    For J = 1 To UBound(Myarray, 2)
        Myarray2(J, i) = Myarray(i - 1, J)
    Next
Next
    

    
'Dictionaryのオブジェクトを作成
Dim a
Set a = CreateObject("Scripting.Dictionary")

'「重複」を保存する配列
Dim Data
ReDim Data(1 To UBound(Myarray2, 1), 1 To 1)
'ReDim Preserve Myarray(LBound(Myarray, 1), UBound(Myarray) + 1)
'「商品」の列をループ
For i = 1 To UBound(Myarray2, 1)

'まだ登録されていない場合

    If a.exists(Myarray2(i, 1)) = False Then
        a.Add Myarray2(i, 1), 1 '登録する（1はなんでもいい）
        s = 1
        Debug.Print Myarray2(i, 1) & ":" & s
        Myarray2(i, 32) = s
    '既に登録されている場合
    Else '(26,i)(27,i)使用
        s = s + 1
        Debug.Print Myarray2(i, 1) & ":" & s
        Myarray2(i, 32) = s
        
    End If
            
Next



Debug.Print 二次元配列フィルター関数(Myarray2, 1, 105592501)

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
    '27が1で処理を開始する
    If Myarray(27, i) = 1 Then '対処日CK
        '配列に格納
        
        
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

Dim dic    '// 重複を除いた値を格納するDictionary
Set dic = CreateObject("Scripting.Dictionary")
    
    Dim i                       '// ループカウンタ１
    Dim ii                      '// ループカウンタ２
    Dim iLen                    '// 配列要素数
    Dim arEdit()                '// 編集後の配列
    
    ReDim arEdit(0)
    iLen = UBound(ar)
    
    '// 配列ループ
    For i = 0 To iLen
        '// 配列に未登録の値の場合
        If (dic.exists(ar(i)) = False) Then
            '// Dictionaryに追加
            Call dic.Add(ar(i), ar(i))
            
            '// 重複がない値のみを編集後配列に格納する
            arEdit(UBound(arEdit)) = ar(i)
            ReDim Preserve arEdit(UBound(arEdit) + 1)
        End If
    Next
    
    '// 配列に格納済みの場合
    If (IsEmpty(arEdit(0)) = False) Then
        '// 余分な領域を削除
        ReDim Preserve arEdit(UBound(arEdit) - 1)
    End If
    
    '// 引数に編集後配列を設定
    ar = arEdit
End Function

Function ex1(str0, str2, str4, str5, str6, str7, str10, str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
        'テンプレをコピー
        wb_t.Worksheets("temp_1").Copy Before:=wb.Worksheets("Sheet1")
        
    With wb
        .ActiveSheet.Name = str12
    End With
    
    cc = cc + 1
    With wb.Sheets(str12)
        .Cells(1, 2).Value = "　" & str2                    'm_請求先郵便番号
        .Cells(2, 2).Value = "　" & str4 & str5 & str6      'm_住所
        .Cells(3, 2).Value = "　" & str7                    'm_建物名
        .Cells(5, 2).Value = "　" & str0 & "　様"             'm_請求先名（名前）
        .Cells(7, 2).Value = "　お客様番号：" & str57       's_顧客番号
        .Cells(8, 2).Value = "　伝票番号：" & str12         's_伝票番号
        .Cells(12, 2).Value = "電気料金等のお知らせ　　" & Date_ss(str14)               's_電気料金のお知らせ
        .Cells(15, 4).Value = " " & nk1(str39) & "円"            's_売上金(税込)
        .Cells(15, 5).Value = "（うち消費税相当額　　　　　" & nk1(str40) & "円）"      's_消費税等相当額
        .Cells(18, 4).Value = " " & str57                   's_顧客番号
        .Cells(19, 4).Value = " " & str10                   's_氏名
        .Cells(20, 4).Value = " " & str19                   's_利用場所
        .Cells(22, 4).Value = str20 & " ～ " & str21        's_使用開始日 s_使用終了日
        .Cells(23, 4).Value = Date_k(str22)                 's_検針日
        .Cells(24, 4).Value = str26                         's_契約種別
        .Cells(25, 4).Value = str13                         's_供給地点特定番号
        .Cells(22, 7).Value = str23                         's_日数
        .Cells(23, 7).Value = Date_k(str28)                 's_次回検針日
        .Cells(24, 7).Value = str27                         's_契約容量
        .Cells(25, 7).Value = "口座振替"                    '支払い方法
        .Cells(26, 7).Value = Date_sa(str22)                'ご請求月
        .Cells(29, 2).Value = str15                         's_明細
        .Cells(29, 5).Value = str16                         's_単価
        .Cells(29, 7).Value = int_0(nk(str17))              's_契約電力/電力量)
        .Cells(29, 9).Value = nk(str18) & "円"              's_売上額(税込)
        .Cells(41, 9).Value = "当月            " & nk(touge2) & "円"                    '燃料費調整単価_当月
        .Cells(42, 9).Value = "翌月            " & nk(yokuge2) & "円"                     '燃料費調整単価_翌月
        .Cells(43, 9).Value = "翌月は当月と比べ" & pm(Val(yokuge2) - Val(touge2)) & "円"   '前月と今月の比較値
    End With
   
End Function

Function ex2(str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
            
    With wb.Sheets(str12)
        .Cells(28 + C, 2).Value = str15                     's_明細
        .Cells(28 + C, 5).Value = nk(str16)                     's_単価
        .Cells(28 + C, 7).Value = str17                     's_契約電力/電力量
        .Cells(28 + C, 9).Value = str18 & "円"              's_売上額(税込)
    End With
    
End Function

Sub excell_create()
    Dim db As Database
    Dim rs As Recordset

    touge2 = Form_TOP.tougetu
    yokuge2 = Form_TOP.yokugetu
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Q_sales_member")         '取込み元クエリ
    Dim rs_CT As Long
    Dim C As Long
    excel_file = "C:\Users\OCN060605102\Desktop\作業フォルダ\書面交付\excel\書面交付ご契約内容書.xls"                            'エクエルファイル
    temp_excel_file = "C:\Users\OCN060605102\Desktop\作業フォルダ\書面交付\excel\temp\書面交付ご契約内容書_temp.xls"    'エクセルテンプレート
    Set exAPP = CreateObject("Excel.Application")       'エクセルセット
    'exAPP.Visible = True                               '非表示
    exAPP.DisplayAlerts = False                         '警告無視
    
    Set wb = exAPP.Workbooks.Add                            'エクセルファイル作成
    wb.SaveAs (excel_file)                                  'エクセルファイル指定
    Set wb_t = exAPP.Workbooks.Open(temp_excel_file)        'エクセルテンプレート
    
    
    'ExAPP.Workbooks.Add (excel_file)                      'エクセルファイル作成
    'Set wb = ExAPP.Workbooks.Open(excel_file)
    'wb.Save
    'Set wb = ExAPP.wb.Open(excel_file)
    'Set wb_t = ExAPP.Workbooks.Open(temp_excel_file)        'エクセルテンプレート
    
    cc = 0
    rs.MoveFirst
    rs.MoveLast
    rs.MoveFirst
    
    rs_CT = rs.RecordCount
    Myarray = rs.GetRows(rs_CT)     '２次元配列セット
    


    
    
    'ループで次のレコード条件分岐でセル書き込み
    For i = 0 To (rs_CT - 1)

        str0 = Myarray(0, i) 'm_請求先名
        str1 = Myarray(1, i) 'm_請求先名(カナ)
        str2 = Myarray(2, i) 'm_請求先郵便番号
        str3 = Myarray(3, i) 'm_請求先都道府県
        str4 = Myarray(4, i) 'm_請求先市区町村
        str5 = Myarray(5, i) 'm_請求先住所
        str6 = Myarray(6, i) 'm_請求先番地
        str7 = Myarray(7, i) 'm_請求先建物名
        str8 = Myarray(8, i) 's_会員ID
        str9 = Myarray(9, i) 's_請求先ID
        str10 = Myarray(10, i) 's_氏名
        str11 = Myarray(11, i) 's_会員区分
        str12 = Myarray(12, i) 's_伝票番号（重複確認項目）
        str13 = Myarray(13, i) 's_供給地点特定番号
        str14 = Myarray(14, i) 's_請求年月
        str15 = Myarray(15, i) 's_明細
        str16 = Myarray(16, i) 's_単価
        str17 = Myarray(17, i) 's_契約電力/電力量
        str18 = Myarray(18, i) 's_売上額(税込)
        str19 = Myarray(19, i) 's_利用場所
        str20 = Myarray(20, i) 's_使用開始日
        str21 = Myarray(21, i) 's_使用終了日
        str22 = Myarray(22, i) 's_検針日
        str23 = Myarray(23, i) 's_日数
        str24 = Myarray(24, i) 's_使用量
        str25 = Myarray(25, i) 's_プランID
        str26 = Myarray(26, i) 's_契約種別
        str27 = Myarray(27, i) 's_契約容量
        str28 = Myarray(28, i) 's_次回検針日
        str29 = Myarray(29, i) 's_確定使用量取込日
        str30 = Myarray(30, i) 's_料金算定日
        str31 = Myarray(31, i) 's_請求番号
        str32 = Myarray(32, i) 's_請求状況
        str33 = Myarray(33, i) 's_請求エラー
        str34 = Myarray(34, i) 's_支払方法
        str35 = Myarray(35, i) 's_売上(振替)指定日
        str36 = Myarray(36, i) 's_コンビニ支払期限
        str37 = Myarray(37, i) 's_請求依頼連携日
        str38 = Myarray(38, i) 's_公開日
        str39 = Myarray(39, i) 's_売上金(税込)
        str40 = Myarray(40, i) 's_消費税等相当額
        str41 = Myarray(41, i) 's_まとめ請求額
        str42 = Myarray(42, i) 's_残金
        str43 = Myarray(43, i) 's_入金日
        str44 = Myarray(44, i) 's_入金額
        str45 = Myarray(45, i) 's_入金方法
        str46 = Myarray(46, i) 's_最終入金日
        str47 = Myarray(47, i) 's_督促回数
        str48 = Myarray(48, i) 's_督促ステータス
        str49 = Myarray(49, i) 's_支払期日
        str50 = Myarray(50, i) 's_延滞日数
        str51 = Myarray(51, i) 's_計上日
        str52 = Myarray(52, i) 's_登録者
        str53 = Myarray(53, i) 's_最終更新日
        str54 = Myarray(54, i) 's_請求書郵送
        str55 = Myarray(55, i) 's_メール送信
        str56 = Myarray(56, i) 's_代理店
        str57 = Myarray(57, i) 's_外部キー

        Dim tc As Long
                
        If i = 0 Then '最初
        
                '①次レコード同じならカウント＋１
                If str12 = Myarray(12, i + 1) Then
                    C = C + 1
                    Debug.Print "①"
                    Call ex1(str0, str2, str4, str5, str6, str7, str10, str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    
                    Debug.Print str12
                    Debug.Print str15 & C
                    
                '②次のレコードが違う
                    Else
                    C = 1
                    Debug.Print "②"
                    Call ex1(str0, str2, str4, str5, str6, str7, str10, str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str12
                    Debug.Print str15 & C
                    C = 0
                    tc = tc + 1
                End If
                
        ElseIf i = rs_CT - 1 Then '最後
        
                '③前レコード同じならカウント＋１
                If str12 = Myarray(12, i - 1) Then
                    C = C + 1
                    Debug.Print "③"
                    Call ex2(str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str15 & C
                    
                '④前レコードと違うが最後１カウント
                Else
                    C = 1
                    Debug.Print "④"
                    Call ex1(str0, str2, str4, str5, str6, str7, str10, str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str12
                    Debug.Print str15 & C
                    tc = tc + 1
                End If
                
        Else
                
                '⑤次レコード同じで& C = 0ならカウント＋１
                If str12 = Myarray(12, i + 1) And C = 0 Then
                    C = C + 1
                    Debug.Print "⑤"
                    Call ex1(str0, str2, str4, str5, str6, str7, str10, str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str12
                    Debug.Print str15 & C
                    
                '⑥次レコード同じならカウント＋１
                ElseIf str12 = Myarray(12, i + 1) Then
                    C = C + 1
                    Debug.Print "⑥"
                    Call ex2(str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str15 & C
                End If
                
                '⑦次レコード違うが前レコードと同じ
                If str12 <> Myarray(12, i + 1) And str12 = Myarray(1, i - 1) Then
                    C = C + 1
                    Debug.Print "⑦"
                    Call ex2(str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str15 & C
                    C = 0
                    tc = tc + 1
                
                '⑧次レコード違とうならカウント０
                ElseIf str12 <> Myarray(12, i + 1) Then
                    
                    If str12 = Myarray(12, 0) Then '最初の取込み者を確認
                    C = C + 1
                    Debug.Print "⑨"
                    Call ex2(str12, str13, _
            str14, str15, str16, str17, str18, str19, str20, str21, _
            str22, str23, str26, str27, str28, str39, str40, str57, C)
                    Debug.Print str15 & C
                    Else
                    C = C + 1
                    Debug.Print "⑩" '次の人へ
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
        
    'エクセル終了処理
    With wb
        .Worksheets("Sheet1").Delete
        .SaveAs (excel_file)
    End With
    
    'exAPP.DisplayAlerts = True

    
    Set wb = Nothing            'ワークシート解放
    Set wb_t = Nothing          'ワークシート解放


    exAPP.Quit
    Set exAPP = Nothing         'エクセル解放

'Public exAPP As Object     'オブジェクトエクセル宣言
'Public wb As Object        'オブジェクトワークブック宣言
'Public wb_t As Object      'オブジェクトワークブック宣言


MsgBox "伝票番号の数：" & cc
End Sub

'************************************************************************************************************
'**2次元配列をフィルターする（特定の１列に対してひとつの値に一致するものを抽出
'**第一引数：data => フィルターしたい2次元配列 バリアント型
'**第二引数：col_num => フィルターする列の番号（左から1から数える）整数型
'**第三引数：key_array => 抽出したいキーワード
'**※dataにヘッダーが含まれる前提、ヘッダーも含まれた配列を返す
'**※Option Base 1 を必ず使用した上で利用すること
'*************************************************************************************************************
Function 二次元配列フィルター関数(ByVal Data As Variant, ByVal col_num As Integer, ByVal key As String) As Variant
    Dim cnt As Long 'フィルター後の配列の行数
    Dim n_col As Long '配列の列数
    Dim dic As Object
    Dim r_array As Variant '条件に合う列の各列を一時的に格納するための配列
    Dim data_fil As Variant 'フィルター後の2次元配列
    '元の配列(data)の列数（=フィルター後の列数
    n_col = UBound(Data, 2)
    '上記の列数（変数）を用いてr_arrayを再定義
    ReDim r_array(n_col) As String
    '辞書型配列を定義
    Set dic = CreateObject("Scripting.Dictionary")
    'フィルター後の配列の行数をカウントし、該当した行を配列に格納する
    cnt = 1 'ヘッダー分の列数をあらかじめ考慮
    For i = 2 To UBound(Data)
        If CStr(Data(i, col_num)) = key Then
            '取り出したい行の列を配列に格納する
            For J = 1 To n_col
                r_array(J) = Data(i, J)
            Next J
            cnt = cnt + 1
            dic.Add cnt, r_array
        End If
    Next i
    'ヘッダー部分を辞書型配列に格納
    For J = 1 To n_col
        r_array(J) = Data(1, J)
    Next J
    dic.Add 1, r_array
    'フィルター後の配列の再定義
    ReDim data_fil(cnt, n_col) As String
    'フィルター後の配列にヘッダーを格納
    For J = 1 To n_col
        data_fil(1, J) = dic.Item(1)(J)
    Next J
    'ヘッダー以外の値をフィルター後の配列に格納
    For i = 2 To cnt
        For J = 1 To n_col
            data_fil(i, J) = dic.Item(i)(J)
        Next J
    Next i
    '関数に代入
    二次元配列フィルター関数 = data_fil
End Function


