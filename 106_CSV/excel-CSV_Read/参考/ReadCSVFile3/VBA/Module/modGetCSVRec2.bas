Attribute VB_Name = "modGetCSVRec2"
'*******************************************************************************
'   CSV形式ファイル読み込み処理(カンマ数不定処理)   ※FSO用(2008/02/03修正)
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'*******************************************************************************
' [参照設定]
'   ・Microsoft Scripting Runtime
'*******************************************************************************
Option Explicit
Const cnsDQ = """"
Const cnsDQCM = ""","
Const cnsCOM = ","
Const cnsBLNK = ""
Const cnsPROD = "."

'*******************************************************************************
' CSV形式の1レコードの受け取り  ※現時点ではエラー処理なし
' (引数はTextStreamオブジェクト,戻り値はCSVレコード内容の1次配列)
'*******************************************************************************
Public Function FP_GET_CSV_REC2(objTS As TextStream) As Variant
    Dim strREC As String        ' レコード内容(連結後)
    Dim strREC2 As String       ' レコード内容(連結前)
    Dim strTEXT As String       ' 1項目の内容
    Dim strTEXT2 As String      ' Work
    Dim POS As Long             ' 項目先頭カラム
    Dim POS2 As Long            ' 項目間のカンマ位置
    Dim X() As Variant          ' レコード内容WORK
    Dim IX As Long              ' 項目INDEX
    Dim lngLEN As Long          ' レコード長
    Dim swTRUE As Boolean       ' レコード連結判断スイッチ
    Dim swDQ As Byte            ' ダブルクォーテーションスイッチ
    Dim lngERR As Long          ' エラーコード
    Dim dteDate As Date         ' 日付試験用Work

    strREC = ""
    ' 項目途中改行対応のため､行末判定は独自に行なう
    Do Until swTRUE = True
        
        swTRUE = True
        ' レコードの読込み(TextStream)
        strREC2 = objTS.ReadLine
        ' 分断レコード対応のため文字列を接合する
        strREC = strREC & strREC2
        lngLEN = Len(strREC)
        
        ' 配列を初期化
        IX = -1
        ReDim X(0)
        ' レコード内容を1文字ずつ判定する
        POS = 1
        Do While POS <= lngLEN
            
            POS2 = POS + 1
            swDQ = 0
            If Mid$(strREC, POS, 1) = cnsDQ Then
                swDQ = 1
                ' 先頭がダブルクォーテーションの場合､項目末の同文字を探す
                Do While POS2 < lngLEN
                    If Mid$(strREC, POS2, 2) = cnsDQCM Then Exit Do
                    POS2 = POS2 + 1
                Loop
                
                If POS2 >= lngLEN Then
                    ' 行末に達した場合は正しい文字列か判定する
                    If Right$(strREC, 1) = cnsDQ Then
                        strTEXT = Trim$(Mid$(strREC, POS + 1, lngLEN - POS - 1))
                    Else
                        ' 不揃いの場合は次レコードを読み込むように指示する
                        strTEXT = cnsBLNK
                        swTRUE = False
                    End If
                ElseIf POS2 > (POS + 1) Then
                    ' 両端のダブルクォーテーションを外す
                    strTEXT = Trim$(Mid$(strREC, POS + 1, POS2 - POS - 1))
                Else
                    strTEXT = cnsBLNK
                End If
                POS2 = POS2 + 1
            ElseIf Mid$(strREC, POS, 1) = cnsCOM Then
                ' カンマのみの場合はEmptyをセットさせる
                strTEXT = ""
                POS2 = POS2 - 1
            Else
                ' 先頭がダブルクォーテーションでない場合は単純にカンマを探す
                Do While POS2 <= lngLEN
                    If Mid$(strREC, POS2, 1) = cnsCOM Then Exit Do
                    POS2 = POS2 + 1
                Loop
                If POS2 > POS Then
                    strTEXT = Trim$(Mid$(strREC, POS, POS2 - POS))
                Else
                    strTEXT = cnsBLNK
                End If
            End If
            
            ' テーブル要素数を追加して内容をｾｯﾄ
            IX = IX + 1
            ReDim Preserve X(IX)
            strTEXT2 = StrConv(strTEXT, vbUpperCase)
            If ((IsNumeric(strTEXT) = True) And (swDQ <> 1)) Then
                ' 数値でダブルクォーテーションで囲われていない
                If InStr(1, strTEXT, cnsPROD, vbTextCompare) <> 0 Then
                    X(IX) = CDbl(strTEXT)   ' 実数は浮動小数点型
                Else
                    X(IX) = CCur(strTEXT)   ' 整数は通貨型
                End If
            ElseIf IsDate(strTEXT) Then
                X(IX) = CDate(strTEXT)      ' 日付型
                On Error Resume Next
                dteDate = X(IX)
                lngERR = Err.Number
                On Error GoTo 0
                If lngERR <> 0 Then
                    ' 日付エラー!
                    X(IX) = strTEXT         ' 文字列型に変更
                ElseIf dteDate < #1/1/1900# Then
                    ' 日付範囲外
                    X(IX) = strTEXT         ' 文字列型に変更
                Else
                    X(IX) = dteDate
                End If
            ElseIf ((strTEXT2 = "TRUE") Or (strTEXT2 = "FALSE")) Then
                X(IX) = CBool(strTEXT)      ' Boolean型
            ElseIf strTEXT <> cnsBLNK Then
                X(IX) = strTEXT             ' 文字列型
            Else
                ' ブランクの場合は初期化(Empty)
                X(IX) = Empty
            End If
            
            POS = POS2 + 1
        Loop
        
        ' EOFの場合は無条件に終了とする
        If objTS.AtEndOfStream Then swTRUE = True
    
    Loop
    
    ' 配列を戻り値にセット
    FP_GET_CSV_REC2 = X
End Function

'-----------------------------<< End of Source >>-------------------------------

