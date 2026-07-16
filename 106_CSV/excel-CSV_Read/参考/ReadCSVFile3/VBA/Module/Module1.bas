Attribute VB_Name = "Module1"
'*******************************************************************************
'   CSV形式ファイル読み込み処理(カンマ数不定処理)
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'*******************************************************************************
' [参照設定]
'   ・Microsoft Scripting Runtime
'*******************************************************************************
Option Explicit
Const cnsTITLE = "テキストファイル読み込み処理"
Const cnsFILTER = "全てのファイル (*.*),*.*"
Const cnsStartRow = 2               ' 読み込み開始行
Const cnsStartCol = 1               ' 読み込み開始カラム番号

'*******************************************************************************
' CSV形式テキストファイル(不定カラム)読み込みサンプル
'*******************************************************************************
Sub READ_TextFile()
    Dim xlAPP As Application        ' Applicationオブジェクト
    Dim objFSO As FileSystemObject  ' FSOオブジェクト
    Dim objTS As TextStream         ' TextStreamオブジェクト
    Dim strFILENAME As String       ' OPENするﾌｧｲﾙ名(ﾌﾙﾊﾟｽ)
    Dim IX1 As Long                 ' CSV項目ｶﾗﾑINDEX
    Dim GYO As Long                 ' 収容するｾﾙの行
    Dim lngREC As Long              ' ﾚｺｰﾄﾞ件数ｶｳﾝﾀ
    Dim vntREC As Variant           ' ﾚｺｰﾄﾞ内容(配列)
    
    
    ' Applicationｵﾌﾞｼﾞｪｸﾄ取得
    Set xlAPP = Application
    ' ｢ﾌｧｲﾙを開く｣のﾌｫｰﾑでﾌｧｲﾙ名の指定を受ける
    xlAPP.StatusBar = "読み込むファイル名を指定して下さい。"
    strFILENAME = xlAPP.GetOpenFilename(FileFilter:=cnsFILTER, _
        Title:=cnsTITLE)
    ' ｷｬﾝｾﾙされた場合は以降の処理は行なわない
    If StrConv(strFILENAME, vbUpperCase) = "FALSE" Then Exit Sub
    
    ' 指定ﾌｧｲﾙをOPEN(入力ﾓｰﾄﾞ)
    Set objFSO = New FileSystemObject
    Set objTS = objFSO.OpenTextFile(strFILENAME, ForReading)
    
    GYO = cnsStartRow - 1
    ' ﾌｧｲﾙのEOF(End of File)まで繰り返す
    Do Until objTS.AtEndOfStream
        ' ﾚｺｰﾄﾞ件数ｶｳﾝﾀの加算
        lngREC = lngREC + 1
        xlAPP.StatusBar = "読み込み中です．．．．(" & lngREC & "レコード目)"
        
        ' 行単位にﾚｺｰﾄﾞを読み込む(共通処理)
        vntREC = modGetCSVRec2.FP_GET_CSV_REC2(objTS)
        IX1 = UBound(vntREC) + cnsStartCol
        
        ' 行を加算しﾚｺｰﾄﾞ内容を表示(先頭は2行目)
        GYO = GYO + 1
        Range(Cells(GYO, cnsStartCol), Cells(GYO, IX1)).Value = vntREC   ' 配列渡し
    Loop
    ' 指定ﾌｧｲﾙをCLOSE
    objTS.Close
    xlAPP.StatusBar = False
    ' 処理終了
    Set objTS = Nothing
    Set objFSO = Nothing
    ' 終了の表示
    MsgBox "ファイル読み込みが完了しました。" & vbCr & _
        "レコード件数=" & lngREC & "件", vbInformation, cnsTITLE
    Set xlAPP = Nothing
End Sub

'-----------------------------<< End of Source >>-------------------------------

