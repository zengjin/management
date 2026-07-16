Attribute VB_Name = "modWRITE_CSVFile4"
'*******************************************************************************
'   CSV形式テキストファイル書き出すサンプル(FSO)    ※汎用型
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'*******************************************************************************
' [参照設定]
'   ・Microsoft Scripting Runtime
'*******************************************************************************
Option Explicit
'-------------------------------------------------------------------------------
' ※この下の各項目を設定することで、ソースコードを改変しなくても
' 　いろいろなケースのCSV形式ファイル出力に対応できると思います。
'-------------------------------------------------------------------------------
'            ↓↓↓ CSV形式ファイルの出力方法の指定 ↓↓↓
'-------------------------------------------------------------------------------
' ■出力ファイルに関する指定
Private Const g_cnsDEFFILE = "SAMPLE.csv"   ' ファイル名(デフォルト)
Private Const g_cnsGETFILE = 1              ' ファイル名可変指定(0=固定、1=可変)
' ※「0=固定」の場合はこのブックのフォルダに出力
' ※「1=可変」の場合は「名前を付けて保存」のダイアログが表示される

' ■シート上の出力範囲の指定
Private Const g_cnsSTRGYO = 2               ' 出力開始行
Private Const g_cnsCHKCOL = 1               ' 最終行判定カラム番号
Private Const g_cnsSTRCOL = 1               ' 出力開始カラム番号
Private Const g_cnsENDCOL = 5               ' 出力終了カラム番号
Private Const g_cnsMIDASHI = 0              ' 先頭行見出し指定
' ※(先頭行見出し指定の値の説明)
'  0=見出しなし(先頭行からデータとして編集)
'  1=先頭行のみセパレータのみ(下記「カラム単位の編集」は無視)
'  2=先頭行は全て「"」囲い(下記「カラム単位の編集」は無視)
'  3=先頭行は全て「'」囲い(下記「カラム単位の編集」は無視)

' ■CSV形式ファイルのセパレータの指定(通常はカンマ)
Private Const g_cnsSEPCHR = ","             ' セパレータ文字
'Private Const g_cnsSEPCHR = vbTab           ' Tab区切りはこちらを有効にする

' ■カラム単位の編集方法の指定
'                カラム位置⇒ ....*....1....*....2....*....3....*....4....*....5
Private Const g_cnsCOLEDIT = "29910000000000000000000000000000000000000000000000"
' ※(カラム単位の編集方法の値の説明)
'  0=セパレータのみ
'  1=数値(ブランクはゼロ出力)
'  2=文字列(「"」囲い：無条件)
'  3=文字列(「"」囲い：ブランクを除く)
'  4=文字列(「'」囲い：無条件)
'  5=文字列(「'」囲い：ブランクを除く)
'  6=日付(「#」囲い：ブランクを除く)
'  8=自動判定(数値,ブランク以外は「"」囲い)
'  9=自動判定(数値,ブランク以外は、日付は「#」囲い、その他は「"」囲い)
'-------------------------------------------------------------------------------
'            ↑↑↑ CSV形式ファイルの出力方法の指定 ↑↑↑
'-------------------------------------------------------------------------------
Private Const g_cnsDQ = """"
Private Const g_cnsSQ = "'"
Private Const g_cnsSH = "#"
Private Const g_cnsYEN = "\"

'*******************************************************************************
' CSV形式テキストファイル書き出すサンプル④(FSO)  ※汎用型
'-------------------------------------------------------------------------------
'   ※シートの指定がない場合はアクティブなシートが対象になります。
'   ※データの条理的なチェックは行ないません。
'   ※処理内では実行時エラーの対応はしていません。
'*******************************************************************************
Sub WRITE_CSVFile4(Optional SH As Worksheet)
    Dim FSO As New FileSystemObject ' FileSystemObject
    Dim TS As TextStream            ' TextStream
    Dim strFILENAME As String       ' 出力ファイル名
    Dim vntFILENAME As Variant      ' 出力ファイル名指定
    Dim GYO As Long                 ' 収容するセルの行
    Dim GYOMAX As Long              ' データが収容された最終行
    Dim strREC As String            ' 出力レコード(見出し編集用)
    Dim COL As Long                 ' カラム(見出し編集用)
    Dim strDC As String             ' 囲い文字(見出し編集用)

    ' 出力ファイル名の指定
    If g_cnsGETFILE = 1 Then
        ' ファイル名をダイアログで指定
        vntFILENAME = Application.GetSaveAsFilename(g_cnsDEFFILE)
        If VarType(vntFILENAME) = vbBoolean Then Exit Sub
        strFILENAME = vntFILENAME
    Else
        ' ファイル名は初期値固定(自フォルダ)
        strFILENAME = ThisWorkbook.Path & g_cnsYEN & g_cnsDEFFILE
    End If
    ' ワークシート判定(Nothingの場合はActiveSheet)
    If SH Is Nothing Then Set SH = ActiveSheet
    ' 最終行の取得
    GYOMAX = SH.Cells(65536, g_cnsCHKCOL).End(xlUp).Row
    ' 指定ファイルをOPEN(出力モード、既存は無条件上書き)
    Set TS = FSO.CreateTextFile(Filename:=strFILENAME, Overwrite:=True)
    ' 指定行目から開始
    GYO = g_cnsSTRGYO
    ' 見出し処理
    If g_cnsMIDASHI <> 0 Then
        ' 見出し指定がある場合は見出しのみ別指定方法で編集
        Select Case g_cnsMIDASHI
            Case 1: strDC = ""
            Case 2: strDC = g_cnsDQ
            Case 3: strDC = g_cnsSQ
        End Select
        strREC = strDC & SH.Cells(GYO, g_cnsSTRCOL).Value & strDC
        COL = g_cnsSTRCOL + 1
        Do While COL <= g_cnsENDCOL
            strREC = strREC & g_cnsSEPCHR & _
                strDC & Trim(SH.Cells(GYO, COL).Value) & strDC
            COL = COL + 1
        Loop
        ' レコードを出力
        TS.WriteLine strREC
        GYO = GYO + 1
    End If
    ' 最終行まで繰り返す
    Do Until GYO > GYOMAX
        ' レコードを出力(REC編集処理より受け取る)
        TS.WriteLine FP_EDIT_CSVREC(SH, GYO)
        ' 行を加算
        GYO = GYO + 1
    Loop
    ' 指定ファイルをCLOSE
    TS.Close
    Set TS = Nothing
    Set FSO = Nothing
End Sub

'*******************************************************************************
' CSV形式テキストの1レコードの編集処理(引数はシート、シート上の行位置)
'*******************************************************************************
Private Function FP_EDIT_CSVREC(SH As Worksheet, GYO As Long) As String
    Dim strREC As String
    Dim COL As Long
    
    ' 先頭カラムの編集
    COL = g_cnsSTRCOL
    strREC = FP_EDIT_COLUMN(SH, GYO, COL)
    ' ２番目以降のカラムの編集
    Do While COL <= g_cnsENDCOL
        strREC = strREC & g_cnsSEPCHR & FP_EDIT_COLUMN(SH, GYO, COL)
    Loop
    ' 編集したレコード内容を戻り値にセット
    FP_EDIT_CSVREC = strREC
End Function

'*******************************************************************************
' １カラム分の編集処理(引数はシート、行位置、カラム位置)
'*******************************************************************************
Private Function FP_EDIT_COLUMN(SH As Worksheet, GYO As Long, COL As Long) As String
    Dim strTEXT As String

    strTEXT = Trim(SH.Cells(GYO, COL).Value)
    Select Case Mid(g_cnsCOLEDIT, COL, 1)
        Case "0"                ' セパレータのみ
            FP_EDIT_COLUMN = strTEXT
        Case "1"                ' 数値(ブランクはゼロ出力)
            If strTEXT = "" Then
                FP_EDIT_COLUMN = "0"
            Else
                FP_EDIT_COLUMN = strTEXT
            End If
        Case "2"                ' 文字列(「"」囲い：無条件)
            FP_EDIT_COLUMN = g_cnsDQ & strTEXT & g_cnsDQ
        Case "3"                ' 文字列(「"」囲い：ブランクを除く)
            If strTEXT = "" Then
                FP_EDIT_COLUMN = strTEXT
            Else
                FP_EDIT_COLUMN = g_cnsDQ & strTEXT & g_cnsDQ
            End If
        Case "4"                ' 文字列(「'」囲い：無条件)
            FP_EDIT_COLUMN = g_cnsSQ & strTEXT & g_cnsSQ
        Case "5"                ' 文字列(「'」囲い：ブランクを除く)
            If strTEXT = "" Then
                FP_EDIT_COLUMN = strTEXT
            Else
                FP_EDIT_COLUMN = g_cnsSQ & strTEXT & g_cnsSQ
            End If
        Case "6"                ' 日付(「#」囲い：ブランクを除く)
            If strTEXT = "" Then
                FP_EDIT_COLUMN = strTEXT
            Else
                FP_EDIT_COLUMN = g_cnsSH & strTEXT & g_cnsSH
            End If
        Case "8"
            ' 自動判定(数値,ブランク以外は「"」囲い)
            If strTEXT = "" Then
                FP_EDIT_COLUMN = strTEXT
            ElseIf IsNumeric(strTEXT) = True Then
                FP_EDIT_COLUMN = CStr(CDbl(strTEXT))            ' 数値
            Else
                FP_EDIT_COLUMN = g_cnsDQ & strTEXT & g_cnsDQ    ' その他(文字列)
            End If
        Case Else
            ' 自動判定(数値,ブランク以外は、日付は「#」、その他は「"」囲い)
            If strTEXT = "" Then
                FP_EDIT_COLUMN = strTEXT
            ElseIf IsDate(strTEXT) = True Then
                FP_EDIT_COLUMN = g_cnsSH & strTEXT & g_cnsSH    ' 日付
            ElseIf IsNumeric(strTEXT) = True Then
                FP_EDIT_COLUMN = CStr(CDbl(strTEXT))            ' 数値
            Else
                FP_EDIT_COLUMN = g_cnsDQ & strTEXT & g_cnsDQ    ' その他(文字列)
            End If
    End Select
    ' カラムを加算
    COL = COL + 1
End Function

'*******************************************************************************
' テスト用起動プロシージャ
'*******************************************************************************
Sub WRITE_CSV_TEST()
    If g_cnsGETFILE <> 1 Then
        If MsgBox("CSV形式ファイルへの出力を行ないます。" & vbCr & _
            "よろしいですね。", vbInformation + vbYesNo) <> vbYes Then Exit Sub
    End If
    Call WRITE_CSVFile4
End Sub

'-----------------------------<< End of Source >>-------------------------------
