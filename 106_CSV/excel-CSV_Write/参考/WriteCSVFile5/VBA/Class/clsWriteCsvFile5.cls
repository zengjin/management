VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWriteCsvFile5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   CSV形式テキストファイル書き出しクラス                       clsWriteCsvFile5(Class)
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'***************************************************************************************************
'   [参照設定]
'   ・Microsoft Scription Runtime
'***************************************************************************************************
'変更日付 Rev  変更履歴内容------------------------------------------------------------------------>
'19/12/22(1.00)新規作成
'***************************************************************************************************
Option Explicit
'===================================================================================================
' 共通定数
Private Const g_cnsComm As String = ","
Private Const g_cnsDQ As String = """"
'---------------------------------------------------------------------------------------------------
' 呼び元から引き渡される変数
Private g_lngCntColumns As Long                                 ' 項目カラム数
Private g_lngFirstRow As Long                                   ' 読出し先頭行番号
Private g_lngFirstCol As Long                                   ' 読出し先頭カラム番号
Private g_lngCntMidashiRow As Long                              ' 見出し行数
Private g_lngColCntMaxRows As Long                              ' 最終行判定列番号
Private g_blnUseMidashi As Boolean                              ' 見出し出力有無
Private g_blnUseUnicode As Boolean                              ' Unicode指定
Private g_strSepChar As String                                  ' 項目境界文字
Private g_strFilename As String                                 ' 出力ファイル名
Private g_strErrMSG As String                                   ' エラーメッセージ

'***************************************************************************************************
'   ■■■ 共通サブ処理 ■■■
'***************************************************************************************************
'* 処理名　：Class_Initialize
'* 機能　　：クラス初期処理
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：(なし)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2019年12月22日
'* 作成者　：井上　治
'* 更新日　：2019年12月22日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub Class_Initialize()
    '-----------------------------------------------------------------------------------------------
    g_lngCntColumns = 0
    g_lngFirstRow = 1
    g_lngCntMidashiRow = 1
    g_lngColCntMaxRows = 1
    g_blnUseMidashi = True
    g_blnUseUnicode = False
    g_strSepChar = g_cnsComm
    g_strErrMSG = ""
End Sub

'***************************************************************************************************
'   ■■■ ワークシート側からの呼び出し処理 ■■■
'***************************************************************************************************
'* 処理名　：WriteCsvFile5
'* 機能　　：CSV形式テキストファイル出力
'---------------------------------------------------------------------------------------------------
'* 返り値　：処理成否(Boolean)
'* 引数　　：Arg1 = 処理対象シート(Object)
'* 　　　　　Arg2 = 追記出力指定(Boolean)               ※Option
'---------------------------------------------------------------------------------------------------
'* 作成日　：2019年12月22日
'* 作成者　：井上　治
'* 更新日　：2019年12月22日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：引数パラメータの条理チェックは行なっていない(正しい前提)
'***************************************************************************************************
Friend Function WriteCsvFile5(ByVal objSh As Worksheet, _
                              Optional ByVal blnAppend As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------
    Dim objFso As FileSystemObject                                  ' FileSystemObject
    Dim objTs As TextStream                                         ' TextStream
    Dim lngRow As Long                                              ' 行INDEX
    Dim lngRowMax As Long                                           ' 最終行INDEX
    Dim blnOpen As Boolean                                          ' ファイルOPEN判定
    WriteCsvFile5 = False
    On Error GoTo WriteCsvFile5_ERROR
    Set objFso = New FileSystemObject
    ' 出力ファイルOPEN
    Set objTs = objFso.CreateTextFile(g_strFilename, True, g_blnUseUnicode)
    blnOpen = True
    With objSh
        ' 保護解除(PW不対応)
        If .ProtectContents Then .Unprotect
        ' フィルタ解除
        If .FilterMode Then .ShowAllData
        ' 先頭行設定
        lngRow = g_lngFirstRow
        ' 項目カラム数がゼロの場合は先頭行で判定
        If g_lngCntColumns = 0 Then
            g_lngCntColumns = .Cells(lngRow, .Columns.Count).End(xlToLeft).Column
        End If
        ' 見出し出力無しの時は見出し行数を加算(見出しSKIP)
        If Not g_blnUseMidashi Then
            lngRow = lngRow + g_lngCntMidashiRow
        End If
        ' 最終行判定
        lngRowMax = .Cells(.Rows.Count, g_lngColCntMaxRows).End(xlUp).Row
        ' 最終行まで繰り返す
        Do While lngRow <= lngRowMax
            ' 1行分のCSV編集・出力
            objTs.WriteLine FP_EditCsvRec(objSh, lngRow)
            ' 次の行へ
            lngRow = lngRow + 1
        Loop
    End With
    WriteCsvFile5 = g_strErrMSG = ""
    GoTo WriteCsvFile5_EXIT
    
'===================================================================================================
' エラートラップ
WriteCsvFile5_ERROR:
    Call GP_AppendMessage2(Err.Description)

'===================================================================================================
' 終了
WriteCsvFile5_EXIT:
    ' 出力ファイルCLOSE
    If blnOpen Then objTs.Close
    Set objFso = Nothing
End Function

'***************************************************************************************************
'   ■■■ サブ処理(Private) ■■■
'***************************************************************************************************
'* 処理名　：FP_EditCsvRec
'* 機能　　：1行分のCSV編集
'---------------------------------------------------------------------------------------------------
'* 返り値　：1行分の編集結果(String)
'* 引数　　：Arg1 = 処理対象シート(Object)
'* 　　　　　Arg2 = 現在行INDEX(Long)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2019年12月22日
'* 作成者　：井上　治
'* 更新日　：2019年12月22日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Function FP_EditCsvRec(ByVal objSh As Worksheet, ByVal lngRow As Long) As String
    '-----------------------------------------------------------------------------------------------
    Dim lngCol As Long                                              ' カラムINDEX
    Dim strRec As String                                            ' CSVレコード
    With objSh
        ' 先頭カラム処理
        strRec = FP_EditCsvFld(.Cells(lngRow, g_lngFirstCol).Value)
        lngCol = g_lngFirstCol + 1
        ' 次カラム以降
        Do While lngCol <= g_lngCntColumns
            ' 1セル分のCSV編集
            strRec = strRec & g_strSepChar & FP_EditCsvFld(.Cells(lngRow, lngCol).Value)
            ' 次の列へ
            lngCol = lngCol + 1
        Loop
    End With
    FP_EditCsvRec = strRec
End Function

'***************************************************************************************************
'* 処理名　：FP_EditCsvFld
'* 機能　　：1セル分のCSV編集
'---------------------------------------------------------------------------------------------------
'* 返り値　：1セル分の編集結果(String)
'* 引数　　：Arg1 = 現在セル値(Variant)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2019年12月22日
'* 作成者　：井上　治
'* 更新日　：2019年12月22日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Function FP_EditCsvFld(ByVal vntValue As Variant) As String
    '-----------------------------------------------------------------------------------------------
    Dim lngPos As Long                                              ' 文字位置
    Dim blnNeedDQ As Boolean                                        ' ダブルクォーテーション囲い要否
    Dim blnNewLine As Boolean                                       ' セル内改行有無
    Dim strFld As String                                            ' セル入力項目値
    Dim strFldO As String                                           ' CSV出力項目値
    Dim strChar As String * 1                                       ' 1文字
    ' タイプ判定
    Select Case True
        Case IsDate(vntValue)
            FP_EditCsvFld = CStr(vntValue)
        Case IsNumeric(vntValue)
            FP_EditCsvFld = CStr(vntValue)
        Case Else
            strFld = vntValue
            ' ブランクセルは終了
            If Trim(strFld) = "" Then
                FP_EditCsvFld = ""
                Exit Function
            End If
            lngPos = 1
            ' 1文字ずつ判定
            Do While lngPos <= Len(strFld)
                strChar = Mid(strFld, lngPos, 1)
                ' 文字判定
                Select Case strChar
                    Case g_cnsDQ                            ' ダブルクォーテーション
                        blnNeedDQ = True
                        ' ダブルクォーテーションは連記
                        strFldO = strFldO & strChar
                    Case g_cnsComm                          ' カンマ
                        blnNeedDQ = True
                    Case vbCr, vbLf                         ' 改行コード
                        blnNewLine = True
                End Select
                strFldO = strFldO & strChar
                ' 次の文字へ
                lngPos = lngPos + 1
            Loop
            ' セル内改行有無
            If blnNewLine Then
                blnNeedDQ = True
                ' CrLfを一旦Lfのみに変換
                strFldO = Replace(strFldO, vbCrLf, vbLf)
                ' Crのみを一旦Lfのみに変換
                strFldO = Replace(strFldO, vbCr, vbLf)
                ' LfをCrLfに変換
                strFldO = Replace(strFldO, vbLf, vbCrLf)
            End If
            ' ダブルクォーテーション囲い要否
            If blnNeedDQ Then
                FP_EditCsvFld = g_cnsDQ & strFldO & g_cnsDQ
            Else
                FP_EditCsvFld = strFldO
            End If
    End Select
End Function

'***************************************************************************************************
'* 処理名　：GP_AppendMessage2
'* 機能　　：エラーメッセージ累積
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 今回メッセージ(String)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2019年12月22日
'* 作成者　：井上　治
'* 更新日　：2019年12月22日
'* 更新者　：井上　治
'* 機能説明：改行を付加して追加
'* 注意事項：
'***************************************************************************************************
Private Sub GP_AppendMessage2(ByVal strAddMSG As String)
    '-----------------------------------------------------------------------------------------------
    If g_strErrMSG <> "" Then g_strErrMSG = g_strErrMSG & vbCrLf
    g_strErrMSG = g_strErrMSG & strAddMSG
End Sub

'***************************************************************************************************
'   ■■■ プロパティ ■■■
'***************************************************************************************************
' 項目カラム数(Long)
'---------------------------------------------------------------------------------------------------
Friend Property Let prpCntColumns(ByVal lngValue As Long)
    g_lngCntColumns = lngValue
End Property

'===================================================================================================
' 読出し先頭行番号(Long)
'---------------------------------------------------------------------------------------------------
Friend Property Let prpFirstRow(ByVal lngValue As Long)
    g_lngFirstRow = lngValue
End Property

'===================================================================================================
' 読出し先頭カラム番号(Long)
'---------------------------------------------------------------------------------------------------
Friend Property Let prpFirstCol(ByVal lngValue As Long)
    g_lngFirstCol = lngValue
End Property

'===================================================================================================
' 見出し行数(Long)
'---------------------------------------------------------------------------------------------------
Friend Property Let prpCntMidashiRow(ByVal lngValue As Long)
    g_lngCntMidashiRow = lngValue
End Property

'===================================================================================================
' 最終行判定列番号(Long)
'---------------------------------------------------------------------------------------------------
Friend Property Let prpColCntMaxRows(ByVal lngValue As Long)
    g_lngColCntMaxRows = lngValue
End Property

'===================================================================================================
' 見出し出力有無(Boolean)
'---------------------------------------------------------------------------------------------------
Friend Property Let prpUseMidashi(ByVal blnValue As Boolean)
    g_blnUseMidashi = blnValue
End Property

'===================================================================================================
' Unicode指定(Boolean)
'---------------------------------------------------------------------------------------------------
Friend Property Let prpUseUnicode(ByVal blnValue As Boolean)
    g_blnUseUnicode = blnValue
End Property

'===================================================================================================
' 項目境界文字(String)
'---------------------------------------------------------------------------------------------------
Friend Property Let prpSepChar(ByVal strValue As String)
    g_strSepChar = strValue
End Property

'===================================================================================================
' 出力ファイル名(String)
'---------------------------------------------------------------------------------------------------
Friend Property Let prpFilename(ByVal strValue As String)
    g_strFilename = strValue
End Property

'===================================================================================================
' エラーメッセージ(String)
'---------------------------------------------------------------------------------------------------
Friend Property Get prpErrMSG() As String
    prpErrMSG = g_strErrMSG
End Property

'----------------------------------------<< End of Source >>----------------------------------------
