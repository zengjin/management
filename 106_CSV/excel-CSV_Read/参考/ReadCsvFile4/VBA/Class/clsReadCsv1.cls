VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReadCsv1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   CSV読み込みクラス                                               clsReadCsv1(Class)
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'***************************************************************************************************
'   [前提条件]
'   ・カンマ区切りテキストファイルであること
'   ・文字コードはAscii/シフトJIS形式
'   ・改行コードはCR、LF、CRLFに対応
'   ・フィールド数は全行一致していること(チェック有効時)
'   ・フィールド内で改行、あるいはフィールド内にカンマ、ダブルクォーテーションがある場合は
'   　フィールド自体をダブルクォーテーションで囲うこと
'***************************************************************************************************
'   [参照設定]
'   ・Microsoft Scripting Runtime
'***************************************************************************************************
'変更日付 Rev  変更履歴内容------------------------------------------------------------------------>
'19/11/06(1.00)新規作成
'19/11/09(1.10)LFコードのみの改行に対応、カラム数チェック有無の指定を追加
'***************************************************************************************************
Option Explicit
'===================================================================================================
Private Const g_cnsCom As String = ","
Private Const g_cnsDq As String = """"
'---------------------------------------------------------------------------------------------------
Private g_blnCheckColCnt As Boolean                                 ' カラム数チェック有無
Private g_lngColCnt As Long                                         ' チェック時のカラム数
Private g_strErrMSG As String                                       ' エラーメッセージ

'***************************************************************************************************
'   ■■■ 公開プロシージャ ■■■
'***************************************************************************************************
'* 処理名　：ReadCsv
'* 機能　　：CSV読み込み
'---------------------------------------------------------------------------------------------------
'* 返り値　：処理成否(Boolean)
'* 引数　　：Arg1 = ファイル名(String)                      ※フルパス
'* 　　　　　Arg2 = CSV内容JAG配列テーブル(Variant)         ※Ref参照
'---------------------------------------------------------------------------------------------------
'* 作成日　：2019年11月06日
'* 作成者　：井上　治
'* 更新日　：2019年11月09日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Friend Function ReadCsv(ByVal strFilename As String, _
                        ByRef tblRec As Variant) As Boolean
    '-----------------------------------------------------------------------------------------------
    Dim strAllRec As String                                         ' レコード全体
    ReadCsv = False
    g_strErrMSG = ""
    ReDim tblRec(0)
    ' カラム数未設定は終了(チェック有効時)
    If g_blnCheckColCnt And g_lngColCnt = 0 Then
        g_strErrMSG = "カラム数が設定されていません。"
        Exit Function
    End If
    ' CSVファイル読み込み
    If Not FP_ReadFile(strFilename, strAllRec) Then Exit Function
    ' CSVファイル内容をJAG配列に変換
    If Not FP_GetJagTable(strAllRec, tblRec) Then Exit Function
    ' 終了
    ReadCsv = True
End Function

'***************************************************************************************************
'   ■■■ サブ処理 ■■■
'***************************************************************************************************
'* 処理名　：FP_ReadFile
'* 機能　　：CSVファイル読み込み
'---------------------------------------------------------------------------------------------------
'* 返り値　：処理成否(Boolean)
'* 引数　　：Arg1 = ファイル名(String)                      ※フルパス
'* 　　　　　Arg2 = レコード全体(String)                    ※Ref参照
'---------------------------------------------------------------------------------------------------
'* 作成日　：2019年11月06日
'* 作成者　：井上　治
'* 更新日　：2019年11月06日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Function FP_ReadFile(ByVal strFilename As String, _
                             ByRef strAllRec As String) As Boolean
    '-----------------------------------------------------------------------------------------------
    Dim objFso As FileSystemObject                                  ' FileSystemObject
    Dim objTs As TextStream                                         ' TextStream
    Dim blnOpen As Boolean                                          ' OPEN判定
    FP_ReadFile = False
    On Error GoTo ReadFile_ERROR
    Set objFso = New FileSystemObject
    ' CSVファイルOPEN
    Set objTs = objFso.OpenTextFile(strFilename, ForReading, False)
    blnOpen = True
    ' ファイル全量を読み込み
    strAllRec = objTs.ReadAll
    FP_ReadFile = True
    GoTo ReadFile_EXIT

'===================================================================================================
ReadFile_ERROR:
    g_strErrMSG = Err.Description

'===================================================================================================
ReadFile_EXIT:
    ' CSVファイルCLOSE
    If blnOpen Then objTs.Close
    Set objTs = Nothing
    Set objFso = Nothing
    On Error GoTo 0
End Function

'***************************************************************************************************
'* 処理名　：FP_GetJagTable
'* 機能　　：CSVファイル内容をJAG配列に変換
'---------------------------------------------------------------------------------------------------
'* 返り値　：処理成否(Boolean)
'* 引数　　：Arg1 = レコード全体(String)
'* 　　　　　Arg2 = CSV内容JAG配列テーブル(Variant)         ※Ref参照
'---------------------------------------------------------------------------------------------------
'* 作成日　：2019年11月06日
'* 作成者　：井上　治
'* 更新日　：2019年11月09日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Function FP_GetJagTable(ByRef strAllRec As String, _
                                ByRef tblRec As Variant) As Boolean
    '-----------------------------------------------------------------------------------------------
    ' 全体で使用する変数
    Dim lngIxRec As Long                                            ' レコードINDEX
    Dim lngPos As Long                                              ' 現在文字位置
    Dim lngPosEnd As Long                                           ' 最終文字位置
    Dim lngCntRec As Long                                           ' レコードカウンタ
    ' レコード単位で使用する変数
    Dim lngIxFld As Long                                            ' フィールドINDEX
    Dim lngCntFld As Long                                           ' フィールドカウンタ
    Dim blnCrLf As Boolean                                          ' 改行判定
    Dim blnOkRec As Boolean                                         ' 有効レコード判定
    Dim tblFld() As String                                          ' フィールドテーブル
    ' フィールド単位で使用する変数
    Dim lngPosS As Long                                             ' 項目開始位置
    Dim lngPosE As Long                                             ' 項目終了位置
    Dim lngPosS2 As Long                                            ' 項目開始位置(DQ調整後)
    Dim lngPosE2 As Long                                            ' 項目終了位置(DQ調整後)
    Dim blnInText As Boolean                                        ' 文字列項目内判定
    Dim blnDq As Boolean                                            ' ダブルクォーテーション判定
    Dim blnDq2 As Boolean                                           ' 文字列内DQ発生判定
    Dim strText As String                                           ' 項目テキストWORK
    '-----------------------------------------------------------------------------------------------
    FP_GetJagTable = False
    lngPos = 1
    lngIxRec = -1
    lngPosEnd = Len(strAllRec)
    ' 内容無しはエラー
    If lngPosEnd = 0 Then
        g_strErrMSG = "ファイルに内容が書き込まれていません。"
        Exit Function
    End If
    '-----------------------------------------------------------------------------------------------
    ' 最終文字位置まで繰り返す
    Do While lngPos <= lngPosEnd
        '=============================================================
        ' レコード単位[前]処理
        '-------------------------------------------------------------
        blnCrLf = False
        blnOkRec = False
        lngIxFld = -1
        lngCntFld = 0
        ReDim tblFld(0)
        '=============================================================
        ' レコード単位[主]処理(レコード内項目巡回)
        '-------------------------------------------------------------
        Do While lngPos <= lngPosEnd
            '-----------------------------------------------
            ' フィールド単位[前]処理
            '-----------------------------------------------
            blnInText = False
            blnDq = False
            blnDq2 = False
            ' 項目先頭位置
            lngPosS = lngPos
            '-----------------------------------------------
            ' フィールド単位[主]処理
            '-----------------------------------------------
            Do While lngPos <= lngPosEnd
                ' 現在文字を判定
                Select Case Mid(strAllRec, lngPos, 1)
                    Case vbCr                               ' CRコード
                        ' 文字列項目内でなければレコード終了
                        If Not blnInText Then
                            blnCrLf = True
                            Exit Do
                        End If
                    Case vbLf                               ' LFコード
                        ' 文字列項目内でなければレコード終了
                        If Not blnInText Then
                            blnCrLf = True
                            Exit Do
                        End If
                    Case g_cnsDq                            ' ダブルクォーテーション
                        ' 文字列内DQ発生判定か
                        If blnInText Then
                            blnDq2 = True
                        End If
                        blnInText = Not blnInText
                        blnDq = True
                    Case g_cnsCom                           ' カンマ
                        ' 文字列項目内でなければ項目終了
                        If Not blnInText Then Exit Do
                End Select
                ' 次の文字へ
                lngPos = lngPos + 1
            Loop
            '-----------------------------------------------
            ' フィールド単位[後]処理
            '-----------------------------------------------
            ' 項目終了位置(実際は次のカンマ位置)
            lngPosE = lngPos
            lngPosS2 = lngPosS
            lngPosE2 = lngPosE
            ' ダブルクォーテーションで囲われているか
            If blnDq Then
                ' 両端文字(ダブルクォーテーション)を除外
                lngPosS2 = lngPosS + 1
                lngPosE2 = lngPosE - 1
            End If
            lngCntFld = lngCntFld + 1
            ' フィールドテーブル要素を追加
            lngIxFld = lngIxFld + 1
            ReDim Preserve tblFld(lngIxFld)
            ' 有効なフィールドか(ブランクでないか)
            If lngPosE2 > lngPosS2 Then
                blnOkRec = True
                ' 項目テキストの取り出し
                strText = Mid(strAllRec, lngPosS2, lngPosE2 - lngPosS2)
                ' 文字列内ダブルクォーテーション発生判定
                If blnDq2 Then
                    ' ダブルクォーテーション連記を解除(完璧な方法ではない)
                    strText = Replace(strText, """""", """")
                End If
                ' フィールドテーブルに格納
                tblFld(lngIxFld) = strText
            End If
            ' カンマ､改行を除外するため+1
            lngPos = lngPosE + 1
            ' レコード終了判定
            If blnCrLf Then Exit Do
        Loop
        '=============================================================
        ' レコード単位[後]処理
        '-------------------------------------------------------------
        ' Cr後のLfは除外
        If Mid(strAllRec, lngPos, 1) = vbLf Then
            lngPos = lngPos + 1
        End If
        ' 有効レコードならJAG配列テーブルに格納
        If blnOkRec Then
            lngCntRec = lngCntRec + 1
            ' 項目数判定
            If g_blnCheckColCnt And lngCntFld <> g_lngColCnt Then
                Call GP_AppendMessage(CStr(lngCntRec) & "REC目、項目数不正(" & lngCntFld & ")")
            End If
            lngIxRec = lngIxRec + 1
            ReDim Preserve tblRec(lngIxRec)
            tblRec(lngIxRec) = tblFld
        End If
    Loop
    ' 有効内容無し
    If lngCntRec = 0 Then
        g_strErrMSG = "有効な内容がありません。"
    End If
    ' エラーがなければOK
    FP_GetJagTable = g_strErrMSG = ""
End Function

'***************************************************************************************************
'* 処理名　：GP_AppendMessage
'* 機能　　：メッセージ累積
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 今回メッセージ(String)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2019年11月06日
'* 作成者　：井上　治
'* 更新日　：2019年11月06日
'* 更新者　：井上　治
'* 機能説明：改行を挟んでメッセージを累積する
'* 注意事項：
'***************************************************************************************************
Private Sub GP_AppendMessage(ByVal strAddMSG As String)
    '-----------------------------------------------------------------------------------------------
    If g_strErrMSG <> "" Then g_strErrMSG = g_strErrMSG & vbCrLf
    g_strErrMSG = g_strErrMSG & strAddMSG
End Sub

'***************************************************************************************************
'   ■■■ プロパティ ■■■
'***************************************************************************************************
'   カラム数チェック有無(Boolean)
'---------------------------------------------------------------------------------------------------
Friend Property Let prpCheckColCnt(ByVal blnValue As Boolean)
    g_blnCheckColCnt = blnValue
End Property

'===================================================================================================
'   チェック時のカラム数(Long)
'---------------------------------------------------------------------------------------------------
Friend Property Let prpColCnt(ByVal lngValue As Long)
    g_lngColCnt = lngValue
End Property

'===================================================================================================
'   エラーメッセージ(String)
'---------------------------------------------------------------------------------------------------
Friend Property Get prpErrMSG() As String
    prpErrMSG = g_strErrMSG
End Property

'----------------------------------------<< End of Source >>----------------------------------------
