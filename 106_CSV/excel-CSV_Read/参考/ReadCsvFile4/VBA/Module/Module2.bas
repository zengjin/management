Attribute VB_Name = "Module2"
'***************************************************************************************************
'   CSV読み込みテスト                                                   Module1(Module)
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'***************************************************************************************************
'変更日付 Rev  変更履歴内容------------------------------------------------------------------------>
'19/11/06(1.00)新規作成
'19/11/09(1.10)カラム数チェック有無の指定を追加
'***************************************************************************************************
Option Explicit
'===================================================================================================
Private Const g_cnsTitle As String = "CSV読み込みテスト"
Private Const g_cnsCheckColCnt As Boolean = True                    ' カラム数チェック有無
Private Const g_cnsColCnt As Long = 10                              ' カラム数
' ↑CSVファイルのカラム数をg_cnsColCntにセットしてから実行して下さい
'   チェック不要の場合はg_cnsCheckColCntをFalseに変更して下さい

'***************************************************************************************************
'* 処理名　：TEST2
'* 機能　　：CSV読み込みテスト
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：(なし)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2019年11月06日
'* 作成者　：井上　治
'* 更新日　：2019年11月09日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Public Sub TEST2()
    '-----------------------------------------------------------------------------------------------
    Dim lngIxR As Long                                              ' テーブルINDEX(Row)
    Dim lngIxC As Long                                              ' テーブルINDEX(Col)
    Dim lngIxRMax As Long                                           ' テーブルINDEX上限(Row)
    Dim strFilename As String                                       ' ファイル名
    Dim strErrMSG As String                                         ' エラーメッセージ
    Dim vntFilename As Variant                                      ' ファイル名(受け取り)
    Dim tblRec As Variant                                           ' JAG配列テーブル
    Dim tblRec2() As Variant                                        ' 二次元配列テーブル
    '-----------------------------------------------------------------------------------------------
    ' ファイル名受け取り
    vntFilename = Application.GetOpenFilename("CSVファイル (*.csv;*.txt),*.csv;*.txt", , g_cnsTitle)
    ' キャンセルは終了
    If VarType(vntFilename) = vbBoolean Then Exit Sub
    strFilename = vntFilename
    '-----------------------------------------------------------------------------------------------
    ' CSV読み込みクラス(UTF-8)
    With New clsReadCsv2
        ' カラム数チェック有無を指定
        .prpCheckColCnt = g_cnsCheckColCnt
        ' カラム数を指定
        .prpColCnt = g_cnsColCnt
        ' CSV読み込み
        If Not .ReadCsv(strFilename, tblRec) Then
            MsgBox .prpErrMSG, vbCritical, g_cnsTitle
            Exit Sub
        End If
    End With
    '===============================================================================================
    ' 以下はサンプルとしての結果検証用の記述です
    ' カラム数チェックを行なわない指定でもg_cnsColCntの指定列までしかシートには書き出されません
    '-----------------------------------------------------------------------------------------------
    lngIxRMax = UBound(tblRec)
    ReDim tblRec2(lngIxRMax, g_cnsColCnt - 1)
    ' JAG配列テーブルを二次元配列テーブルに変換
    Do While lngIxR <= lngIxRMax
        lngIxC = 0
        ' カラム方向ループ
        Do While lngIxC < g_cnsColCnt
            tblRec2(lngIxR, lngIxC) = tblRec(lngIxR)(lngIxC)
            ' 次へ
            lngIxC = lngIxC + 1
        Loop
        ' 次へ
        lngIxR = lngIxR + 1
    Loop
    ReDim tblRec(0)
    '-----------------------------------------------------------------------------------------------
    ' 現在シートに貼り付け
    With ActiveSheet
        .Range(.Cells(1, 1), .Cells(lngIxR, g_cnsColCnt)).Value = tblRec2
    End With
End Sub

'----------------------------------------<< End of Source >>----------------------------------------

