Attribute VB_Name = "Module1"
'***************************************************************************************************
'   CSV形式テキストファイル書き出すサンプル                         Module1(Module)
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'***************************************************************************************************
'   [参照設定]
'   ・Microsoft Scription Runtime
'   ・Windows Script Host Object Model
'***************************************************************************************************
'変更日付 Rev  変更履歴内容------------------------------------------------------------------------>
'19/12/22(1.00)新規作成
'***************************************************************************************************
Option Explicit
'===================================================================================================
Public Const g_cnsTitle As String = "CSV形式テキストファイル書き出すサンプル"
'---------------------------------------------------------------------------------------------------
' CSV形式指定項目
Private Const g_cnsCntColumns As Long = 0                           ' 列数(0=先頭行で自動判定)
Private Const g_cnsFirstRow As Long = 1                             ' 読出し先頭行番号(0不可,見出し含む)
Private Const g_cnsFirstCol As Long = 1                             ' 読出し先頭カラム番号(0不可)
Private Const g_cnsCntMidashiRow As Long = 1                        ' 見出し行数
Private Const g_cnsColCntMaxRows As Long = 1                        ' 最終行判定列番号(0不可)
Private Const g_cnsUseMidashi As Boolean = True                     ' 見出し出力有無(行数0時はFalse)
Private Const g_cnsUseUnicode As Boolean = False                    ' Unicode指定

'***************************************************************************************************
'   ■■■ ユーザー起動処理 ■■■
'***************************************************************************************************
'* 処理名　：WRITE_CSV_TEST
'* 機能　　：CSV形式テキストファイル書き出すサンプル
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
Public Sub WRITE_CSV_TEST()
    '-----------------------------------------------------------------------------------------------
    Dim objShell As WshShell                                        ' WshShell
    Dim objClass As clsWriteCsvFile5                                ' CSV出力クラス
    Dim strCurrPathSV As String                                     ' カレントフォルダ退避
    Dim strFilename As String                                       ' ファイル名
    Dim vntFilename As Variant                                      ' ファイル名(受け取り)
    ' アクティブシートチェック
    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "CSV出力させるワークシートをアクティブにした状態で起動して下さい。", vbExclamation, g_cnsTitle
        Exit Sub
    End If
    '-----------------------------------------------------------------------------------------------
    Set objShell = New WshShell
    strCurrPathSV = objShell.CurrentDirectory
    objShell.CurrentDirectory = ThisWorkbook.Path
    ' ファイル名受け取り
    vntFilename = Application.GetSaveAsFilename("SAMPLE.csv", _
                                                "CSV形式ファイル (*.csv),*.csv", , _
                                                "出力ファイル名の指定")
    objShell.CurrentDirectory = strCurrPathSV
    ' キャンセルは終了
    If VarType(vntFilename) = vbBoolean Then Exit Sub
    strFilename = vntFilename
    '-----------------------------------------------------------------------------------------------
    ' CSV形式テキストファイル書き出しクラス
    Set objClass = New clsWriteCsvFile5
    With objClass
        .prpCntColumns = g_cnsCntColumns                    ' 項目カラム数
        .prpFirstRow = g_cnsFirstRow                        ' 読出し先頭行番号
        .prpFirstCol = g_cnsFirstCol                        ' 読出し先頭カラム番号
        .prpCntMidashiRow = g_cnsCntMidashiRow              ' 見出し行数
        .prpColCntMaxRows = g_cnsColCntMaxRows              ' 最終行判定列番号
        .prpUseMidashi = g_cnsUseMidashi                    ' 見出し出力有無
        .prpUseUnicode = g_cnsUseUnicode                    ' Unicode指定
        .prpFilename = strFilename                          ' 出力ファイル名
        ' 実際のCSV出力処理
        If Not .WriteCsvFile5(ActiveSheet) Then
            ' エラー表示
            MsgBox .prpErrMSG, vbCritical, g_cnsTitle
        Else
            ' 終了表示
            MsgBox "処理完了しました。", vbInformation, g_cnsTitle
        End If
    End With
End Sub

'----------------------------------------<< End of Source >>----------------------------------------
