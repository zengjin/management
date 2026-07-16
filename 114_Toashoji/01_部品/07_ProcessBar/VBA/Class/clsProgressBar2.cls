VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProgressBar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   プログレスバーをコントロールするクラス                      clsProgressBar2(Class)
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'***************************************************************************************************
'変更日付 Rev  変更履歴内容------------------------------------------------------------------------>
'03/08/19(1.00)新規作成
'19/12/29(1.11)処理記述整理(標準化準拠)
'***************************************************************************************************
Option Explicit
'===================================================================================================
' クラス側で保持する変数(隠蔽)
Private g_lngWidth As Long                                          ' Width
Private g_lngNow As Long                                            ' 現在件数
Private g_lngMax As Long                                            ' 最大件数
Private g_strCaption As String                                      ' プログレスバーのタイトル
Private g_strMessage As String                                      ' Window内上部のメッセージ
Private g_strStatusBar As String                                    ' ステータスバー
Private g_lngTag As Long                                            ' 中断指示
Private g_blnSwShow As Boolean                                      ' 閉じる[×]ボタン表示スイッチ
Private g_blnSwStop As Boolean                                      ' 中断機能の使用有無スイッチ
Private g_lngEnableCancelKey As Long                                ' Escキー動作の保管

'***************************************************************************************************
' ■■■ イベント ■■■
'***************************************************************************************************
'* 処理名　：Class_Initialize
'* 機能　　：クラス初期化
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：(なし)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2003年08月19日
'* 作成者　：井上　治
'* 更新日　：2019年12月29日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub Class_Initialize()
    '-----------------------------------------------------------------------------------------------
#If VBA6 Or VBA7 Then
    ' デフォルト値のセット
    g_blnSwShow = False                         ' 閉じる[×]ボタン表示
    g_blnSwStop = True                          ' 中断機能の使用有無
    g_strMessage = "処理中です。しばらくお待ち下さい．．．．"
    g_strStatusBar = ""
    ' Escキー動作を保管
    g_lngEnableCancelKey = Application.EnableCancelKey
#Else
    ' Excel97以前ではModeLessユーザーフォームは利用不可
    MsgBox "この処理はExcel2000以降で動作します。", vbExclamation
    End
#End If
End Sub

'***************************************************************************************************
'* 処理名　：Class_Terminate
'* 機能　　：クラス終了
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：(なし)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2003年08月19日
'* 作成者　：井上　治
'* 更新日　：2019年12月29日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub Class_Terminate()
    '-----------------------------------------------------------------------------------------------
    Unload frmProgressBar2
    ' Escキー動作を復帰
    Application.EnableCancelKey = g_lngEnableCancelKey
End Sub

'***************************************************************************************************
' ■■■ メソッド ■■■
'***************************************************************************************************
'* 処理名　：Show
'* 機能　　：Showメソッド
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：(なし)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2003年08月19日
'* 作成者　：井上　治
'* 更新日　：2019年12月29日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Friend Sub Show()
    '-----------------------------------------------------------------------------------------------
    g_lngWidth = 0
    g_lngTag = 0
    With frmProgressBar2
        .Height = 110
        ' 制御情報は一旦ラベルに持たせておく
        .LBL_Macro.Visible = g_blnSwShow        ' 閉じる[×]ボタン表示
        .LBL_Macro.WordWrap = g_blnSwStop       ' 中断機能の使用有無
        If g_blnSwStop = True Then
            ' 中断キーを使用する場合
            Application.EnableCancelKey = xlErrorHandler
        Else
            ' 中断キーを使用しない場合
            Application.EnableCancelKey = xlDisabled
        End If
        .Caption = g_strCaption                 ' フォームタイトル
        .LBL_Message = g_strMessage             ' フォーム画面メッセージ
        .LBL_StatusBar = g_strStatusBar         ' フォームステータス表示
        .Show vbModeless                        ' プログレスバーの表示(モードレス)
        DoEvents
    End With
End Sub

'***************************************************************************************************
'* 処理名　：Hide
'* 機能　　：Hideメソッド
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：(なし)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2003年08月19日
'* 作成者　：井上　治
'* 更新日　：2019年12月29日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Friend Sub Hide()
    '-----------------------------------------------------------------------------------------------
    frmProgressBar2.Hide
    DoEvents
End Sub

'***************************************************************************************************
'* 処理名　：EscKey
'* 機能　　：Escキー押下メソッド(フォームがアクティブでない場合の対応)
'---------------------------------------------------------------------------------------------------
'* 返り値　：終了判定(Boolean)
'* 引数　　：(なし)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2003年08月19日
'* 作成者　：井上　治
'* 更新日　：2019年12月29日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Friend Function EscKey() As Boolean
    '-----------------------------------------------------------------------------------------------
    If MsgBox("キャンセルキーが押されました。" & _
        "ここで中断して終了しますか？", _
        vbInformation + vbYesNo, Me.Caption) = vbYes Then
        frmProgressBar2.Tag = 9
        g_lngTag = 9
        EscKey = True                   ' 終了する(True)
    Else
        EscKey = False                  ' 終了しない(False)
    End If
End Function

'***************************************************************************************************
' ■■■ プロパティ ■■■
'***************************************************************************************************
' 最大件数の取得、設定
'---------------------------------------------------------------------------------------------------
Friend Property Get Max() As Long
    Max = g_lngMax
End Property
Friend Property Let Max(ByVal lngNewMax As Long)
    g_lngMax = lngNewMax
End Property

'===================================================================================================
' 現在件数の取得、設定
'---------------------------------------------------------------------------------------------------
Friend Property Get Value() As Long
    Value = g_lngNow
End Property
Friend Property Let Value(ByVal lngNewValue As Long)
    Dim lngWidth As Long, lngTag As Long
    
    On Error GoTo Value_Error
    g_lngNow = lngNewValue
    If g_lngNow >= g_lngMax Then
        frmProgressBar2.Fm_ProgressBar.Width = 200
        frmProgressBar2.Hide
        Exit Property
    End If
    ' Widthの計算
    lngWidth = Int(g_lngNow * 200& / g_lngMax)
    ' 長さに変動がなければ終了
    If lngWidth = g_lngWidth Then Exit Property
    
    ' プログレスバーの長さを変更
    frmProgressBar2.Fm_ProgressBar.Width = lngWidth
    DoEvents
    g_lngWidth = lngWidth
    lngTag = Val(frmProgressBar2.Tag)
    If lngTag = 9 Then g_lngTag = lngTag
    Exit Property
    
'===================================================================================================
' エラー処理(中断キーの対応)
Value_Error:
    If Err.Number = 18 Then
        Call EscKey
    Else
        MsgBox Err.Description, vbCritical, g_strCaption
        frmProgressBar2.Tag = 9
        g_lngTag = 9
        DoEvents
    End If
End Property

'===================================================================================================
' プログレスバーのタイトルの取得、設定
'---------------------------------------------------------------------------------------------------
Friend Property Get Caption() As String
    Caption = g_strCaption
End Property
Friend Property Let Caption(ByVal strNewCaption As String)
    g_strCaption = strNewCaption
End Property

'===================================================================================================
' プログレスバーの画面メッセージの取得、設定
'---------------------------------------------------------------------------------------------------
Friend Property Get Message() As String
    Message = g_strMessage
End Property
Friend Property Let Message(ByVal strNewMessage As String)
    g_strMessage = strNewMessage
    If g_blnShowProgressBar2 = True Then
        frmProgressBar2.LBL_Message.Caption = g_strMessage
        DoEvents
    End If
End Property

'===================================================================================================
' プログレスバーのステータス表示の取得、設定
'---------------------------------------------------------------------------------------------------
Friend Property Get StatusBar() As String
    StatusBar = g_strStatusBar
End Property
Friend Property Let StatusBar(ByVal strNewStatusBar As String)
    g_strStatusBar = strNewStatusBar
    If g_blnShowProgressBar2 = True Then
        frmProgressBar2.LBL_StatusBar.Caption = g_strStatusBar
        DoEvents
    End If
End Property

'===================================================================================================
' 閉じる[×]ボタン表示スイッチの取得、設定
'---------------------------------------------------------------------------------------------------
Friend Property Get ShowBox() As Boolean
    ShowBox = g_blnSwShow
End Property
Friend Property Let ShowBox(ByVal blnSwShow As Boolean)
    g_blnSwShow = blnSwShow
End Property

'===================================================================================================
' 中断機能の使用有無スイッチの取得、設定
'---------------------------------------------------------------------------------------------------
Friend Property Get EscStop() As Boolean
    ShowBox = g_blnSwStop
End Property
Friend Property Let EscStop(ByVal blnSwStop As Boolean)
    g_blnSwStop = blnSwStop
End Property

'===================================================================================================
' 中断指示の取得(Tag=9が中断指示)
'---------------------------------------------------------------------------------------------------
Friend Property Get Tag() As Long
    On Error Resume Next
    Tag = g_lngTag
    If Err.Number = 18 Then Call EscKey
End Property

'----------------------------------------<< End of Source >>----------------------------------------
