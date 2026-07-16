VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgressBar2 
   Caption         =   "プログレスバー"
   ClientHeight    =   1632
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4332
   OleObjectBlob   =   "frmProgressBar2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmProgressBar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************************
'   プログレスバーを表示するフォーム本体                        frmProgressBar2(UserForm)
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'***************************************************************************************************
'変更日付 Rev  変更履歴内容------------------------------------------------------------------------>
'03/08/19(1.00)新規作成
'19/11/02(1.10)64ビット版Excelの対応(API関連)
'19/12/29(1.11)処理記述整理(標準化準拠)
'***************************************************************************************************
Option Explicit
'===================================================================================================
Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000
' 64ビット対応
#If VBA7 Then
    #If Win64 Then
        ' ウィンドウに関する情報を返す
        Private Declare PtrSafe Function GetWindowLongPtr Lib "USER32.dll" _
            Alias "GetWindowLongPtrA" _
            (ByVal hWnd As LongPtr, _
             ByVal nIndex As Long) As LongPtr
        ' ウィンドウの属性を変更
        Private Declare PtrSafe Function SetWindowLongPtr Lib "USER32.dll" _
            Alias "SetWindowLongPtrA" _
            (ByVal hWnd As LongPtr, _
             ByVal nIndex As Long, _
             ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        ' ウィンドウに関する情報を返す
        Private Declare PtrSafe Function GetWindowLongPtr Lib "USER32.dll" _
            Alias "GetWindowLongA" _
            (ByVal hWnd As LongPtr, _
             ByVal nIndex As Long) As LongPtr
        ' ウィンドウの属性を変更
        Private Declare Function SetWindowLongPtr Lib "USER32.dll" _
            Alias "SetWindowLongA" _
            (ByVal hWnd As LongPtr, _
             ByVal nIndex As Long, _
             ByVal dwNewLong As LongPtr) As LongPtr
    #End If
' Activeなウィンドウのハンドルを取得
Private Declare PtrSafe Function GetActiveWindow Lib "USER32.dll" _
    () As LongPtr
' メニューバーを再描画
Private Declare PtrSafe Function DrawMenuBar Lib "USER32.dll" _
    (ByVal hWnd As LongPtr) As LongPtr
#Else
' ウィンドウに関する情報を返す
Private Declare Function GetWindowLong Lib "USER32.dll" _
    Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long
' ウィンドウの属性を変更
Private Declare Function SetWindowLong Lib "USER32.dll" _
    Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long
' Activeなウィンドウのハンドルを取得
Private Declare Function GetActiveWindow Lib "USER32.dll" _
    () As Long
' メニューバーを再描画
Private Declare Function DrawMenuBar Lib "USER32.dll" _
    (ByVal hWnd As Long) As Long
#End If

'***************************************************************************************************
'   ■■■ フォームイベント ■■■
'***************************************************************************************************
'* 処理名　：CMD_CANCEL_Click
'* 機能　　：キャンセル(隠しボタン)のイベント(Escキーで本イベントが発生する)
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
Private Sub CMD_CANCEL_Click()
    '-----------------------------------------------------------------------------------------------
    ' 中断機能の使用有無はラベルのWordWrapプロパティを使う
    If LBL_Macro.WordWrap = True Then
        If MsgBox("キャンセルキーが押されました。" & _
            "ここで中断して終了しますか？", _
            vbInformation + vbYesNo, Me.Caption) = vbYes Then
            Me.Tag = 9
        End If
    End If
End Sub

'***************************************************************************************************
'* 処理名　：UserForm_Activate
'* 機能　　：フォームの表示初期化
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
Private Sub UserForm_Activate()
    '-----------------------------------------------------------------------------------------------
#If VBA7 Then
    Dim hWnd As LongPtr
    Dim Wnd_STYLE As LongPtr
#Else
    Dim hWnd As Long
    Dim Wnd_STYLE As Long
#End If
    ' 閉じる[×]ボタンの有効／無効はラベルのVisibleプロパティを使う
    If LBL_Macro.Visible <> True Then
        ' 閉じる[×]ボタンを無効にする
        hWnd = GetActiveWindow()
#If VBA7 Then
        Wnd_STYLE = GetWindowLongPtr(hWnd, GWL_STYLE)
        Wnd_STYLE = Wnd_STYLE And (Not WS_SYSMENU)
        SetWindowLongPtr hWnd, GWL_STYLE, Wnd_STYLE
#Else
        Wnd_STYLE = GetWindowLong(hWnd, GWL_STYLE)
        Wnd_STYLE = Wnd_STYLE And (Not WS_SYSMENU)
        SetWindowLong hWnd, GWL_STYLE, Wnd_STYLE
#End If
        DrawMenuBar hWnd
    End If
    Me.Tag = 0
    Fm_ProgressBar.Width = 0.01
    g_blnShowProgressBar2 = True
End Sub

'***************************************************************************************************
'* 処理名　：UserForm_Deactivate
'* 機能　　：フォーム不活性化
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
Private Sub UserForm_Deactivate()
    '-----------------------------------------------------------------------------------------------
    g_blnShowProgressBar2 = False
End Sub

'***************************************************************************************************
'* 処理名　：UserForm_QueryClose
'* 機能　　：Closeイベント
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：(既定)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2003年08月19日
'* 作成者　：井上　治
'* 更新日　：2019年12月29日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '-----------------------------------------------------------------------------------------------
    ' 閉じる[×]ボタンは許可しない(表示されない場合は動作しない)
    If CloseMode = vbFormControlMenu Then
        ' Escキーと同様処理とする
        Call CMD_CANCEL_Click
        Cancel = True
    End If
End Sub

'----------------------------------------<< End of Source >>----------------------------------------
