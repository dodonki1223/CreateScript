'**************************************************************************************
'* プログラム名 ： Access実行スクリプト                                               *
'* 処理概要     ： Accessを実行するだけのスクリプト                                   *
'* メモ         ： Accessがインストールされていなければ実行されません                 *
'* 設定         ：                                                                    *
'**************************************************************************************

'*****************************************
'* 変数                                  *
'*****************************************
Dim mObjShell : Set mObjShell = WScript.CreateObject("WScript.Shell")

Main()

'***********************************************************************
'* 処理名   ： メイン処理                                              *
'* 引数     ： なし                                                    *
'* 処理内容 ： メイン処理                                              *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub Main()

    'Access実行処理
    mObjShell.Run "msaccess.exe"

    'オブジェクトの破棄処理
    Set mObjShell = Nothing

End Sub
