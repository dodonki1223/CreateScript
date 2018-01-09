'************************************************************************************************
'* プログラム名 ： IPアドレスチェックスクリプト                                                 *
'* 処理概要     ： 自身のIPアドレスをメッセージボックスで表示させます                           *
'* メモ         ： 以下のコマンドを使用してIPアドレスを取得します                               *
'*                   @for /F "delims=: tokens=2" %a in ('ipconfig ^| findstr "IP"') do @echo %a *
'* 設定         ：                                                                              *
'************************************************************************************************

'*****************************************
'* 変数                                  *
'*****************************************
Dim mObjShell : Set mObjShell = WScript.CreateObject("WScript.Shell") 
Dim mObjExec  : Set mObjExec  = mObjShell.Exec("cmd.exe /c @for /F " & """delims=: tokens=2""" & " %a in ('ipconfig ^| findstr " & """IPv4""" & "') do @echo %a")

Main()

'***********************************************************************
'* 処理名   ： メイン処理                                              *
'* 引数     ： なし                                                    *
'* 処理内容 ： メイン処理                                              *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub Main()

    'コマンドプロンプトの結果行分繰り返す
    Do Until mObjExec.StdOut.AtEndOfStream

        'IPアドレスを変数にセット(半角スペースを削除して)
        Dim mIpAdress : mIpAdress = Trim(mObjExec.StdOut.ReadLine)
    
        'IPアドレスをメッセージボックスに表示
        Dim mMsgBoxResult : mMsgBoxResult = MsgBox(mIpAdress & vbCrLf & vbCrLf & "クリップボードにコピーする場合は「はい」を押してください", vbYesNo, "IPアドレスチェック")
    
        '「はい」ボタンが押された時
        If mMsgBoxResult = vbYes Then
    
            'クリップボードコピー処理
            Dim mClipBoardCopyText : mClipBoardCopyText = "cmd.exe /c ""echo " & mIpAdress & "| clip"""
            mObjShell.Run mClipBoardCopyText, 0
        
        End If
    
    Loop

End Sub

