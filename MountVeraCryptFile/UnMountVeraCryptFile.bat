Echo off
Rem ********************************************************************
Rem * バッチ名：VeraCryptでマウントしたドライブをアンマウントするバッチ
Rem * 処理内容：VeraCryptでマウントされたドライブをアンマウントします
Rem ********************************************************************

Rem --------------------------- 設定 --------------------------- 
Rem VeraCryptパス ※「実行ドライブ文字列:/Tools\VeraCrypt」になります
Set VeraCryptPath=%~d0/Tools\VeraCrypt\

Rem ---------------------- メイン処理部分 ---------------------- 
Rem タイトル表示処理
Call :DisplayTitle

Rem 現在のディレクトリをVeraCryptディレクトリに変更
Call :ChangeVeraCryptDirectory

Rem アンマウントドライブ名を取得
Call :InputUnMountDriveNmae

Rem VeraCryptでマウントしたドライブのアンマウント処理
Call :UnMountVeraCryptDrive

Rem 終了処理
Call :EndProcess

Rem ---------------------- ラベル処理部分 ---------------------- 

Rem ********************************************************************
Rem * 処 理 名：タイトルの表示
Rem * 引    数：なし
Rem * 処理内容：VeraCryptでマウントされたドライブをアンマウントするバッ
Rem *           チのタイトルを表示します
Rem ********************************************************************
:DisplayTitle

    Call :DisplayNewLine
    Call :DisplayMessage "******************************************************************** " 0 0
    Call :DisplayMessage "  バッチ名：VeraCryptでマウントしたドライブをアンマウントするバッチ  " 0 0
    Call :DisplayMessage "  処理概要：VeraCryptでマウントされたドライブをアンマウントします    " 0 0
    Call :DisplayMessage "******************************************************************** " 0 0
    Call :DisplayNewLine
    
    Exit /b

Rem ***************************************************************
Rem * 処 理 名：現在のディレクトリをVeraCryptディレクトリに変更
Rem * 引    数：なし
Rem * 処理内容：現在のディレクトリ（batを実行した場所）からVeraCrypt
Rem *           のディレクトリに変更する
Rem ***************************************************************
:ChangeVeraCryptDirectory

    Call :DisplayMessage "★現在のディレクトリをVeraCryptのディレクトリに変更します．．．  " 0 0
    
    Rem ディレクトリを「VeraCrypt」ディレクトリに変更
    Cd %VeraCryptPath%
    
    Call :DisplayMessage "現在のディレクトリをVeraCryptのディレクトリに変更しました．．．" 1 0
    Call :DisplayMessage "＞ %VeraCryptPath%                                             " 0 0
    
    Exit /b

Rem ***************************************************************
Rem * 処 理 名：アンマウント対象ドライブ名を取得
Rem * 引    数：なし
Rem * 処理内容：アンマウントドライブ名をユーザーに対話し取得する。「Y」
Rem             又は「y」が指定された場合は再度ユーザーに対話
Rem ***************************************************************
:InputUnMountDriveNmae

    Call :DisplayMessage "★アンマウント対象ドライブ指定処理を行います．．．                    " 1 1

    Rem アンマウントするドライブ指定（ユーザーに対話）
    Set /p UnMountDrive="アンマウントするドライブを指定して下さい　＞　"
    
    Rem ドライブ指定が空文字の時
    If "%UnMountDrive%" Equ "" (
    
        Rem エラーを表示し、アンマウント対象ドライブ名取得処理をもう一度実行
        Call :DisplayMessage "エラー：ドライブの指定が正しくありません" 1 0
        Call :InputUnMountDriveNmae
        Exit /b
        
    )
    
    Rem 対象ドライブ表示処理
    Call :DisplayMessage "アンマウント対象ドライブ：%UnMountDrive%                              " 1 1

    Rem 処理を続行するかユーザーに対話 
    Rem ※空Enterすると落ちるため、予め適当な値をセットしておく
    Call :DisplayMessage "下記メッセージは(「Y」又は「y」)以外はドライブ指定を再度行います．．．" 0 0
    Set RunContinueResult=KaraMojiTaiou
    Set /p RunContinueResult="上記情報を使用して処理を実行しますか？(y/n)　＞　"
    
    Rem 大文字/小文字変換(Y以外は全てキャンセル扱い) 
    Set RunContinueResult=%RunContinueResult:y=Y%%

    Rem Y以外の入力の時は「アンマウント対象ドライブ名を取得」へ
    If /i Not %RunContinueResult%==Y Call :InputUnMountDriveNmae

    Exit /b

Rem ***************************************************************
Rem * 処 理 名：VeraCryptでマウントしたドライブのアンマウント処理
Rem * 引    数：なし
Rem * 処理内容：VeraCryptでマウントしたドライブをユーザーに対話し
Rem *           対象のドライブをアンマウントする
Rem ***************************************************************
:UnMountVeraCryptDrive

    Call :DisplayMessage "★VeraCryptでマウントされたドライブのアンマウント処理を行います．．．" 1 1
    
    Call :DisplayMessage "※実行されるコマンドは以下になります          " 0 0
    Call :DisplayMessage " ＞ VeraCrypt /q /d %UnMountDrive%            " 0 1

    Rem 処理を続行するかユーザーに対話
    Rem ※空Enterすると落ちるため、予め適当な値をセットしておく
    Call :DisplayMessage "下記メッセージは(「Y」又は「y」)以外は処理を終了します．．．         " 0 0
    Set RunContinueResult=KaraMojiTaiou
    Set /p RunContinueResult="上記情報を使用して処理を実行しますか？(y/n)　＞　"

    Rem 大文字/小文字変換(Y以外は全てキャンセル扱い) 
    Set RunContinueResult=%RunContinueResult:y=Y%%

    Rem Y以外の入力の時は処理を終了
    If /i Not %RunContinueResult%==Y Call :EndProcess

    Rem 暗号化コンテナのアンマウント処理
    Rem /q：バックグラウンドでVeraCryptを実行させる ※パスワード入力ボックスのみ表示させる
    Rem /d：アンマウントするドライブ名
    VeraCrypt /q /d %UnMountDrive%

    Exit /b

Rem ***************************************************************
Rem * 処 理 名：メッセージの表示
Rem * 引    数：1 表示させるメッセージ
Rem *           2 表示させるメッセージ前に改行を含めるかどうか（1は含める、それ以外は含めない）
Rem *           3 表示させるメッセージ後に改行を含めるかどうか（1は含める、それ以外は含めない）
Rem * 処理内容：表示させるメッセージの前後に改行を含めて表示するか
Rem *           どうかを引数に応じて行う
Rem *           使用方法
Rem *             Call :DisplayMessage "ああああ aaaaa bbbbb" 1 1
Rem *             ※引数は必ず3つ渡すこと
Rem *               表示するメッセージは必ずダブルクォーテーションで囲むこと
Rem ***************************************************************
:DisplayMessage

    Rem メッセージ前に改行を含める
    If %~2 Equ 1 (
        Call :DisplayNewLine
    ) 

    Rem ダウブルクォーテーションを削除して表示
    Echo %~1

    Rem メッセージ後に改行を含める
    If %~3 Equ 1 (
        Call :DisplayNewLine
    )
    
    Exit /b

Rem ***************************************************************
Rem * 処 理 名：改行メッセージの表示
Rem * 引    数：なし
Rem * 処理内容：コマンドプロンプトに改行を表示させる
Rem ***************************************************************
:DisplayNewLine

    Rem 改行を表示
    Echo;
    
    Exit /b

Rem ***************************************************************
Rem * 処 理 名：終了処理
Rem * 引    数：なし
Rem * 処理内容：バッチ処理を終了させる
Rem ***************************************************************
:EndProcess

    Call :DisplayMessage "★アンマウント処理が終了しました．．．" 1 1
    Pause

    Rem 処理の終了
    Exit
    