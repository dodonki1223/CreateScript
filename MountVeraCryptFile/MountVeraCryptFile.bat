Echo off

Rem For文内の値を変化させるための宣言（遅延環境変数）
Setlocal enabledelayedexpansion

Rem ***************************************************************
Rem * バッチ名：VeraCryptで作成した暗号化コンテナのマウントバッチ
Rem * 処理内容：VeraCryptで作成した暗号化コンテナをマウントします
Rem ***************************************************************

Rem --------------------------- 設定 --------------------------- 
Rem VeraCryptパス ※「実行ドライブ文字列:/Tools\VeraCrypt」になります
Set VeraCryptPath=%~d0/Tools\VeraCrypt\
Rem 暗号化コンテナファイル格納フォルダ ※「実行ドライブ文字列:/Tools\VeraCrypt\暗号化コンテナファイル格納フォルダ」
Set EncryptionContainerStorageFolder=EncryptionContainerFile\

Rem ---------------------- メイン処理部分 ---------------------- 
Rem タイトル表示処理
Call :DisplayTitle

Rem 現在のディレクトリをVeraCryptディレクトリに変更
Call :ChangeVeraCryptDirectory

Rem マウント対象ファイルを選択します
Call :ChooseTargeFile

Rem 暗号化コンテナのマウント処理
Call :MountEncryptionContainer

Rem 終了処理
Call :EndProcess

Rem ---------------------- ラベル処理部分 ---------------------- 

Rem ***************************************************************
Rem * 処 理 名：タイトルの表示
Rem * 引    数：なし
Rem * 処理内容：VeraCryptで作成した暗号化コンテナのマウントバッチの
Rem *           タイトルを表示する
Rem ***************************************************************
:DisplayTitle

    Call :DisplayNewLine
    Call :DisplayMessage "************************************************************** " 0 0
    Call :DisplayMessage "  バッチ名：VeraCryptで作成した暗号化コンテナのマウントバッチ  " 0 0
    Call :DisplayMessage "  処理概要：VeraCryptで作成した暗号化コンテナを空いているドラ  " 0 0
    Call :DisplayMessage "            イブレターにマウント処理を行う                     " 0 0
    Call :DisplayMessage "************************************************************** " 0 0
    Call :DisplayNewLine
    
    Exit /b

Rem ***************************************************************
Rem * 処 理 名：現在のディレクトリをVeraCryptディレクトリに変更
Rem * 引    数：なし
Rem * 処理内容：現在のディレクトリ（batを実行した場所）からVeraCrypt
Rem *           のディレクトリに変更する
Rem ***************************************************************
:ChangeVeraCryptDirectory

    Call :DisplayMessage "★現在のディレクトリをVeraCryptのディレクトリに変更します．．．" 0 0
    
    Rem ディレクトリを「VeraCrypt」ディレクトリに変更
    Cd %VeraCryptPath%
    
    Call :DisplayMessage "現在のディレクトリをVeraCryptのディレクトリに変更しました．．．" 1 0
    Call :DisplayMessage "＞ %VeraCryptPath%                                             " 0 0
    
    Exit /b

Rem ***************************************************************
Rem * 処 理 名：マウント対象ファイルの選択
Rem * 引    数：なし
Rem * 処理内容：暗号化コンテナファイルが格納されているフォルダ内の
Rem *           ファイルを表示し対象ファイルをユーザーに対話して選
Rem *           択してもらう
Rem ***************************************************************
:ChooseTargeFile

    Call :DisplayMessage "★マウント対象ファイルを選択します．．．" 1 1

    Rem 暗号化コンテナファイルが格納されているフォルダ内のファイルの一覧を表示
    Set FileListCounter=0
    For /f %%i In ('Dir %EncryptionContainerStorageFolder% /b') Do (

        Set /a FileListCounter=FileListCounter+1
        Echo !FileListCounter!：「%%i」

    )

    Rem 対象コンテナファイルを選択
    Call :DisplayMessage "マウントするコンテナファイルの番号を入力して下さい（1～!FileListCounter!）" 1 0
    Set /p TargetFileNo="コンテナファイル番号の入力　＞　"

    Rem 対象コンテナファイルの表示
    Set ChooseFileNoCounter=0
    For /f %%i In ('Dir EncryptionContainerFile\ /b') Do (
    
        Set /a ChooseFileNoCounter=ChooseFileNoCounter+1
        
        Rem 対象コンテナファイルの番号と一致した時
        If !ChooseFileNoCounter! Equ %TargetFileNo% (
        
            Rem ファイル名をセットする
            Set TargetFileName=%%i
            
        )
    )
    
    Rem 対象コンテナファイルの表示
    Call :DisplayMessage "対象ファイル名：%TargetFileName%" 1 0

    Rem 対象コンテナファイルが未入力の時（不正な値を入力した時）
    Rem ※もう一度、「マウント対象ファイルの選択」処理を実行する
    If "%TargetFileName%" Equ "" (

        Call :DisplayMessage "エラー：対象ファイル名が取得できませんでした。「1～!FileListCounter!」で入力して下さい。" 1 0
        Call :ChooseTargeFile

    ) 

    Exit /b

Rem ***************************************************************
Rem * 処 理 名：暗号化コンテナのマウント処理
Rem * 引    数：なし
Rem * 処理内容：暗号化コンテナをマウントするかどうかユーザーに対話
Rem ***************************************************************
:MountEncryptionContainer

    Rem マウント対象コンテナファイルを取得
    Set TargetContainerFile=%EncryptionContainerStorageFolder%%TargetFileName%

    Call :DisplayMessage "★暗号化コンテナのマウント処理を行います．．．              " 1 1
    Call :DisplayMessage "※実行されるコマンドは以下になります                        " 0 0
    Call :DisplayMessage " ＞ VeraCrypt /q /e /v %TargetContainerFile%                " 0 0

    Rem 処理を続行するかユーザーに対話
    Rem ※空Enterすると落ちるため、予め適当な値をセットしておく
    Call :DisplayMessage "下記メッセージは(「Y」又は「y」)以外は処理を終了します．．．" 1 0
    Set RunContinueResult=KaraMojiTaiou
    Set /p RunContinueResult="上記情報を使用して処理を実行しますか？(y/n)　＞　"

    Rem 大文字/小文字変換(Y以外は全てキャンセル扱い) 
    Set RunContinueResult=%RunContinueResult:y=Y%%

    Rem Y以外の入力の時は処理を終了
    If /i Not %RunContinueResult%==Y Call :EndProcess

    Rem 暗号化コンテナのマウント処理
    Rem /q：バックグラウンドでVeraCryptを実行させる ※パスワード入力ボックスのみ表示させる
    Rem /e：マウント後エクスプローラーで開く
    Rem /v：マウントするファイル名
    VeraCrypt /q /e /v %TargetContainerFile%

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

    Call :DisplayMessage "★マウント処理が終了しました．．．" 1 1
    Pause

    Rem 処理の終了
    Exit
    