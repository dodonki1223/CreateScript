Echo off

Rem For文内の値を変化させるための宣言（遅延環境変数）
Setlocal enabledelayedexpansion

Rem ************************************************************************
Rem * バッチ名：Laravelのアプリケーションの環境切り替え
Rem * 処理内容：Laravelのアプリケーションの環境を切り替えるバッチです。アプリ
Rem *           ケーションフォルダ内のenvファイル(.env以外)を選択させ「.env」
Rem *           に選択したenvファイルを置き換えることでアプリケーションの環境
Rem *           を切り替えます
Rem ************************************************************************

Rem --------------------------- 設定 --------------------------- 
Rem アプリケーションパス
Set ApplicationPath=%UserProfile%\LaravelProjects\Fukully
Rem 除外ファイル名
Set ExclusionFile=.env
Rem 切り替え先ファイルパス（.envのフルパス）
Set SwitchDestFilePath=%ApplicationPath%\%ExclusionFile%
Rem 終了時に実行するエディタのパス（VSCode）
Set TextEditorPath=%SystemRoot%\..\Tools\VSCode\Code.exe

Rem ---------------------- メイン処理部分 ---------------------- 
Rem タイトル表示処理
Call :DisplayTitle

Rem 切り替えファイルを選択します
Call :ChooseTargetFile

Rem 切り替えファイルを.envファイルに変更
Call :SwitchChangeFileToEnvFile

Rem 終了処理
Call :EndProcess

Rem ---------------------- ラベル処理部分 ---------------------- 

Rem ************************************************************************
Rem * バッチ名：Laravelのアプリケーションの環境切り替え
Rem * 処理内容：Laravelのアプリケーションの環境を切り替えるバッチです。アプリ
Rem *           ケーションフォルダ内のenvファイル(.env以外)を選択させ「.env」
Rem *           に選択したenvファイルを置き換えることでアプリケーションの環境
Rem *           を切り替えます
Rem ************************************************************************
:DisplayTitle

    Call :DisplayNewLine
    Call :DisplayMessage " ************************************************************************  " 0 0
    Call :DisplayMessage " * バッチ名：Laravelのアプリケーションの環境切り替え                       " 0 0
    Call :DisplayMessage " * 処理内容：Laravelのアプリケーションの環境を切り替えるバッチです。アプリ " 0 0
    Call :DisplayMessage " *           ケーションフォルダ内のenvファイル(.env以外)を選択させ「.env」 " 0 0
    Call :DisplayMessage " *           に選択したenvファイルを置き換えることでアプリケーションの環境 " 0 0
    Call :DisplayMessage " *           を切り替えます                                                " 0 0
    Call :DisplayMessage " ************************************************************************  " 0 0
    Call :DisplayNewLine
    
    Exit /b

Rem ***************************************************************
Rem * 処 理 名：切り替え対象ファイルの選択
Rem * 引    数：なし
Rem * 処理内容：Laravelのアプリケーションフォルダ内に格納されている
Rem *           envファイル(.env以外)を表示し切り替えファイルをユー
Rem *           ザーに対話して選択してもらう
Rem ***************************************************************
:ChooseTargetFile

    Call :DisplayMessage "★切り替えファイルを選択します．．．" 1 1

    Rem アプリケーションフォルダ内のenvファイル(.env以外)の一覧を表示
    Set FileListCounter=0
    For /f %%i In ('Dir %ApplicationPath%\*.env /b') Do (

        Rem .envファイル以外の時のみ表示
        if not %%i==%ExclusionFile% (
            Set /a FileListCounter=FileListCounter+1
            Echo !FileListCounter!：「%%i」
        ) 

    )

    Rem 切り替えファイルを選択
    Call :DisplayMessage "切り替えファイルの番号を入力して下さい（1～!FileListCounter!）" 1 0
    Set /p TargetFileNo="切り替えファイル番号の入力　＞　"

    Rem 切り替えファイルの表示
    Set ChooseFileNoCounter=0
    For /f %%i In ('Dir %ApplicationPath%\*.env /b') Do (

        Rem .envファイル以外の時
        if not %%i==%ExclusionFile% (
    
            Set /a ChooseFileNoCounter=ChooseFileNoCounter+1
        
            Rem 切り替えファイルの番号と一致した時
            If !ChooseFileNoCounter! Equ %TargetFileNo% (
        
                Rem ファイル名をセットする
                Set TargetFileName=%%i
            
            )
        )
    )
    
    Rem 切り替えファイルの表示
    Call :DisplayMessage "切り替えファイル名：%TargetFileName%" 1 0

    Rem 切り替えファイルが未入力の時（不正な値を入力した時）
    Rem ※もう一度、「切り替えファイルの選択」処理を実行する
    If "%TargetFileName%" Equ "" (

        Call :DisplayMessage "エラー：切り替えファイル名が取得できませんでした。「1～!FileListCounter!」で入力して下さい。" 1 0
        Call :ChooseTargetFile

    ) 

    Exit /b

Rem ***************************************************************
Rem * 処 理 名：切り替えファイルを.envファイルへ切り替え
Rem * 引    数：なし
Rem * 処理内容：選択された切り替えファイルを.envファイルに切り替えます
Rem ***************************************************************
:SwitchChangeFileToEnvFile

    Rem 切り替えファイルのフルパスを取得
    Set ChangeFilePath=%ApplicationPath%\%TargetFileName%
    
    Call :DisplayMessage "ファイルの切り替えを行います" 1 1
    
    Rem .envファイルの切り替え処理を実行
    Rem ※コピーコマンドを実行して切り替える
    copy %ChangeFilePath% %SwitchDestFilePath%

    Call :DisplayMessage "ファイルの切り替えが完了しました" 1 1

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

    Call :DisplayMessage "★ファイルの切り替え処理が終了しました．．．           " 1 1
    Call :DisplayMessage "  テキストエディタで「%ExclusionFile%」を開きます      " 0 1
    Call :DisplayMessage "※念のためキャッシュ系のコマンドを実行しておいて下さい " 1 1
    Pause

    start %TextEditorPath% %SwitchDestFilePath%

    Rem 処理の終了
    Exit
    