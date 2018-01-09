@Echo off
Rem ***************************************************************
Rem * バッチ名：シンボリック　リンク作成バッチ
Rem * 処理内容：シンボリック　リンク作成条件をユーザーに入力させシ
Rem *           ンボリック　リンクを作成します
Rem ***************************************************************

Rem --------------------------- 設定 --------------------------- 
Rem ログファイルの保存場所
Set LogFileFullPath=%~dp0\CreateSymbolicLink.log

Rem ---------------------- メイン処理部分 ---------------------- 
Rem タイトル表示処理
Call :DisplayTitle

Rem 設定の入力
Call :InputSettings

Rem 作成リンクの種類をコマンド変換処理
Call :ConvertCreateLinkTypeToCommand

Rem バッチファイル実行日時を作成する
Call :CreateRunDateAndTime

Rem 処理を続行するかユーザーに対話
Call :IsRun

Rem シンボリックリンク作成処理の実行
Call :CreateSymbolicLink

Rem 終了処理
Call :EndProcess 1

Rem ---------------------- ラベル処理部分 ---------------------- 

Rem ***************************************************************
Rem * 処 理 名：タイトルの表示
Rem * 引    数：なし
Rem * 処理内容：シンボリック　リンク作成バッチのタイトルを表示する
Rem *           パラメータ等も一緒に表示させる
Rem ***************************************************************
:DisplayTitle

    Call :DisplayNewLine
    Call :DisplayMessage "*********************************************************               " 0 0
    Call :DisplayMessage "  バッチ名：シンボリック　リンク作成バッチ                              " 0 0
    Call :DisplayMessage "  処理概要：①作成リンクの種類の指定                                    " 0 0
    Call :DisplayMessage "            ②リンクの指定                                              " 0 0
    Call :DisplayMessage "            ③ターゲットの指定                                          " 0 0
    Call :DisplayMessage "            ④シンボリック　リンク作成処理                              " 0 0
    Call :DisplayMessage "*********************************************************               " 0 0
    Call :DisplayNewLine
    Call :DisplayMessage "  MKLINK [ [/D] [/H] [/J] ] リンク ターゲット                           " 0 0
    Call :DisplayNewLine
    Call :DisplayMessage "          /D：ディレクトリのシンボリック　リンクを作成します。          " 0 0
    Call :DisplayMessage "              既定では、ファイルのシンボリック　リンクが作成されます。  " 0 0
    Call :DisplayMessage "          /H：シンボリック　リンクではなく、ハード　リンクを作成します。" 0 0
    Call :DisplayMessage "          /J：ディレクトリ　ジャンクションを作成します。                " 0 0
    Call :DisplayMessage "      リンク：新しいシンボリック　リンク名を指定します。                " 0 0
    Call :DisplayMessage "  ターゲット：新しいリンクが参照するパス（相対または絶対）を指定します。" 0 0
    
    Exit /b
    
Rem ***************************************************************
Rem * 処 理 名：設定の入力
Rem * 引    数：なし
Rem * 処理内容：作成リンクの種類、リンク、ターゲットの指定をユーザー
Rem *           に対話し入力させる。入力内容が不正の場合は再度、入力
Rem *           させる
Rem ***************************************************************
:InputSettings

    Call :DisplayMessage "★設定の入力を行います．．．" 1 0

    Rem 作成リンクの種類を指定
    Call :InputCreateLinkType
    
    Rem リンクを指定
    Call :InputLink
    
    Rem ターゲットを指定
    Call :InputTarget
    
    Exit /b
    
    Rem 作成リンクの種類の指定
    :InputCreateLinkType
    
        Rem ユーザーに作成リンクの入力を対話
        Call :DisplayMessage "①作成リンクの種類の指定を行います。（1：/D、2：/H、3：/J）" 1 0
        Set /p CreateLinkType="作成リンクの種類の入力　＞　"
       
        Rem 入力されたリンクの種類が空文字の時
        If "%CreateLinkType%" Equ "" (
            Call :DisplayMessage "エラー：作成リンクの種類の指定が正しくありません「1」、「2」、「3」で入力して下さい。" 1 0
            Call :InputCreateLinkType
        ) 
       
        Rem 入力された作成リンクの種類が「1,2,3」以外の時はもう１度入力させる
        If %CreateLinkType% Equ 1 (
            Exit /b
        ) Else If %CreateLinkType% Equ 2 (
            Exit /b
        ) Else If %CreateLinkType% Equ 3 (
            Exit /b
        ) Else (
            Call :DisplayMessage "エラー：作成リンクの種類の指定が正しくありません「1」、「2」、「3」で入力して下さい。" 1 0
            Call :InputCreateLinkType
        )
        
        Exit /b
        
    Rem リンクを指定
    :InputLink
    
        Rem ユーザーにリンクの入力を対話
        Call :DisplayMessage "②リンクの指定を行います。" 1 0
        Set /p Link="リンクの入力　＞　"

        If "%Link%" Equ "" (
            Call :DisplayMessage "エラー：リンクの指定がありません。" 1 0
            Call :InputLink
        )
                
        Exit /b

    Rem ターゲットを指定
    :InputTarget

        Rem ユーザーにターゲットの入力を対話
        Call :DisplayMessage "③ターゲットの指定を行います。" 1 0
        Set /p Target="ターゲットの入力　＞　"
    
        if "%Target%" Equ "" (
            Call :DisplayMessage "エラー：ターゲットの指定がありません。" 1 0
            Call :InputTarget
        ) 
    
        Exit /b

Rem ***************************************************************
Rem * 処 理 名：作成リンクの種類をコマンドに変換
Rem * 引    数：なし
Rem * 処理内容：入力された作成リンクの種類（1,2,3）を実際にコマンド
Rem *           に変換する
Rem ***************************************************************
:ConvertCreateLinkTypeToCommand

    If %CreateLinkType% Equ 1 (
        Set CreateLinkType=/D
    ) Else If %CreateLinkType% Equ 2 (
        Set CreateLinkType=/H
    ) Else If %CreateLinkType% Equ 3 (
        Set CreateLinkType=/J
    )

    Exit /b

Rem ***************************************************************
Rem * 処 理 名：処理を実行するかユーザーに対話
Rem * 引    数：なし
Rem * 処理内容：ユーザーの入力内容を表示し、処理を実行するかユー
Rem *           ザーに対話し実行しない場合は設定の入力を再度行う
Rem ***************************************************************
:IsRun

    Rem ユーザーの入力内容を表示
    Call :DisplayNewLine
    Call :DisplayMessage "********************************************************" 0 0
    Call :DisplayMessage "作成リンクの種類：%CreateLinkType%                      " 0 0
    Call :DisplayMessage "          リンク：%Link%                                " 0 0
    Call :DisplayMessage "      ターゲット：%Target%                              " 0 0
    Call :DisplayMessage "        実行日時：%RunDateAndTime%                      " 0 0
    Call :DisplayMessage "********************************************************" 0 0
    Call :DisplayNewLine
    
    Rem 処理を続行するかユーザーに対話
    Rem ※空Enterすると落ちるため、予め適当な値をセットしておく
    Call :DisplayMessage "★下記メッセージは(「Y」又は「y」)以外は処理を終了します．．．" 0 0
    Set RunContinueResult=KaraMojiTaiou
    Set /p RunContinueResult="上記情報を使用して処理を実行しますか？(y/n)　＞　"

    Rem 大文字/小文字変換(Y以外は全てキャンセル扱い) 
    Set RunContinueResult=%RunContinueResult:y=Y%%
    
    Rem Y以外の入力の時は処理を終了
    If /i Not %RunContinueResult%==Y Call :EndProcess 0

    Exit /b


Rem ***************************************************************
Rem * 処 理 名：シンボリック　リンクの作成処理
Rem * 引    数：なし
Rem * 処理内容：入力された設定を使用し「MKLINK」コマンドを実行する
Rem ***************************************************************
:CreateSymbolicLink

    Rem 実行コマンドを表示させる
    Call :DisplayMessage "★実際に実行されるコマンドは以下になります " 1 0
    Call :DisplayMessage " ＞ mklink %CreateLinkType% %Link% %Target%" 0 1
    Pause

    Rem ログファイルへ結果を書き込む
    Call :WriteLog "★実行日★"
    Call :WriteLog "%RunDateAndTime%"
    Call :WriteLog "★実行コマンド★"
    Call :WriteLog "mklink %CreateLinkType% %Link% %Target%"
    Call :WriteLog "★実行結果★"

    Rem 「mklink」コマンドを実行 ※ログファイルへ実行の結果を書き込む
    Call :DisplayNewLine
    mklink %CreateLinkType% %Link% %Target% >> %LogFileFullPath%
    
    Rem 改行をログファイルへ書き込む
    Call :WriteLog ";"
    Call :WriteLog ";"
    
    Exit /b

Rem ***************************************************************
Rem * 処 理 名：バッチファイルの実行日時を作成
Rem * 引    数：なし
Rem * 処理内容：実行の日時を作成する
Rem *           実行日時の形式は「9999/99/99 HH:mm:ss」になります
Rem ***************************************************************
:CreateRunDateAndTime

    Rem 現在日付の作成
    Set Date=%date:~-10,4%/%date:~-5,2%/%date:~-2,2%

    Rem 現在時間の作成
    Set Time=%time: =0%
    Set Time=%time:~0,2%:%time:~3,2%:%time:~6,2%

    Rem 現在日付と現在時間をつなげる
    Set RunDateAndTime=%Date% %Time%
    
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
Rem * 処 理 名：ログファイルへ書き込み
Rem * 引    数：1 ログファイルへ書き込む内容
Rem * 処理内容：ログファイルへログを書き込む
Rem             ログファイルへ書き込む内容が「;」の時は改行を書き込む
Rem *           使用方法
Rem *             Call :WriteLog "ああああ aaaaa bbbbb"
Rem *             Call :WriteLog ";" 改行を書き込む
Rem *             ※引数は必ず渡すこと
Rem *               ログファイルへ書き込む内容は必ずダブルクォーテーションで囲むこと
Rem ***************************************************************
:WriteLog

    Rem 引数が「;」の時
    If "%~1"==";" (
    
        Rem 改行をログファイルへ書き込む
        Echo; >> %LogFileFullPath%
        
    ) Else (
    
        Rem ダウブルクォーテーションを削除してログファイルへ書き込み
        Echo %~1 >> %LogFileFullPath%
    
    )

    Exit /b

Rem ***************************************************************
Rem * 処 理 名：終了処理
Rem * 引    数：1 ログファイルを表示するか（1は表示、それ以外は非表示）
Rem * 処理内容：バッチ処理を終了させる
Rem *           引数が1の時はログファイルを実行してユーザーにログ
Rem *           ファイルの内容を表示する
Rem ***************************************************************
:EndProcess

    Rem 1つ目の引数が「1」の時
    If %~1 Equ 1 (
    
        Rem 処理の終了確認
        Call :DisplayMessage "★ログファイルを実行して処理を終了します．．．" 0 1
        Pause

        Rem ログファイルの実行処理
        Start %LogFileFullPath%
        
    ) 

    Rem 処理の終了
    Exit
    