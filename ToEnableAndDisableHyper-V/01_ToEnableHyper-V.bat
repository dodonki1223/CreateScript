Echo off

Rem For文内の値を変化させるための宣言（遅延環境変数）
Setlocal enabledelayedexpansion

Rem ***************************************************************
Rem * バッチ名：Hyper-V機能をON・OFFにするバッチ
Rem * 処理内容：Hyper-V機能の現在の状態を表示し、Hyper-V機能をON・OFF
Rem *           のどちらを実行するかユーザーに対話して選択してもらう
Rem *           Hyper-V機能の変更後は再起動するかどうかユーザーに対
Rem *           話して実行する
Rem ***************************************************************

Rem --------------------------- 設定 --------------------------- 
Rem Hyper-Vの現在の状態を表示するbatファイル名
Set NowHyper-VStatusBat=04_CheckHyper-V-Status.bat

Rem ---------------------- メイン処理部分 ---------------------- 
Rem タイトル表示処理
Call :DisplayTitle

Rem カレントディレクトリをbatファイルの実行されたディレクトリに変更
Call :ChangeRunBatDir

Rem Hyper-V機能の現在の状態を表示
Call :DisplayHyper-VStatus

Rem 実行batファイルを選択します
Call :ChooseRunBatFile

Rem 選択した実行batファイルを管理者で実行する
Call :RunBatFileAsAdmin

Rem 再起動を行うかどうかユーザーに対話
Call :IsReStart

Rem 終了処理
Call :EndProcess

Rem ---------------------- ラベル処理部分 ---------------------- 

Rem ************************************************************************
Rem * バッチ名：タイトルの表示
Rem * 引    数：なし
Rem * 処理内容：Hyper-V機能をON・OFFにするバッチのタイトルを表示します
Rem ************************************************************************
:DisplayTitle

    Call :DisplayNewLine
    Call :DisplayMessage "  ****************************************************************** " 0 0
    Call :DisplayMessage "  * バッチ名：Hyper-V機能をON・OFFにするバッチ                       " 0 0
    Call :DisplayMessage "  * 処理内容：Hyper-V機能の現在の状態を表示し、Hyper-V機能をON・OFF  " 0 0
    Call :DisplayMessage "  *           のどちらを実行するかユーザーに対話して選択してもらう   " 0 0
    Call :DisplayMessage "  *           Hyper-V機能の変更後は再起動するかどうかユーザーに対    " 0 0
    Call :DisplayMessage "  *           話して実行する                                         " 0 0
    Call :DisplayMessage "  ****************************************************************** " 0 0
    Call :DisplayNewLine
    
    Exit /b

Rem ***************************************************************
Rem * 処 理 名：カレントディレクトリをbatファイルが実行されたディレ
Rem *           クトリに変更する
Rem * 引    数：なし
Rem * 処理内容：「cd」コマンドを実行してカレントディレクトリをbat
Rem *           ファイルが実行されたディレクトリに変更する
Rem ***************************************************************
:ChangeRunBatDir

    Call :DisplayMessage "★現在のディレクトリをbatファイルを実行したディレクトリに変更します……" 1 0

    Rem カレントディレクトリをbatファイルが実行したフォルダに変更
    cd /d %~dp0
    Exit /b

Rem ***************************************************************
Rem * 処 理 名：Hyper-Vの機能の現在の状態を画面に表示する
Rem * 引    数：なし
Rem * 処理内容：管理者で実行しないとHyper-Vの現在の状態を取得できな
Rem *           いのでpowershellを使用して別のコマンドプロンプトを
Rem *           管理者で実行してHyper-Vの現在の状態を表示する
Rem ***************************************************************
:DisplayHyper-VStatus

    Call :DisplayMessage "★Hyper-Vの現在の状態を表示します……" 1 0

    Rem Hyper-V機能の現在の状態をpowershellを使用して別のコマンドプロンプトを
    Rem 管理者で実行して表示する
    powershell start-process %NowHyper-VStatusBat% -verb runas

    Call :DisplayNewLine

    Exit /b

Rem ***************************************************************
Rem * 処 理 名：実行batファイルの選択
Rem * 引    数：なし
Rem * 処理内容：実行対象のbatファイルを選択します
Rem *           Hyper-VをONにするかOFFにするかのどちらかのbatファ
Rem *           イルをユーザーに対話して選択してもらう
Rem ***************************************************************
:ChooseRunBatFile

    Call :DisplayMessage "★実行batファイルを選択します……" 0 1

    Rem このbatファイルのフォルダ内のcommandと名のつくbatファイルの一覧を表示
    Set FileListCounter=0
    For /f %%i In ('Dir *command* /b') Do (

        Rem Hyper-V機能をON・OFFにするbatファイルを表示します
        Set /a FileListCounter=FileListCounter+1
        Echo !FileListCounter!：「%%i」

    )

    Rem 実行batファイルを選択
    Call :DisplayMessage "実行するbatファイルの番号を入力して下さい（1～!FileListCounter!）" 1 0
    Set /p TargetFileNo="実行するbatファイル番号の入力　＞　"

    Rem 実行batファイルの表示
    Set ChooseFileNoCounter=0
    For /f %%i In ('Dir *command* /b') Do (
    
        Set /a ChooseFileNoCounter=ChooseFileNoCounter+1
        
        Rem 実行batファイルの番号と一致した時
        If !ChooseFileNoCounter! Equ %TargetFileNo% (
        
            Rem ファイル名をセットする
            Set TargetFileName=%%i
            
        )
    )
    
    Rem 実行batファイルの表示
    Call :DisplayMessage "実行batファイル名：%TargetFileName%" 1 0

    Rem 実行batファイルが未入力の時（不正な値を入力した時）
    Rem ※もう一度、「実行batファイルの選択」処理を実行する
    If "%TargetFileName%" Equ "" (

        Call :DisplayMessage "エラー：実行batファイル名が取得できませんでした。「1～!FileListCounter!」で入力して下さい。" 1 0
        Call :ChooseRunBatFile

    ) 

    Exit /b

Rem ***************************************************************
Rem * 処 理 名：管理者でbatファイルを実行する
Rem * 引    数：なし
Rem * 処理内容：powershellを使用してbatファイルを別のコマンドプロン
Rem *           プトを管理者で実行する
Rem ***************************************************************
:RunBatFileAsAdmin

    Call :DisplayMessage "★選択されたbatファイルを実行します……" 1 0

    Rem batファイルをpowershellを使用して別のコマンドプロンプトを管理者で実行して表示する
    Rem ※管理者でないと使用できないコマンドを使用するため
    powershell start-process %TargetFileName% -verb runas

    Call :DisplayNewLine

    Exit /b

Rem ***************************************************************
Rem * 処 理 名：再起動するかどうかユーザーに対話
Rem * 引    数：なし
Rem * 処理内容：再起動するかどうかをユーザーに対話し結果を変数に
Rem *           セットする
Rem ***************************************************************
:IsReStart

    Rem 処理を続行するかユーザーに対話
    Rem ※空Enterすると落ちるため、予め適当な値をセットしておく
    Call :DisplayMessage "下記メッセージは(「Y」又は「y」)以外は処理を終了します……" 0 0
    Set RunReStartResult=KaraMojiTaiou
    Set /p RunReStartResult="再起動を実行しますか？(y/n)　＞　"

    Rem 大文字/小文字変換(Y以外は全てキャンセル扱い) 
    Set RunReStartResult=%RunReStartResult:y=Y%%

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

    If /i %RunReStartResult%==Y (

        Rem 再起動を実行して処理を終了する
        Call :DisplayMessage "☆再起動を実行して処理を終了します……" 1 1
        pause
        shutdown.exe /r /t 0        
        Exit
    )

    Rem 再起動を実行しないで終了する
    Call :DisplayMessage "☆再起動を実行しないで処理を終了します……" 1 1
    pause
    Exit
