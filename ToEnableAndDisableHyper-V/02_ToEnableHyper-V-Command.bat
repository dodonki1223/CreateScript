@Echo off
Rem ***************************************************************
Rem * バッチ名：Hyper-V機能をONにするコマンドバッチ
Rem * 処理内容：Hyper-V機能をコマンドラインからONにします
Rem ***************************************************************

Rem ---------------------- メイン処理部分 ---------------------- 
Rem タイトル表示処理
Call :DisplayTitle

Rem Hyper-Vの機能をONにする
Call :RunEnableHyper-V

Rem 終了処理
Call :EndProcess

Rem ---------------------- ラベル処理部分 ---------------------- 

Rem ***************************************************************
Rem * 処 理 名：タイトルの表示
Rem * 引    数：なし
Rem * 処理内容：Hyper-V機能をONにするコマンドバッチのタイトルを表
Rem *           示する
Rem ***************************************************************
:DisplayTitle

    Call :DisplayNewLine
    Call :DisplayMessage "*************************************************************** " 0 0
    Call :DisplayMessage "* バッチ名：Hyper-V機能をONにするコマンドバッチ                 " 0 0
    Call :DisplayMessage "* 処理内容：Hyper-V機能をコマンドラインからONにします           " 0 0
    Call :DisplayMessage "*************************************************************** " 0 0
    Call :DisplayNewLine
    
    Exit /b

Rem ***************************************************************
Rem * 処 理 名：Hyper-Vの機能をONにするコマンドを実行
Rem * 引    数：なし
Rem * 処理内容：「BCDEdit」コマンドを使用してHyper-V機能をONにする
Rem *           コマンドを実行する
Rem ***************************************************************
:RunEnableHyper-V

    Rem 実行するコマンドを画面に表示、ユーザーに処理の続行を対話
    Call :DisplayMessage "下記のコマンドを実行してHyper-V機能をONにします " 0 0
    Call :DisplayMessage "bcdedit /set hypervisorlaunchtype auto          " 0 1
    Pause

    Rem 「BCDEdit」コマンドを使用してHyper-Vの機能をONする
    Call :DisplayNewLine
    bcdedit /set hypervisorlaunchtype auto
    Call :DisplayNewLine

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

    Call :DisplayMessage "Hyper-V機能をONしました！！          " 0 0
    Call :DisplayMessage "再起動をすることで設定が反映されます " 0 1
    Pause

    Rem 処理の終了
    Exit
        