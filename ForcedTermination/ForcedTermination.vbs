'**************************************************************************************
'* プログラム名 ： EXE強制終了スクリプト                                              *
'* 処理概要     ： 終了したいEXE名をユーザーに入力させ、入力したEXE名に一致するEXEを  *
'*                 強制終了します                                                     *
'* メモ         ：                                                                    *
'* 設定         ：                                                                    *
'**************************************************************************************

'メッセージボックス・インプットボックスで使用する文言
Dim MsgInputBoxTitle   : MsgInputBoxTitle   = "EXE強制終了スクリプト"
Dim MsgInputExeName    : MsgInputExeName    = "終了したいEXE名を入力してください。" & VbCrLf & VbCrLf & "※タスクマネージャーに表示されているEXE名を入力して下さい"
Dim MsgNotExistsExe    : MsgNotExistsExe    = "入力されたEXE名のプログラムが起動していません。" & VbCrLf & "スクリプトを終了します。"
Dim MsgIsRunEndProcess : MsgIsRunEndProcess = "を終了します。" & VbCrLf & "よろしいですか？"
Dim MsgRanEndProcess   : MsgRanEndProcess   = "を強制終了しました。"

Main()

'***********************************************************************
'* 処理名   ： メイン処理                                              *
'* 引数     ： なし                                                    *
'* 処理内容 ： メイン処理                                              *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub Main()

    '*********************************
    '* 終了するExe名を取得           *
    '*********************************
    'ユーザーに対話し終了するExe名を取得
    Dim mExeName : mExeName = InputBox(MsgInputExeName, MsgInputBoxTitle)
    
    '取得したExe名に「.exe」が含まれてなかったら、末尾に「.exe」を追加
    If InStr(mExeName, ".exe") <= 0 then mExeName = mExeName & ".exe"

    '*********************************
    '* 入力されたEXE名の存在チェック *
    '*********************************
    '入力されたEXEが存在するかどうか取得
    Dim mIsRunExe : mIsRunExe = IsRunProgram(mExeName)
    
    '入力されたEXEが存在しない時
    If mIsRunExe = False Then
    
        '「EXEが存在しません」メッセージを表示しスクリプトを終了
        MsgBox MsgNotExistsExe, vbOkOnly, MsgInputBoxTitle
        Wscript.Quit()
    
    End If

    '*********************************
    '* EXEの終了処理を行うか確認     *
    '*********************************
    'EXEの終了処理を行うかユーザーに対話
    Dim mIsRunEndProcess : mIsRunEndProcess = MsgBox("「" & mExeName & "」" & MsgIsRunEndProcess, vbOKCancel, MsgInputBoxTitle)
    
    'キャンセルが押された時は処理を終了
    If mIsRunEndProcess = vbCancel Then Wscript.Quit()

    '*********************************
    '* Exeの終了処理                 *
    '*********************************
    For Each Process in GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_Process where Name='" & mExeName & "'")

        'プログラムの強制終了処理
        Process.terminate
        
        '「プログラム名」を強制終了しましたメッセージを表示
        MsgBox "「" & mExeName & "」" & MsgRanEndProcess, vbOkOnly, MsgInputBoxTitle

    Next

End Sub

'***********************************************************************
'* 処理名   ： プログラム起動状態を取得                                *
'* 引数     ： pProgramExe  対象プログラム（○○○.exe形式）           *
'* 処理内容 ： 対象プログラムが起動中かどうかを取得する                *
'* 戻り値   ： 対象プログラムが起動中有無 / True、False                *
'***********************************************************************
Function IsRunProgram(ByVal pProgramExe)

    '起動中有無、デフォルト値設定
    IsRunProgram = False

    '対象プログラムを取得
    Set mPrograms = GetObject("winmgmts:").ExecQuery("Select * from Win32_Process where Name='" & pProgramExe & "'")

    '対象プログラムが取得出来たら起動中とする（１件でもあったら）
    For Each Program in mPrograms

        IsRunProgram = True
        Exit For

    Next

End Function