'**************************************************************************************
'* プログラム名 ： スタートアッププログラム一括終了スクリプト                         *
'* 処理概要     ： スタートアップで起動しうるプログラムを１つずつ終了するかどうかユー *
'*                 ザーに対話で確認（メッセージボックスのはい、いいえ）。「はい」を   *
'*                 選択されたプログラムをすべて終了します。                           *
'*                 ※起動されていないプログラムを「はい」と選択しても何も起きません   *
'*                   Cliborはクリップボードの履歴が保存されないので注意！！           *
'* メモ         ：                                                                    *
'* 設定         ：                                                                    *
'**************************************************************************************

Main()

'***********************************************************************
'* 処理名   ： メイン処理                                              *
'* 引数     ： なし                                                    *
'* 処理内容 ： メイン処理                                              *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub Main()

    '*****************************************
    '* 処理続行かユーザーに対話              *
    '*****************************************
    'メッセージの表示
    Dim mContinueProcessingResult : mContinueProcessingResult = MsgBox("スタートアッププログラム一括終了処理を実行します。" & vbCrLf & "よろしいですか？", vbOKCancel, "スタートアッププログラム一括終了処理")

    'キャンセルが押された時は処理を終了
    If mContinueProcessingResult = vbCancel Then Wscript.Quit()

    '*****************************************
    '* 終了プログラムの選定処理              *
    '*****************************************
    '終了するプログラム格納Dictionary
    Dim mExitExes : Set mExitExes = WScript.CreateObject("Scripting.Dictionary")

    '終了するプログラムの追加処理
    Set mExitExes = AddExitExe(mExitExes)

    '*****************************************
    '* プログラムの一括終了処理              *
    '*****************************************
    'Dictionaryに格納されているプログラム数分繰り返す
    For Each mExeName In mExitExes.Keys

		    'プログラムの終了処理
        For Each Process in GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_Process where Name='" & mExitExes(mExeName) & "'")

            Process.terminate

        Next

    Next

End Sub

'***********************************************************************
'* 処理名   ： 終了対象プログラムの追加処理                            *
'* 引数     ： pExitExes 終了するプログラム格納Dictionary              *
'* 処理内容 ： 終了対象プログラムをDictionaryに追加する                *
'*             メッセージボックスで１つずつ確認していく                *
'* 戻り値   ： pExitExes                                               *
'***********************************************************************
Function AddExitExe(ByVal pExitExes)

    '7+ Taskbar Tweakerの終了可否
    Dim mIsRun7TaskbarTweaker : mIsRun7TaskbarTweaker = IsRunProgram("7+ Taskbar Tweaker.ex2")
    If mIsRun7TaskbarTweaker = True Then

        Dim mIsRun7TaskbarTweakerExitResult : mIsRun7TaskbarTweakerExitResult = GetSelectedUserResultForExitProgram("7+ Taskbar Tweakerを終了しますか？", "7+ Taskbar Tweaker終了可否")
        If mIsRun7TaskbarTweakerExitResult = vbYes Then

            pExitExes.Add "7+ Taskbar Tweaker" , "7+ Taskbar Tweaker.ex2"

        End If

    End If

    'AkabeiMonitorの終了可否
    Dim mIsRunAkabeiMonitor : mIsRunAkabeiMonitor = IsRunProgram("akamoni.exe")
    If mIsRunAkabeiMonitor = True Then

        Dim mAkabeiMonitorExitResult : mAkabeiMonitorExitResult = GetSelectedUserResultForExitProgram("AkabeiMonitorを終了しますか？", "AkabeiMonitor終了可否")
        If mAkabeiMonitorExitResult = vbYes Then

            pExitExes.Add "AkabeiMonitor", "akamoni.exe"

        End If

    End If

    'AutoHotKeyToolの終了可否
    Dim mIsRunAutoHotKeyTool : mIsRunAutoHotKeyTool = IsRunProgram("AutoHotKeyTool.exe")
    If mIsRunAutoHotKeyTool = True Then

        Dim mAutoHotKeyToolExitResult : mAutoHotKeyToolExitResult  = GetSelectedUserResultForExitProgram("AutoHotKeyToolを終了しますか？", "AutoHotKeyTool終了可否")
        If mAutoHotKeyToolExitResult = vbYes Then

            pExitExes.Add "AutoHotKeyTool", "AutoHotKeyTool.exe"

        End If

    End If

    'befooの終了可否
    Dim mIsRunbefoo : mIsRunbefoo = IsRunProgram("befooPtan.exe")
    If mIsRunbefoo = True Then

        Dim mbefooExitResult : mbefooExitResult  = GetSelectedUserResultForExitProgram("befooを終了しますか？", "befoo終了可否")
        If mbefooExitResult = vbYes Then

            pExitExes.Add "befoo", "befooPtan.exe"

        End If

    End If
    
    'BijinTokeiGadgetの終了可否
    Dim mIsRunBijinTokeiGadget : mIsRunBijinTokeiGadget = IsRunProgram("BijinTokeiGadget.exe")
    If mIsRunBijinTokeiGadget = True Then

        Dim mBijinTokeiGadgetExitResult : mBijinTokeiGadgetExitResult  = GetSelectedUserResultForExitProgram("BijinTokeiGadgetを終了しますか？", "BijinTokeiGadget終了可否")
        If mBijinTokeiGadgetExitResult = vbYes Then

            pExitExes.Add "BijinTokeiGadget", "BijinTokeiGadget.exe"

        End If

    End If


    'BijoLinuxGadgetの終了可否
    Dim mIsRunBijoLinuxGadget : mIsRunBijoLinuxGadget = IsRunProgram("BijoLinuxGadget.exe")
    If mIsRunBijoLinuxGadget = True Then

        Dim mBijoLinuxGadgetExitResult : mBijoLinuxGadgetExitResult = GetSelectedUserResultForExitProgram("BijoLinuxGadgetを終了しますか？", "BijoLinuxGadget終了可否")
        If mBijoLinuxGadgetExitResult = vbYes Then

            pExitExes.Add "BijoLinuxGadget", "BijoLinuxGadget.exe"

        End If

    End If



    'Chronusの終了可否
    Dim mIsRunChronus : mIsRunChronus = IsRunProgram("Chronus.exe")
    If mIsRunChronus = True Then

        Dim mChronusExitResult : mChronusExitResult  = GetSelectedUserResultForExitProgram("Chronusを終了しますか？", "Chronus終了可否")
        If mChronusExitResult = vbYes Then

            pExitExes.Add "Chronus", "Chronus.exe"

        End If

    End If

    'Cliborの終了可否
    Dim mIsRunClibor : mIsRunClibor = IsRunProgram("Clibor.exe")
    If mIsRunClibor = True Then

        Dim mCliborExitResult : mCliborExitResult  = GetSelectedUserResultForExitProgram("Cliborを終了しますか？", "Clibor終了可否")
        If mCliborExitResult = vbYes Then

            pExitExes.Add "Clibor", "Clibor.exe"

        End If

    End If

    'Evernoteの終了可否
    Dim mIsRunEvernote : mIsRunEvernote = IsRunProgram("Evernote.exe")
    If mIsRunEvernote = True Then

        Dim mEvernoteExitResult : mEvernoteExitResult = GetSelectedUserResultForExitProgram("Evernoteを終了しますか？", "Evernote終了可否")
        If mEvernoteExitResult = vbYes Then

            pExitExes.Add "Evernote"        , "Evernote.exe"
            pExitExes.Add "EvernoteClipper" , "EvernoteClipper.exe"
            pExitExes.Add "EvernotePortable", "EvernotePortable.exe"
            pExitExes.Add "EvernoteTray"    , "EvernoteTray.exe"

        End If

    End If

    'GoogleDriveの終了可否
    Dim mIsRunGoogleDrive : mIsRunGoogleDrive = IsRunProgram("googledrivesync.exe")
    If mIsRunGoogleDrive = True Then

        Dim mGoogleDriveExitResult : mGoogleDriveExitResult = GetSelectedUserResultForExitProgram("GoogleDriveを終了しますか？", "GoogleDrive終了可否")
        If mGoogleDriveExitResult = vbYes Then

            pExitExes.Add "GoogleDrive", "googledrivesync.exe"

        End If

    End If

    'Freeraserの終了可否
    Dim mIsRunFreeraser : mIsRunFreeraser = IsRunProgram("Freeraser.exe")
    If mIsRunFreeraser = True Then

        Dim mFreeraserExitResult : mFreeraserExitResult = GetSelectedUserResultForExitProgram("Freeraserを終了しますか？", "Freeraser終了可否")
        If mFreeraserExitResult = vbYes Then

            pExitExes.Add "Freeraser", "Freeraser.exe"

        End If

    End If

    'MouseGestureLの終了可否
    Dim mIsRunMouseGestureL : mIsRunMouseGestureL = IsRunProgram("MouseGestureL.exe")
    If mIsRunMouseGestureL = True Then

        Dim mMouseGestureLExitResult : mMouseGestureLExitResult  = GetSelectedUserResultForExitProgram("MouseGestureLを終了しますか？", "MouseGestureL終了可否")
        If mMouseGestureLExitResult = vbYes Then

            pExitExes.Add "MouseGestureL", "MouseGestureL.exe"

        End If

    End If

    'Orchisの終了可否
    Dim mIsRunOrchis : mIsRunOrchis = IsRunProgram("orchis.exe")
    If mIsRunOrchis = True Then

        Dim mOrchisExitResult : mOrchisExitResult  = GetSelectedUserResultForExitProgram("Orchisを終了しますか？", "Orchis終了可否")
        If mOrchisExitResult = vbYes Then

            pExitExes.Add "OrchisService", "ocobsv.exe"
            pExitExes.Add "Orchis"       , "orchis.exe"

        End If

    End If

    'RocketDockの終了可否
    Dim mIsRunRocketDock : mIsRunRocketDock = IsRunProgram("RocketDock.exe")
    If mIsRunRocketDock = True Then

        Dim mRocketDockExitResult : mRocketDockExitResult  = GetSelectedUserResultForExitProgram("RocketDockを終了しますか？", "RocketDock終了可否")
        If mRocketDockExitResult = vbYes Then

            pExitExes.Add "RocketDock", "RocketDock.exe"

        End If

    End If

    'Rss速報の終了可否
    Dim mIsRunRssNews : mIsRunRssNews = IsRunProgram("RssNews.exe")
    If mIsRunRssNews = True Then

        Dim mRssNewsExitResult : mRssNewsExitResult  = GetSelectedUserResultForExitProgram("Rss速報を終了しますか？", "Rss速報終了可否")
        If mRssNewsExitResult = vbYes Then

            pExitExes.Add "RssNews", "RssNews.exe"

        End If

    End If

    'Slackの終了可否
    Dim mIsRunSlack : mIsRunSlack = IsRunProgram("slack.exe")
    If mIsRunSlack = True Then

        Dim mSlackExitResult : mSlackExitResult = GetSelectedUserResultForExitProgram("Slackを終了しますか？", "Slack終了可否")
        If mSlackExitResult = vbYes Then

            pExitExes.Add "Slack", "slack.exe"

        End If

    End If

    'Stickiesの終了可否
    Dim mIsRunStickies : mIsRunStickies = IsRunProgram("Stickies.exe")
    If mIsRunStickies = True Then

        Dim mStickiesExitResult : mStickiesExitResult  = GetSelectedUserResultForExitProgram("Stickiesを終了しますか？", "Stickies終了可否")
        If mStickiesExitResult = vbYes Then

            pExitExes.Add "Stickies", "Stickies.exe"

        End If

    End If

    'SuperF4の終了可否
    Dim mIsRunSuperF4 : mIsRunSuperF4 = IsRunProgram("SuperF4.exe")
    If mIsRunSuperF4 = True Then

        Dim mSuperF4ExitResult : mSuperF4ExitResult = GetSelectedUserResultForExitProgram("SuperF4を終了しますか？", "SuperF4終了可否")
        If mSuperF4ExitResult = vbYes Then

            pExitExes.Add "SuperF4", "SuperF4.exe"

        End If

    End If

    'TVClockの終了可否
    Dim mIsRunTVClock : mIsRunTVClock = IsRunProgram("TVClock.exe")
    If mIsRunTVClock = True Then

        Dim mTVClockExitResult : mTVClockExitResult  = GetSelectedUserResultForExitProgram("TVClockを終了しますか？", "TVClock終了可否")
        If mTVClockExitResult = vbYes Then

            pExitExes.Add "TVClock", "TVClock.exe"

        End If

    End If

    'WheelAcceleの終了可否
    Dim mIsRunWheelAccele : mIsRunWheelAccele = IsRunProgram("WheelAccele.exe")
    If mIsRunWheelAccele = True Then

        Dim mWheelAcceleExitResult : mWheelAcceleExitResult  = GetSelectedUserResultForExitProgram("WheelAcceleを終了しますか？", "WheelAccele終了可否")
        If mWheelAcceleExitResult  = vbYes Then

            pExitExes.Add "WheelAccele" , "WheelAccele.exe"

        End If

    End If

    'X-Finderの終了可否
    Dim mIsRunXFinder : mIsRunXFinder = IsRunProgram("XF.exe")
    If mIsRunXFinder = True Then

        Dim mXFinderExitResult : mXFinderExitResult = GetSelectedUserResultForExitProgram("X-Finderを終了しますか？", "X-Finder終了可否")
        If mXFinderExitResult = vbYes Then

            pExitExes.Add "X-Finder32" , "XF.exe"
            pExitExes.Add "X-Finder64" , "xf64.exe"

        End If

    End If

    '机上予報の終了可否
    Dim mIsRunWeather : mIsRunWeather = IsRunProgram("Weather64.exe")
    If mIsRunWeather = True Then

        Dim mRunWeatherExitResult : mRunWeatherExitResult  = GetSelectedUserResultForExitProgram("机上予報を終了しますか？", "机上予報終了可否")
        If mRunWeatherExitResult = vbYes Then

            pExitExes.Add "Weather", "Weather64.exe"

        End If

    End If

    '棒読みちゃんの終了可否
    Dim mIsRunBouyomiChan : mIsRunBouyomiChan = IsRunProgram("BouyomiChan.exe")
    If mIsRunBouyomiChan = True Then

        Dim mBouyomiChanExitResult : mBouyomiChanExitResult  = GetSelectedUserResultForExitProgram("棒読みちゃんを終了しますか？", "棒読みちゃん終了可否")
        If mBouyomiChanExitResult = vbYes Then

            pExitExes.Add "BouyomiChan", "BouyomiChan.exe"

        End If

    End If

    Set AddExitExe = pExitExes

End Function

'***********************************************************************
'* 処理名   ： プログラム終了可否                                      *
'* 引数     ： pMsgBoxTitle  メッセージボックスのタイトル              *
'*             pMsgBoxDetail メッセージボックスの内容                  *
'* 処理内容 ： メッセージボックスを表示ユーザーにプログラムを終了する  *
'*             かどうか対話し結果を返す                                *
'* 戻り値   ： メッセージボックスの結果 / vbYes、vbNo                  *
'***********************************************************************
Function GetSelectedUserResultForExitProgram(ByVal pMsgBoxTitle,ByVal pMsgBoxDetail)

    'ユーザーにプログラムの終了可否
    Dim mMsgBoxResult : mMsgBoxResult = MsgBox(pMsgBoxTitle, vbYesNo, pMsgBoxDetail)

    'ユーザーが選択した結果をセット
    GetSelectedUserResultForExitProgram = mMsgBoxResult

End Function

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
