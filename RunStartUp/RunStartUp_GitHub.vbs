'**************************************************************************************
'* プログラム名 ： スタートアップ処理スクリプト                                       *
'* 処理概要     ： スタートアップ時に実行するスクリプト。実行されたドライブからOrchis *
'*                 で使用するショートカットファイルのリンク先を作成し直す。RSS速報の  *
'*                  棒読みちゃんのパス設定を実行されたドライブ用に書き換える。スター  *
'*                 トアップ時に実行されて欲しいプログラムを一括で実行する             *
'* メモ         ： このファイルをショートカットにしてコマンドライン引数を指定すること *
'*                 ★使用例★                                                         *
'*                   C:\Tools\CreateScript\RunStartUp\RunStartUp.vbs "Company"        *
'*                   C:\Tools\CreateScript\RunStartUp\RunStartUp.vbs "House"          *
'*                   C:\Tools\CreateScript\RunStartUp\RunStartUp.vbs "NotePC"         *
'*                 ※実行する環境によりコマンドライン引数を変更する事                 *
'*                 URLファイルの作成方法                                              *
'*                   ファイル名を「○○○.url」形式にしショートカット先にURLを指定    *
'* 設定         ： このスクリプトのデフォルト設定はUSBで実行されます                  *
'**************************************************************************************

'--------------------------------------
' 設定
'--------------------------------------
'※実行区分「Company：会社、House：家、USB：USB、NotePC：ノートパソコン」
'  デフォルトはUSBです
Dim runKbn : runKbn = "USB"

'コマンドライン引数を取得し実行区分にセットする
'※コマンドライン引数を取得出来たときだけセットする
If WScript.Arguments.Count > 0 Then

    runKbn = WScript.Arguments(0)

End If

'--------------------------------------
' 変数宣言・インスタンス作成
'--------------------------------------
Dim objShell   : Set objShell   = WScript.CreateObject("WScript.Shell")              'WScript.Shellオブジェクト
Dim objAppli   : Set objAppli   = WScript.CreateObject("Shell.Application")          'WScript.Applicationオブジェクト
Dim objFso     : set objFso     = WScript.CreateObject("Scripting.FileSystemObject") 'FileSystemObject
Dim fileInfo   : Set fileInfo   = WScript.CreateObject("Scripting.Dictionary")       'ファイル情報格納Dictionary
Dim runFile    : Set runFile    = WScript.CreateObject("Scripting.Dictionary")       '実行EXE格納Dictionary
Dim orchisDirectory                                                                  'Orchis実行ディレクトリ

Main()

'***********************************************************************
'* 処理名   ： メイン処理                                              *
'* 引数     ： なし                                                    *
'* 処理内容 ： メイン処理                                              *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub Main()

    '--------------------------------------
    ' 処理続行かユーザーに対話
    '--------------------------------------
    'メッセージボックスの表示
    Dim msgResult : msgResult = MsgBox("スタートアップ処理を実行します。" & vbCrLf & "よろしいですか？", vbOKCancel, "スタートアップ処理")

    'キャンセルが押された時は処理を終了
    If msgResult = vbCancel Then Wscript.Quit()

    '--------------------------------------
    ' 実行ドライブを取得
    '--------------------------------------
    Dim runDrive : runDrive = objFSo.GetDriveName(WScript.ScriptFullName)

    '--------------------------------------
    ' RSS速報の棒読みちゃんのパス設定変更
    '--------------------------------------
    Dim bouyomiC           : bouyomiC           = "C:\Tools\BouyomiChan\RemoteTalk\RemoteTalk.exe"
    Dim bouyomiD           : bouyomiD           = "D:\Tools\BouyomiChan\RemoteTalk\RemoteTalk.exe"
    Dim bouyomiE           : bouyomiE           = "E:\Tools\BouyomiChan\RemoteTalk\RemoteTalk.exe"
    Dim bouyomiF           : bouyomiF           = "F:\Tools\BouyomiChan\RemoteTalk\RemoteTalk.exe"
    Dim bouyomiG           : bouyomiG           = "G:\Tools\BouyomiChan\RemoteTalk\RemoteTalk.exe"
    Dim bouyomiH           : bouyomiH           = "H:\Tools\BouyomiChan\RemoteTalk\RemoteTalk.exe"
    Dim bouyomiAfterPath   : bouyomiAfterPath   = runDrive & "\Tools\BouyomiChan\RemoteTalk\RemoteTalk.exe"
    Dim rssNewsSettingPath : rssNewsSettingPath = runDrive & "\Tools\RssNews\setting.xml"

    Call ReplaceTextDetail(rssNewsSettingPath,bouyomiC,bouyomiAfterPath,"UTF-8")
    Call ReplaceTextDetail(rssNewsSettingPath,bouyomiD,bouyomiAfterPath,"UTF-8")
    Call ReplaceTextDetail(rssNewsSettingPath,bouyomiE,bouyomiAfterPath,"UTF-8")
    Call ReplaceTextDetail(rssNewsSettingPath,bouyomiF,bouyomiAfterPath,"UTF-8")
    Call ReplaceTextDetail(rssNewsSettingPath,bouyomiG,bouyomiAfterPath,"UTF-8")
    Call ReplaceTextDetail(rssNewsSettingPath,bouyomiH,bouyomiAfterPath,"UTF-8")

    '--------------------------------------
    ' ファイルの実行処理
    '--------------------------------------
    '実行対象ファイルを追加
    Call AddRunFile(runFile,runDrive,orchisDirectory)

    'ファイルの一括実行処理
    For Each key In runFile.Keys

        'EXEの実行処理
        objShell.Run runFile(key)

        'MouseGestureLの時は起動後１０秒間待つ(完全に起動するまで待つ)
        '※MouseGestureLとAutoHotKeyToolの相性の問題、MouseGestureLの起動後に立ち
        '  上げないとAutoHotKeyToolで設定したショートカットが効かなくなるため
        If(key = "MouseGestureL") Then WScript.Sleep(10000)

    Next

    '--------------------------------------
    ' ショートカット一括作成
    '--------------------------------------
    'ショートカットを作成するファイル情報を追加
    Call AddShortCutFile(fileInfo,runDrive,orchisDirectory)

    'ショートカット作成処理
    For Each key In fileInfo.Keys

        '作成情報を切り分ける（①ショートカットパス、②出力フォルダ、③コマンドライン引数、④アイコン情報）
        Dim aryFileInfo : aryFileInfo = Split(fileInfo(key),"|")

        'lnkファイルのファイル名からショートカットを作成するディレクトリを取得
        Dim fileName : fileName = key
        Dim path : path = aryFileInfo(1) & fileName

        'ショートカット作成元のフォルダが無かった場合はフォルダを作成する
        '※Bookmarkフォルダ作成用に記述
        CreateNotExistFolder(aryFileInfo(0))

        '作成先ディレクトリのフォルダが無かった場合はフォルダを作成する
        CreateNotExistFolder(path)

        'ショートカットオブジェクトを作成し出力先パス、コマンドライン引数、アイコンを指定
        Set shortCut = objShell.CreateShortcut(path)                               'ショートカットオブジェクトを作成
        shortCut.TargetPath = aryFileInfo(0)                                       'ショートカット先
        If UBound(aryFileInfo) > 1 Then shortCut.Arguments        = aryFileInfo(2) 'コマンドライン引数設定
        If UBound(aryFileInfo) > 2 Then shortCut.IconLocation     = aryFileInfo(3) 'アイコン情報を設定
        If UBound(aryFileInfo) > 3 Then shortCut.WorkingDirectory = aryFileInfo(4) '作業フォルダを設定

        'ショートカットを作成
        shortCut.Save

    Next

    '--------------------------------------
    ' オブジェクト破棄処理
    '--------------------------------------
    Set objShell   = Nothing
    Set objAppli   = Nothing
    Set objFso     = Nothing
    Set fileInfo   = Nothing
    Set runFile    = Nothing

End Sub

'***********************************************************************
'* 処理名   ： テキストファイル内の文字列変換                          *
'* 引数     ： pTextFilePath 対象テキストファイルのフルパス            *
'*             pBeforeText   変換元文字列                              *
'*             pAfterText    変換後文字列                              *
'*             pStringCode   文字コード                                *
'* 処理内容 ： 指定した文字コードのファイルを開き、文字列をの変換を行  *
'*             う                                                      *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub ReplaceTextDetail(ByVal pTextFilePath,ByVal pBeforeText,ByVal pAfterText,ByVal pStringCode)

    '--------------------------------------
    ' 変数宣言・インスタンス作成
    '--------------------------------------
    Dim objFso    : Set objFso    = WScript.CreateObject("Scripting.FileSystemObject") 'FileSystemObject
    Dim objStream : Set objStream = CreateObject("ADODB.Stream")                       'Streamオブジェクト（データのバイナリまたはテキストストリームを表します）

    '--------------------------------------
    ' テキストファイルを開く
    '--------------------------------------
    'オープンするファイルの情報をセットする
    objStream.type = 2              'テキストデータ
    objStream.Charset = pStringCode '文字コード

    'Straamオブジェクトを開く
    objStream.Open

    '指定されたファイルを開く -1:テキスト内容のすべてを読み込む
    objStream.LoadFromFile pTextFilePath
    Dim readFileDetail : readFileDetail = objStream.ReadText(-1)

    '--------------------------------------
    ' ファイル内容の置換処理
    '--------------------------------------
    readFileDetail = replace(readFileDetail,pBeforeText,pAfterText)

    '--------------------------------------
    ' 置換内容の書き込み処理
    '--------------------------------------
    '書き込み位置をトップにし書き込みを行う
    objStream.Position = 0
    objStream.WriteText readFileDetail , 0

    'ファイルの上書き処理
    objStream.SaveToFile pTextFilePath, 2

    'Streamオブジェクトを閉じる
    objStream.Close

    '--------------------------------------
    ' オブジェクト破棄処理
    '--------------------------------------
    Set objFso    = Nothing
    Set objStream = Nothing

End Sub

'***********************************************************************
'* 処理名   ： 実行対象ファイルの追加処理                              *
'* 引数     ： pRunFile         実行対象ファイル格納Dictionary         *
'*             pRunDrive        実行ドライブパス                       *
'*             pOrchisDirectory Orchisの実行ドライブ格納変数           *
'* 処理内容 ： 実行対象のファイル情報をDictionaryに追加する            *
'*             一部のプログラムはユーザーに対話して追加するか問う      *
'* 戻り値   ： pRunFile                                                *
'*             pOrchisDirectory                                        *
'***********************************************************************
Function AddRunFile(ByRef pRunFile,ByVal pRunDrive,ByRef pOrchisDirectory)

    '--------------------------------------
    ' ファイル情報を設定していく
    ' ※キー：ファイル名、項目：ファイルパス
    '--------------------------------------
    'ファイラー起動可否 いいえが押された時はX-Finderを起動しない(起動ファイル格納Dictionaryに追加しない)
    Dim msgRunFilerResult : msgRunFilerResult = MsgBox("ファイラーを起動しますか？", vbYesNo, "ファイラー起動可否")
    If msgRunFilerResult = vbYes Then

        pRunFile.Add "X-Finder"           , pRunDrive & "\Tools\X-Finder\xf64.exe"

    End If

    'マウス存在可否 いいえが押された時はMouseGestureLを起動しない(起動ファイル格納Dictionaryに追加しない)
    Dim msgMouseExistResult : msgMouseExistResult = MsgBox("お使いのパソコンにマウスはありますか？", vbYesNo, "マウス存在可否")
    If msgMouseExistResult = vbYes Then

        pRunFile.Add "WheelAccele"        , pRunDrive & "\Tools\AutoHotKey\Tools\WheelAccele\WheelAccele.exe"
        pRunFile.Add "MouseGestureL"      , pRunDrive & "\Tools\MouseGestureL\MouseGestureL.exe"

    End If

    'ネットワークが使えるかどうか
    Dim msgIsUseNetworkResult : msgIsUseNetworkResult = MsgBox("ネットワークが使える環境ですか？", vbYesNo, "ネットワーク利用可能可否")
    If msgIsUseNetworkResult = vbYes Then

        '----------------------------
        ' ネットワークが使用可の時
        '----------------------------
        'ブラウザー起動可否 いいえが押された時はGoogle Chromeを起動しない(起動ファイル格納Dictionaryに追加しない)
        Dim msgRunBrowserResult : msgRunBrowserResult = MsgBox("ブラウザーを起動しますか？", vbYesNo, "ブラウザー起動可否")
        If msgRunBrowserResult = vbYes Then

            Select Case runKbn

                Case "Company", "House", "NotePC"

                    pRunFile.Add "GoogleChrome"       , """" & pRunDrive & "\Program Files (x86)\Google\Chrome\Application\chrome.exe"""

                Case "USB"

                    pRunFile.Add "GoogleChrome"       , pRunDrive & "\Tools\GoogleChromePortable\GoogleChromePortable.exe"

            End Select

        End If

        'メーラー起動可否 いいえが押された時はThunderBirdを起動しない(起動ファイル格納Dictionaryに追加しない)
        Dim msgRunMailerResult : msgRunMailerResult = MsgBox("メーラーを起動しますか？", vbYesNo, "メーラー起動可否")
        If msgRunMailerResult = vbYes Then

            Select Case runKbn

                Case "Company"

                    pRunFile.Add "ThunderBird"        , """" & pRunDrive & "\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe"""

                Case "House", "USB", "NotePC"

                    pRunFile.Add "ThunderBird"        , pRunDrive & "\Tools\ThunderbirdPortable\ThunderbirdPortable.exe"

            End Select

        End If

        Select Case runKbn

            Case "House"

                '棒読みちゃん起動可否 いいえが押された時は棒読みちゃんを起動しない(起動ファイル格納Dictionaryに追加しない)
                Dim msgRunBouyoumiChanResult : msgRunBouyoumiChanResult = MsgBox("棒読みちゃんを起動しますか？", vbYesNo, "棒読みちゃん起動可否")
                If msgRunBouyoumiChanResult = vbYes Then

                      pRunFile.Add "BouyomiChan"       , pRunDrive & "\Tools\BouyomiChan\BouyomiChan.exe"

                End If
                
                pRunFile.Add "befoo"              , pRunDrive & "\Tools\befoo\befooPtan.exe"
                pRunFile.Add "RssNews"            , pRunDrive & "\Tools\RssNews\RssNews.exe"
                
            Case "Company"

                pRunFile.Add "befoo"              , pRunDrive & "\Tools\befoo\befooPtan.exe"
                pRunFile.Add "RssNews"            , pRunDrive & "\Tools\RssNews\RssNews.exe"

        End Select

    End If

    pRunFile.Add "Clibor"             , pRunDrive & "\Tools\clibor\Clibor.exe"
    pRunFile.Add "AutoHotKeyTool"     , pRunDrive & "\Tools\AutoHotKey\AutoHotKeyTool.exe"
    pRunFile.Add "AkabeiMonitor"      , pRunDrive & "\Tools\AkabeiMonitor\akamoni.exe"
    pRunFile.Add "Freeraser"          , pRunDrive & "\Tools\Freeraser\Freeraser.exe"
    pRunFile.Add "7+ Taskbar Tweaker" , """" & pRunDrive & "\Tools\7+ Taskbar Tweaker\7+ Taskbar Tweaker.exe"""

    Select Case runKbn
    
        Case "Company"

            pRunFile.Add "BijinTokeiGadget"   , pRunDrive & "\Tools\BijinTokeiGadget\BijinTokeiGadget.exe"
            pRunFile.Add "BijoLinuxGadget"    , pRunDrive & "\Tools\BijoLinuxGadget\BijoLinuxGadget.exe"
            pRunFile.Add "Evernote"           , pRunDrive & "\Tools\EvernotePortable\EvernotePortable.exe"
            pRunFile.Add "Stickies"           , pRunDrive & "\Tools\StickiesPortable\StickiesPortable.exe"
            pRunFile.Add "TVClock"            , pRunDrive & "\Tools\tvclock\TVClock.exe"
            pRunFile.Add "RocketDock"         , pRunDrive & "\Tools\RocketDock\RocketDock.exe"
            pRunFile.Add "Chronus"            , pRunDrive & "\Tools\Chronus\"
            pRunFile.Add "Slack"              , """" & "%UserProfile%\AppData\Local\slack\slack.exe"""
            pRunFile.Add "SuperF4"            , pRunDrive & "\Tools\SuperF4\SuperF4.exe"
            pRunFile.Add "GoogleDrive"        , """" & pRunDrive & "\Program Files\Google\Drive\googledrivesync.exe"""
            pRunFile.Add "机上予報"           , pRunDrive & "\Tools\Weather\Weather64.exe"

        Case "House"

            pRunFile.Add "BijinTokeiGadget"   , pRunDrive & "\Tools\BijinTokeiGadget\BijinTokeiGadget.exe"
            pRunFile.Add "BijoLinuxGadget"    , pRunDrive & "\Tools\BijoLinuxGadget\BijoLinuxGadget.exe"
            pRunFile.Add "Evernote"           , pRunDrive & "\Tools\EvernotePortable\EvernotePortable.exe"
            pRunFile.Add "Stickies"           , pRunDrive & "\Tools\StickiesPortable\StickiesPortable.exe"
            pRunFile.Add "TVClock"            , pRunDrive & "\Tools\tvclock\TVClock.exe"
            pRunFile.Add "RocketDock"         , pRunDrive & "\Tools\RocketDock\RocketDock.exe"
            pRunFile.Add "Chronus"            , pRunDrive & "\Tools\Chronus\"
            pRunFile.Add "Slack"              , """" & "%UserProfile%\AppData\Local\slack\slack.exe"""
            pRunFile.Add "SuperF4"            , pRunDrive & "\Tools\SuperF4\SuperF4.exe"
            pRunFile.Add "GoogleDrive"        , """" & pRunDrive & "\Program Files\Google\Drive\googledrivesync.exe"""
            pRunFile.Add "机上予報"           , pRunDrive & "\Tools\Weather\Weather64.exe"

            
        Case "NotePC"

            pRunFile.Add "TVClock"            , pRunDrive & "\Tools\tvclock\TVClock.exe"

    End Select

    '実行ドライブ文字列を取得
    Dim driveStr : driveStr = Left(pRunDrive, 1)

    'ドライブごと起動するOrchisを変更する
    Select Case driveStr

        Case "C"

            Select Case runKbn

                Case "Company", "House", "NotePC"

                    pOrchisDirectory = """" & pRunDrive & "\Program Files\Orchis\orchis.exe""" 'インストール版

                Case "USB"

                    orchisDirectory = pRunDrive & "\Tools\orchisC\orchis-p.exe"               'ポータブル版

            End Select

        Case "D"

             pOrchisDirectory = pRunDrive & "\Tools\orchisD\orchis-p.exe"

        Case "E"

            pOrchisDirectory = pRunDrive & "\Tools\orchisE\orchis-p.exe"

        Case "F"

            pOrchisDirectory = pRunDrive & "\Tools\orchisF\orchis-p.exe"

        Case "G"

            pOrchisDirectory = pRunDrive & "\Tools\orchisG\orchis-p.exe"

        Case "H"

            pOrchisDirectory = pRunDrive & "\Tools\orchisH\orchis-p.exe"

        Case Else

    End Select
    pRunFile.Add "Orchis"             , pOrchisDirectory

End Function

'***********************************************************************
'* 処理名   ： ショートカット作成ファイルの追加処理                    *
'* 引数     ： pFileInfo        ショートカット作成情報格納Dictionary   *
'*             pRunDrive        実行ドライブパス                       *
'*             pOrchisDirectory Orchisの実行ドライブ格納変数           *
'* 処理内容 ： ショートカットを作成するファイル情報をDictionaryに格納  *
'*             する                                                    *
'* 戻り値   ： pFileInfo                                               *
'***********************************************************************
Function AddShortCutFile(ByRef pFileInfo,ByVal pRunDrive,ByVal pOrchisDirectory)

    '-----------------------------------------------------
    ' ショートカットを作成するファイル情報をセットしていく
    ' ※キー：ファイル名、作成情報：ファイルパス|出力先フォルダ|コマンドライン引数|アイコン情報
    '   ランチャー用に実行されたドライブでショートカットのパスを作りなおす
    '-----------------------------------------------------
    'ファイル名                                                'ショートカット先                                                                            'ファイルの出力先                                                  'コマンドライン引数                                 'アイコンファイル                                                        '作業フォルダ                              

    '★StartUp★
    pFileInfo.Add "7+ Taskbar Tweaker.lnk"                    , """" & pRunDrive & "\Tools\7+ Taskbar Tweaker\7+ Taskbar Tweaker.exe"""               & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "AkabeiMonitor.lnk"                         , pRunDrive & "\Tools\AkabeiMonitor\akamoni.exe"                                        & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "AutoHotKeyTool.lnk"                        , pRunDrive & "\Tools\AutoHotKey\AutoHotKeyTool.exe"                                    & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "befoo.lnk"                                 , pRunDrive & "\Tools\befoo\befooPtan.exe"                                              & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "BijinTokeiGadget.lnk"                      , pRunDrive & "\Tools\BijinTokeiGadget\BijinTokeiGadget.exe"                            & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "BijoLinuxGadget.lnk"                       , pRunDrive & "\Tools\BijoLinuxGadget\BijoLinuxGadget.exe"                              & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "Clibor.lnk"                                , pRunDrive & "\Tools\clibor\Clibor.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "Evernote.lnk"                              , pRunDrive & "\Tools\EvernotePortable\EvernotePortable.exe"                            & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "Freeraser.lnk"                             , pRunDrive & "\Tools\Freeraser\Freeraser.exe"                                          & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "MouseGestureL.lnk"                         , pRunDrive & "\Tools\MouseGestureL\MouseGestureL.exe"                                  & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "Orchis.lnk"                                , pOrchisDirectory                                                                      & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"                            & "|" & ""                                 & "|" & pRunDrive & "\Program Files\Orchis\orchis.exe"
    pFileInfo.Add "RSS速報.lnk"                               , pRunDrive & "\Tools\RssNews\RssNews.exe"                                              & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "Stickies.lnk"                              , pRunDrive & "\Tools\StickiesPortable\StickiesPortable.exe"                            & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "SuperF4.lnk"                               , pRunDrive & "\Tools\SuperF4\SuperF4.exe"                                              & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "WheelAccele.lnk"                           , pRunDrive & "\Tools\AutoHotKey\Tools\WheelAccele\WheelAccele.exe"                     & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "机上予報.lnk"                              , pRunDrive & "\Tools\Weather\Weather64.exe"                                            & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"

    Select Case runKbn

        Case "Company"

            pFileInfo.Add "BijinTokeiGadget_Other.lnk"                , pRunDrive & "\Tools\BijinTokeiGadget_Other\BijinTokeiGadget.exe"                      & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "Chronus.lnk"                               , pRunDrive & "\Tools\Chronus\"                                                         & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "GoogleDrive.lnk"                           , """" & pRunDrive & "\Program Files\Google\Drive\googledrivesync.exe"""                & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "RocketDock.lnk"                            , pRunDrive & "\Tools\RocketDock\RocketDock.exe"                                        & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "Slack.lnk"                                 , """" & "%UserProfile%\AppData\Local\slack\slack.exe"""                                & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "TVClock.lnk"                               , pRunDrive & "\Tools\tvclock\TVClock.exe"                                              & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"

        Case "House"

            pFileInfo.Add "BijinTokeiGadget_Other.lnk"                , pRunDrive & "\Tools\BijinTokeiGadget_Other\BijinTokeiGadget.exe"                      & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "Chronus.lnk"                               , pRunDrive & "\Tools\Chronus\"                                                         & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "GoogleDrive.lnk"                           , """" & pRunDrive & "\Program Files\Google\Drive\googledrivesync.exe"""                & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "RocketDock.lnk"                            , pRunDrive & "\Tools\RocketDock\RocketDock.exe"                                        & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "Slack.lnk"                                 , """" & "%UserProfile%\AppData\Local\slack\slack.exe"""                                & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "TVClock.lnk"                               , pRunDrive & "\Tools\tvclock\TVClock.exe"                                              & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "棒読みちゃん.lnk"                          , pRunDrive & "\Tools\BouyomiChan\BouyomiChan.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"

        Case "NotePC"

            pFileInfo.Add "Chronus.lnk"                               , pRunDrive & "\Tools\Chronus\"                                                         & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "GoogleDrive.lnk"                           , """" & pRunDrive & "\Program Files\Google\Drive\googledrivesync.exe"""                & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "RocketDock.lnk"                            , pRunDrive & "\Tools\RocketDock\RocketDock.exe"                                        & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "Slack.lnk"                                 , """" & "%UserProfile%\AppData\Local\slack\slack.exe"""                                & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "TVClock.lnk"                               , pRunDrive & "\Tools\tvclock\TVClock.exe"                                              & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "棒読みちゃん.lnk"                          , pRunDrive & "\Tools\BouyomiChan\BouyomiChan.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"

        Case "USB"

            pFileInfo.Add "棒読みちゃん.lnk"                          , pRunDrive & "\Tools\BouyomiChan\BouyomiChan.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"

    End Select

    '★OftenUse★
    pFileInfo.Add "Atom.lnk"                                  , pRunDrive & "\Tools\Atom\atom.exe"                                                    & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
    pFileInfo.Add "EmEditorPortable.lnk"                      , pRunDrive & "\Tools\EmEditor\EmEditor.exe"                                            & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
    pFileInfo.Add "FolderFileList.lnk"                        , pRunDrive & "\Tools\FolderFileList\FolderFileList.exe"                                & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
    pFileInfo.Add "FolderFileListDebug.lnk"                   , pRunDrive & "\Tools\FolderFileList\FolderFileList_Debug.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
    pFileInfo.Add "00_ImageForClipboard.lnk"                  , pRunDrive & "\Tools\ImageForClipboard\ImageForClipboard.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\ImageForClipboard\"
    pFileInfo.Add "01_ヘルプ.lnk"                             , pRunDrive & "\Tools\ImageForClipboard\ImageForClipboard.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\ImageForClipboard\"         & "|" & "/?"
    pFileInfo.Add "02_クリップボード内の画像を表示.lnk"       , pRunDrive & "\Tools\ImageForClipboard\ImageForClipboard.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\ImageForClipboard\"         & "|" & "/AutoClose 5 /ImageSize 40"
    pFileInfo.Add "03_クリップボード内の画像を保存.lnk"       , pRunDrive & "\Tools\ImageForClipboard\ImageForClipboard.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\ImageForClipboard\"         & "|" & "/AutoClose 2 /ImageSize 40 /AutoSave %UserProfile%\Downloads\ /Extension png"
    pFileInfo.Add "ReduceMemory.lnk"                          , pRunDrive & "\Tools\ReduceMemory\ReduceMemory.exe"                                    & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
    pFileInfo.Add "Skype.lnk"                                 , pRunDrive & "\Tools\SkypePortable\SkypePortable.exe"                                  & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
    pFileInfo.Add "Skype終了.lnk"                             , pRunDrive & "\Tools\SkypePortable\SkypePortable.exe"                                  & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"                           & "|" & "/shutdown"
    pFileInfo.Add "TeamViewer.lnk"                            , pRunDrive & "\Tools\TeamViewerPortable\TeamViewerPortable.exe"                        & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
    pFileInfo.Add "Visual Studio Code.lnk"                    , pRunDrive & "\Tools\VSCode\Code.exe"                                                  & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
    pFileInfo.Add "X-Finder.lnk"                              , pRunDrive & "\Tools\X-Finder\xf64.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"

    Select Case runKbn

        Case "Company"

            pFileInfo.Add "GoogleChrome.lnk"                          , "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"                            & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
            pFileInfo.Add "Thunderbird.lnk"                           , "%ProgramFiles(x86)%\Mozilla Thunderbird\thunderbird.exe"                             & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
            pFileInfo.Add "EmEditor.lnk"                              , """" & pRunDrive & "\Program Files\EmEditor\EmEditor.exe"""                           & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"                   & "|" & ""                                 & "|" & """" & pRunDrive & "\Program Files\EmEditor\EmEditor.exe"""

        Case "House", "NotePC"

            pFileInfo.Add "GoogleChrome.lnk"                          , "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"                            & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
            pFileInfo.Add "LINE.lnk"                                  , "%UserProfile%\AppData\Local\LINE\bin\LineLauncher.exe"                               & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
            pFileInfo.Add "Thunderbird.lnk"                           , pRunDrive & "\Tools\ThunderbirdPortable\ThunderbirdPortable.exe"                      & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"

        Case "USB"

            pFileInfo.Add "GoogleChrome.lnk"                          , pRunDrive & "\Tools\GoogleChromePortable\GoogleChromePortable.exe"                    & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
            pFileInfo.Add "Thunderbird.lnk"                           , pRunDrive & "\Tools\ThunderbirdPortable\ThunderbirdPortable.exe"                      & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"

    End Select

    '★FileEdit★
    pFileInfo.Add "CaptureSTAFF.lnk"                          , pRunDrive & "\Tools\CaptureSTAFF\Capt_St.exe"                                         & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "GIMP.lnk"                                  , pRunDrive & "\Tools\GIMPPortable\GIMPPortable.exe"                                    & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "Greenshot.lnk"                             , pRunDrive & "\Tools\Greenshot\Greenshot.exe"                                          & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "ImgBurn.lnk"                               , pRunDrive & "\Tools\ImgBurnPortable\ImgBurn.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "PDFTKBuilder.lnk"                          , pRunDrive & "\Tools\PDFTKBuilderPortable\PDFTKBuilderPortable.exe"                    & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "PSSTPSST.lnk"                              , pRunDrive & "\Tools\PSSTPSST\PSSTPSST.exe"                                            & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "ResourceHacker.lnk"                        , pRunDrive & "\Tools\ResourceHacker\ResourceHacker.exe"                                & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "Stirling.lnk"                              , pRunDrive & "\Tools\stir131\Stirling.exe"                                             & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"

    Select Case runKbn

        Case "House", "USB", "NotePC"

            pFileInfo.Add "CDEx.lnk"                                  , pRunDrive & "\Tools\CDExPortable\CDExPortable.exe"                                    & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
            pFileInfo.Add "Mp3tag.lnk"                                , pRunDrive & "\Tools\mp3tag\Mp3tag.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
            pFileInfo.Add "Mp3Gain.lnk"                               , pRunDrive & "\Tools\wxMP3gainPortable\wxMP3gainPortable.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"

    End Select

    '★Player･Viewer★
    pFileInfo.Add "Calibre.lnk"                               , pRunDrive & "\Tools\CalibrePortable\calibre-portable.exe"                             & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"              & "|" & ""                                         & "|" & pRunDrive & "\Tools\CalibrePortable\calibre-portable.exe"         & "|" & pRunDrive & "\Tools\CalibrePortable"
    pFileInfo.Add "FoxitReader.lnk"                           , pRunDrive & "\Tools\FoxitReaderPortable\FoxitReaderPortable.exe"                      & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"
    pFileInfo.Add "PDF-XChangeViewer.lnk"                     , pRunDrive & "\Tools\PDF-XChangeViewerPortable\PDF-XChangeViewerPortable.exe"          & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"
    pFileInfo.Add "XnView.lnk"                                , pRunDrive & "\Tools\XnViewPortable\XnViewPortable.exe"                                & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"


    Select Case runKbn

        Case "House"

            pFileInfo.Add "Kindle.lnk"                                , "%UserProfile%\AppData\Local\Amazon\Kindle\application\Kindle.exe"                    & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"
            pFileInfo.Add "foobar2000.lnk"                            , pRunDrive & "\Tools\foobar2000\foobar2000.exe"                                        & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"
            pFileInfo.Add "MangaMeeya.lnk"                            , pRunDrive & "\Tools\MangaMeeya_73\MangaMeeya.exe"                                     & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"
            pFileInfo.Add "VLC Media Player.lnk"                      , pRunDrive & "\Tools\VLCPortable\VLCPortable.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"

        Case "House", "USB", "NotePC"

            pFileInfo.Add "foobar2000.lnk"                            , pRunDrive & "\Tools\foobar2000\foobar2000.exe"                                        & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"
            pFileInfo.Add "MangaMeeya.lnk"                            , pRunDrive & "\Tools\MangaMeeya_73\MangaMeeya.exe"                                     & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"
            pFileInfo.Add "VLC Media Player.lnk"                      , pRunDrive & "\Tools\VLCPortable\VLCPortable.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"

    End Select

    '★Maintenance★
    pFileInfo.Add "Autoruns.lnk"                              , pRunDrive & "\Tools\AutorunsPortable\AutorunsPortable.exe"                            & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "CCleaner.lnk"                              , pRunDrive & "\Tools\CCleanerPortable\CCleaner64.exe"                                  & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "ChangeKey.lnk"                             , pRunDrive & "\Tools\ChangeKey_v150\ChgKey.exe"                                        & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "CPU-Z.lnk"                                 , pRunDrive & "\Tools\CPU-ZPortable\CPU-ZPortable.exe"                                  & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "CrystalDiskInfo.lnk"                       , pRunDrive & "\Tools\CrystalDiskInfoPortable\CrystalDiskInfoPortable.exe"              & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "CrystalDiskMark.lnk"                       , pRunDrive & "\Tools\CrystalDiskMarkPortable\CrystalDiskMarkPortable.exe"              & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "Defraggler.lnk"                            , pRunDrive & "\Tools\DefragglerPortable\Defraggler64.exe"                              & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "FromHDDtoSSD.lnk"                          , pRunDrive & "\Tools\FromHDDtoSSD\FromHDDtoSSD_64.exe"                                 & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "GPU-Z.lnk"                                 , pRunDrive & "\Tools\GPU-ZPortable\GPU-ZPortable.exe"                                  & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "HWMonitor.lnk"                             , pRunDrive & "\Tools\HWMonitor\HWMonitor_x64.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "IObitUninstaller.lnk"                      , pRunDrive & "\Tools\IObitUninstallerPortable\IObitUninstallerPortable.exe"            & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "PCパフォーマンスチェッカー.lnk"            , pRunDrive & "\Tools\PerformanceChecker\PerformanceChecker.exe"                        & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "ProcessExplorer.lnk"                       , pRunDrive & "\Tools\ProcessExplorerPortable\ProcessExplorerPortable.exe"              & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "ProcessMonitor.lnk"                        , pRunDrive & "\Tools\ProcessMonitorPortable\ProcessMonitorPortable.exe"                & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "TCPView.lnk"                               , pRunDrive & "\Tools\TCPViewPortable\TCPViewPortable.exe"                              & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "SpeedyFox.lnk"                             , pRunDrive & "\Tools\SpeedyFoxPortable\SpeedyFoxPortable.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "SystemExplorer.lnk"                        , pRunDrive & "\Tools\SystemExplorerPortable\SystemExplorerPortable.exe"                & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "Wireshark.lnk"                             , pRunDrive & "\Tools\WiresharkPortable\WiresharkPortable.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"

    '★MicrosoftOffice★
    pFileInfo.Add "Access.lnk"                                , pRunDrive & "\Tools\MicrosoftOffice\RunAccess.exe"                                    & "|" & pRunDrive & "\Tools\Shortcuts\MicrosoftOffice\"
    pFileInfo.Add "Excel.lnk"                                 , pRunDrive & "\Tools\MicrosoftOffice\RunExcel.exe"                                     & "|" & pRunDrive & "\Tools\Shortcuts\MicrosoftOffice\"
    pFileInfo.Add "PowerPoint.lnk"                            , pRunDrive & "\Tools\MicrosoftOffice\RunPowerPoint.exe"                                & "|" & pRunDrive & "\Tools\Shortcuts\MicrosoftOffice\"
    pFileInfo.Add "Word.lnk"                                  , pRunDrive & "\Tools\MicrosoftOffice\RunWord.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\MicrosoftOffice\"

    '★LibreOffice★
    pFileInfo.Add "LibreOffice.lnk"                           , pRunDrive & "\Tools\LibreOfficePortable\LibreOfficePortable.exe"                      & "|" & pRunDrive & "\Tools\Shortcuts\LibreOffice\"
    pFileInfo.Add "Base(Access).lnk"                          , pRunDrive & "\Tools\LibreOfficePortable\LibreOfficeBasePortable.exe"                  & "|" & pRunDrive & "\Tools\Shortcuts\LibreOffice\"
    pFileInfo.Add "Calc(Excel).lnk"                           , pRunDrive & "\Tools\LibreOfficePortable\LibreOfficeCalcPortable.exe"                  & "|" & pRunDrive & "\Tools\Shortcuts\LibreOffice\"
    pFileInfo.Add "Impress(PowerPoint).lnk"                   , pRunDrive & "\Tools\LibreOfficePortable\LibreOfficeImpressPortable.exe"               & "|" & pRunDrive & "\Tools\Shortcuts\LibreOffice\"
    pFileInfo.Add "Writer(Word).lnk"                          , pRunDrive & "\Tools\LibreOfficePortable\LibreOfficeWriterPortable.exe"                & "|" & pRunDrive & "\Tools\Shortcuts\LibreOffice\"

    '★Development★
    pFileInfo.Add "A5SQL Mk-2.lnk"                            , pRunDrive & "\Tools\A5SQLMk-2\A5M2.exe"                                               & "|" & pRunDrive & "\Tools\Shortcuts\Development\"
    pFileInfo.Add "BusyBox.lnk"                               , pRunDrive & "\Tools\BusyBox\RunBusyBox.bat"                                           & "|" & pRunDrive & "\Tools\Shortcuts\Development\"
    pFileInfo.Add "cmd.lnk"                                   , "%windir%\system32\cmd.exe"                                                           & "|" & pRunDrive & "\Tools\Shortcuts\Development\"                & "|" & ""                                         & "|" & "%windir%\system32\cmd.exe"                                       & "|" & "%windir%\system32"
    pFileInfo.Add "HeidiSQL.lnk"                              , pRunDrive & "\Tools\HeidiSQL\heidisql.exe"                                            & "|" & pRunDrive & "\Tools\Shortcuts\Development\"
    pFileInfo.Add "PowerShell.lnk"                            , "%windir%\System32\WindowsPowerShell\v1.0\powershell.exe"                             & "|" & pRunDrive & "\Tools\Shortcuts\Development\"                & "|" & ""                                         & "|" & "%windir%\System32\WindowsPowerShell\v1.0\powershell.exe"         & "|" & "%windir%\system32"
    pFileInfo.Add "RLogin.lnk"                                , pRunDrive & "\Tools\RLogin\RLogin.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\Development\"
    pFileInfo.Add "WinMerge.lnk"                              , pRunDrive & "\Tools\WinMergePortable\WinMergePortable.exe"                            & "|" & pRunDrive & "\Tools\Shortcuts\Development\"
    pFileInfo.Add "Zeal.lnk"                                  , pRunDrive & "\Tools\zeal-portable\zeal.exe"                                           & "|" & pRunDrive & "\Tools\Shortcuts\Development\"

    Select Case runKbn

        Case "Company", "House", "NotePC"

            pFileInfo.Add "Docker For Windows.lnk"                    , """" & pRunDrive & "\Program Files\Docker\Docker\Docker for Windows.exe"""            & "|" & pRunDrive & "\Tools\Shortcuts\Development\"
            pFileInfo.Add "GitBash(Laravel).lnk"                      , """" & pRunDrive & "\Program Files\Git\git-bash.exe"""                                & "|" & pRunDrive & "\Tools\Shortcuts\Development\"                & "|" & ""                                         & "|" & """" & pRunDrive & "\Program Files\Git\git-bash.exe"""            & "|" & "%UserProfile%\LaravelProjects\Homestead"
            pFileInfo.Add "GitBash(Project).lnk"                      , """" & pRunDrive & "\Program Files\Git\git-bash.exe"""                                & "|" & pRunDrive & "\Tools\Shortcuts\Development\"                & "|" & ""                                         & "|" & """" & pRunDrive & "\Program Files\Git\git-bash.exe"""            & "|" & "%SystemRoot%\..\Project"
            pFileInfo.Add "GitBash(UserFolder).lnk"                   , """" & pRunDrive & "\Program Files\Git\git-bash.exe"""                                & "|" & pRunDrive & "\Tools\Shortcuts\Development\"                & "|" & ""                                         & "|" & """" & pRunDrive & "\Program Files\Git\git-bash.exe"""            & "|" & "%UserProfile%"
            pFileInfo.Add "Oracle VM VirtualBox.lnk"                  , """" & pRunDrive & "\Program Files\Oracle\VirtualBox\VirtualBox.exe"""                & "|" & pRunDrive & "\Tools\Shortcuts\Development\"
            pFileInfo.Add "SourceTree.lnk"                            , "%UserProfile%\AppData\Local\SourceTree\SourceTree.exe"                               & "|" & pRunDrive & "\Tools\Shortcuts\Development\"
            pFileInfo.Add "TeraTerm.lnk"                              , "%ProgramFiles(x86)%\teraterm\ttermpro.exe"                                           & "|" & pRunDrive & "\Tools\Shortcuts\Development\"
            pFileInfo.Add "Visual Studio 2017.lnk"                    , "%ProgramFiles(x86)%\Microsoft Visual Studio\2017\Community\Common7\IDE\devenv.exe"   & "|" & pRunDrive & "\Tools\Shortcuts\Development\"
            pFileInfo.Add "WinsSCP.lnk"                               , "%ProgramFiles(x86)%\WinSCP\WinSCP.exe"                                               & "|" & pRunDrive & "\Tools\Shortcuts\Development\"

    End Select

    '★Download・Recording★
    pFileInfo.Add "CamStudio.lnk"                             , pRunDrive & "\Tools\CamStudioPortable\CamStudioPortable.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\Download・Recording\"

    '★Emulator★
    Select Case runKbn

        Case "House", "NotePC"

            pFileInfo.Add "NoxAppPlayer.lnk"                          , "%ProgramFiles(x86)%\Nox\bin\Nox.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\Emulator\"
            pFileInfo.Add "NoxAppPlayerMulti-Drive.lnk"               , "%ProgramFiles(x86)%\Nox\bin\MultiPlayerManager.exe"                                 & "|" & pRunDrive & "\Tools\Shortcuts\Emulator\"

    End Select

    '★OtherTool★
    pFileInfo.Add "7-Zip.lnk"                                 , pRunDrive & "\Tools\7-ZipPortable\7-ZipPortable.exe"                                  & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "CrenaHtml2jpg.lnk"                         , pRunDrive & "\Tools\CrenaHtml2jpg\CrenaHtml2jpg.exe"                                  & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "DeInput.lnk"                               , pRunDrive & "\Tools\DeInput\DeInput.exe"                                              & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "Everything.lnk"                            , pRunDrive & "\Tools\Everything\Everything64.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "FastCopy.lnk"                              , pRunDrive & "\Tools\FastCopyPortable\FastCopyPortable.exe"                            & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "Filistry.lnk"                              , pRunDrive & "\Tools\Filistry136\Filistry.exe"                                         & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "FireFileCopy.lnk"                          , pRunDrive & "\Tools\FireFileCopy\FFC.exe"                                             & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "FitWin.lnk"                                , pRunDrive & "\Tools\fitwin\fitwin.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "Fスクリーンキーボード.lnk"                 , pRunDrive & "\Tools\fkey\fkey.exe"                                                    & "|" & pRunDrive & "\Tools\Shortcuts\Other\"                      & "|" & ""                                         & "|" & pRunDrive & "\Tools\fkey\fkey.exe"                                & "|" & pRunDrive & "\Tools\fkey"
    pFileInfo.Add "IObitUnlocker.lnk"                         , pRunDrive & "\Tools\IObitUnlockerPortable\IObitUnlockerPortable.exe"                  & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "TanZIP.lnk"                                , pRunDrive & "\Tools\TanZIP\TanZIP.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "TransStickies.lnk"                         , pRunDrive & "\Tools\AutoHotKey\Tools\TransStickies\TransStickies.exe"                 & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "makeexe.lnk"                               , pRunDrive & "\Tools\makeexe\"                                                         & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "Meiryo UIも大嫌い!!.lnk"                   , pRunDrive & "\Tools\noMeiryoUI\noMeiryoUI.exe"                                        & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "pointClip.lnk"                             , pRunDrive & "\Tools\PointClip\pointClip.exe"                                          & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "RoboCopyGUI.lnk"                           , pRunDrive & "\Tools\RoboCopyGUI\RoboCopyGUI.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "RunSpp.lnk"                                , pRunDrive & "\Tools\SPP\RunSpp.bat"                                                   & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "StopWatchD.lnk"                            , pRunDrive & "\Tools\StopWatchD\StopWatchD.exe"                                        & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "VeraCrypt.lnk"                             , pRunDrive & "\Tools\VeraCrypt\VeraCrypt.exe"                                          & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "XMind.lnk"                                 , pRunDrive & "\Tools\XMind\XMind.exe"                                                  & "|" & pRunDrive & "\Tools\Shortcuts\Other\"

    Select Case runKbn

        Case "House", "USB" ,"NotePC"

            pFileInfo.Add "JoyToKey.lnk"                              , pRunDrive & "\Tools\JoyToKey\JoyToKey.exe"                                            & "|" & pRunDrive & "\Tools\Shortcuts\Other\"

    End Select

    '★クリップボード整形のリンクを作成★
    pFileInfo.Add "00_FIFOモード切り替え.lnk"                 , pRunDrive & "\Tools\clibor\Clibor.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\ClipConversion\"             & "|" & "/ff"
    pFileInfo.Add "01_各行先頭に「 ＞ 」を挿入（引用文）.lnk" , pRunDrive & "\Tools\clibor\Clibor.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\ClipConversion\"             & "|" & "/sd 1"
    pFileInfo.Add "02_各行先頭に「 001： 」の連番を挿入.lnk"  , pRunDrive & "\Tools\clibor\Clibor.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\ClipConversion\"             & "|" & "/sd 2"
    pFileInfo.Add "03_各行を「 ” 」で囲む.lnk"                , pRunDrive & "\Tools\clibor\Clibor.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\ClipConversion\"             & "|" & "/sd 3"
    pFileInfo.Add "04_各行を「 ' 」で囲む.lnk"                , pRunDrive & "\Tools\clibor\Clibor.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\ClipConversion\"             & "|" & "/sd 4"
    pFileInfo.Add "05_「大文字」に変換.lnk"                   , pRunDrive & "\Tools\clibor\Clibor.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\ClipConversion\"             & "|" & "/sd 5"
    pFileInfo.Add "06_「小文字」に変換.lnk"                   , pRunDrive & "\Tools\clibor\Clibor.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\ClipConversion\"             & "|" & "/sd 6"
    pFileInfo.Add "07_「全角」を「半角」に変換.lnk"           , pRunDrive & "\Tools\clibor\Clibor.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\ClipConversion\"             & "|" & "/sd 7"
    pFileInfo.Add "08_「半角」を「全角」に変換.lnk"           , pRunDrive & "\Tools\clibor\Clibor.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\ClipConversion\"             & "|" & "/sd 8"
    pFileInfo.Add "09_「カタカナ」を「ひらがな」に変換.lnk"   , pRunDrive & "\Tools\clibor\Clibor.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\ClipConversion\"             & "|" & "/sd 9"
    pFileInfo.Add "10_「ひらがな」を「カタカナ」に変換.lnk"   , pRunDrive & "\Tools\clibor\Clibor.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\ClipConversion\"             & "|" & "/sd 10"

    '★お気に入りディレクトリのリンクを作成★
    pFileInfo.Add "コンピュータ.lnk"                          , objAppli.Namespace(17).Self.Path                                                      & "|" & pRunDrive & "\Tools\Shortcuts\Favorite\"
    pFileInfo.Add "デスクトップ.lnk"                          , objShell.SpecialFolders("Desktop")                                                    & "|" & pRunDrive & "\Tools\Shortcuts\Favorite\"
    pFileInfo.Add "Tools.lnk"                                 , pRunDrive & "\Tools\"                                                                 & "|" & pRunDrive & "\Tools\Shortcuts\Favorite\"
    pFileInfo.Add "Bookmark.lnk"                              , pRunDrive & "\Tools\Shortcuts\Bookmark\"                                              & "|" & pRunDrive & "\Tools\Shortcuts\Favorite\"
    pFileInfo.Add "ダウンロード.lnk"                          , objAppli.Namespace(40).Self.Path & "\Downloads\"                                      & "|" & pRunDrive & "\Tools\Shortcuts\Favorite\"
    pFileInfo.Add "CreateScript.lnk"                          , pRunDrive & "\Tools\CreateScript"                                                     & "|" & pRunDrive & "\Tools\Shortcuts\Favorite\"

    '★Windows（アクセサリ）★
    pFileInfo.Add "Windows Media Player.lnk"                  , "%ProgramFiles(x86)%\Windows Media Player\wmplayer.exe"                               & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"
    pFileInfo.Add "コマンドプロンプト.lnk"                    , "%windir%\system32\cmd.exe"                                                           & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"         & "|" & ""                                         & "|" & "%windir%\system32\cmd.exe"                                       & "|" & "%windir%\system32"
    pFileInfo.Add "タスクマネージャー.lnk"                    , "%windir%\system32\taskmgr.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"
    pFileInfo.Add "ペイント.lnk"                              , "%windir%\system32\mspaint.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"
    pFileInfo.Add "メモ帳.lnk"                                , "%windir%\system32\notepad.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"
    pFileInfo.Add "リモートデスクトップ.lnk"                  , "%windir%\system32\mstsc.exe"                                                         & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"
    pFileInfo.Add "ワードパッド.lnk"                          , "%ProgramFiles%\Windows NT\Accessories\wordpad.exe"                                   & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"
    pFileInfo.Add "電卓.lnk"                                  , "%windir%\system32\calc.exe"                                                          & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"

    '★Windows（コントロールパネル）★
    pFileInfo.Add "Windows Update.lnk"                        , "%windir%\system32\wuapp.exe"                                                         & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\" & "|" & ""                                         & "|" & "%WinDir%\System32\shell32.dll, 46"
    pFileInfo.Add "コントロールパネル.lnk"                    , "%windir%\system32\control.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\"
    pFileInfo.Add "システム.lnk"                              , "%windir%\system32\control.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\" & "|" & "/name Microsoft.System"                   & "|" & "%WinDir%\System32\shell32.dll, 272"
    pFileInfo.Add "デバイスマネージャー.lnk"                  , "%windir%\system32\devmgmt.msc"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\"
    pFileInfo.Add "ネットワークと共有センター.lnk"            , "%windir%\system32\control.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\" & "|" & "/name Microsoft.NetworkAndSharingCenter"  & "|" & "%WinDir%\System32\shell32.dll, 276"
    pFileInfo.Add "フォルダーオプション.lnk"                  , "%windir%\system32\control.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\" & "|" & "folders"                                  & "|" & "%WinDir%\System32\shell32.dll, 110"
    pFileInfo.Add "プログラムの追加と削除.lnk"                , "%windir%\system32\appwiz.cpl"                                                        & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\" & "|" & ""                                         & "|" & "%WinDir%\System32\shell32.dll, 162"
    pFileInfo.Add "個人設定.lnk"                              , "%windir%\system32\control.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\" & "|" & "/name Microsoft.Personalization"          & "|" & "%WinDir%\System32\shell32.dll, 141"
    pFileInfo.Add "電源オプション.lnk"                        , "%windir%\system32\powercfg.cpl"                                                      & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\" & "|" & ""                                         & "|" & "%windir%\system32\powercfg.cpl, 0"

    '★Windows（その他）★
    pFileInfo.Add "DirectX診断ツール.lnk"                     , "%windir%\system32\dxdiag.exe"                                                        & "|" & pRunDrive & "\Tools\Shortcuts\Windows\その他\"
    pFileInfo.Add "Internet Explorer.lnk"                     , "%ProgramFiles%\Internet Explorer\iexplore.exe"                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\その他\"
    pFileInfo.Add "Windowsモビリティセンター.lnk"             , "%windir%\system32\mblctr.exe"                                                        & "|" & pRunDrive & "\Tools\Shortcuts\Windows\その他\"
    pFileInfo.Add "システム構成.lnk"                          , "%windir%\system32\msconfig.exe"                                                      & "|" & pRunDrive & "\Tools\Shortcuts\Windows\その他\"
    pFileInfo.Add "ディスクの管理.lnk"                        , "%windir%\system32\diskmgmt.msc"                                                      & "|" & pRunDrive & "\Tools\Shortcuts\Windows\その他\"
    pFileInfo.Add "レジストリエディタ.lnk"                    , "%windir%\SysWOW64\regedit.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\その他\"
    pFileInfo.Add "画面のプロパティ.lnk"                      , "%windir%\system32\desk.cpl"                                                          & "|" & pRunDrive & "\Tools\Shortcuts\Windows\その他\"             & "|" & ""                                         & "|" & "%WinDir%\System32\shell32.dll, 174"
    pFileInfo.Add "ハードウェアの安全な取り外し.lnk"          , "%windir%\system32\RunDll32.exe"                                                      & "|" & pRunDrive & "\Tools\Shortcuts\Windows\その他\"             & "|" & "shell32.dll,Control_RunDLL HotPlug.dll"   & "|" & "%SystemRoot%\system32\hotplug.dll, 0"
    pFileInfo.Add "コンピュータのロック.lnk"                  , "%windir%\System32\rundll32.exe"                                                      & "|" & pRunDrive & "\Tools\Shortcuts\Windows\その他\"             & "|" & "user32.dll,LockWorkStation"               & "|" & "%WinDir%\System32\shell32.dll, 44"
    pFileInfo.Add "ShutDownDialog.lnk"                        , pRunDrive & "\Tools\CreateScript\ShowShutdownWindowsDialog\exe\ShutDownDialog.exe"    & "|" & pRunDrive & "\Tools\Shortcuts\Windows\その他\"

    '★Windows（管理ツール）★
    pFileInfo.Add "イベントビューアー.lnk"                    , "%windir%\system32\eventvwr.exe"                                                      & "|" & pRunDrive & "\Tools\Shortcuts\Windows\管理ツール\"
    pFileInfo.Add "コンピュータの管理.lnk"                    , "%windir%\system32\compmgmt.msc"                                                      & "|" & pRunDrive & "\Tools\Shortcuts\Windows\管理ツール\"
    pFileInfo.Add "サービス.lnk"                              , "%windir%\system32\services.msc"                                                      & "|" & pRunDrive & "\Tools\Shortcuts\Windows\管理ツール\"
    pFileInfo.Add "データソース(ODBC)_32bit.lnk"              , "%windir%\SysWOW64\odbcad32.exe"                                                      & "|" & pRunDrive & "\Tools\Shortcuts\Windows\管理ツール\"         & "|" & ""                                         & "|" & "%windir%\system32\odbcad32.exe"
    pFileInfo.Add "データソース(ODBC)_64bit.lnk"              , "%windir%\system32\odbcad32.exe"                                                      & "|" & pRunDrive & "\Tools\Shortcuts\Windows\管理ツール\"
    pFileInfo.Add "パフォーマンスモニター.lnk"                , "%windir%\system32\perfmon.msc"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\管理ツール\"
    pFileInfo.Add "管理ツール.lnk"                            , "%windir%\system32\control.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\管理ツール\"         & "|" & "admintools"                               & "|" & "%windir%\system32\imageres.dll, 109"

    '★Explorersのショートカットを作成★
    pFileInfo.Add "01.Explorer.lnk"                           , "%windir%\explorer.exe"                                                               & "|" & pRunDrive & "\Tools\Shortcuts\Explorers\"                  & "|" & "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    pFileInfo.Add "02.DoubleExplorer.lnk"                     , pRunDrive & "\Tools\AutoHotKey\Tools\Explorers\Explorers.exe"                         & "|" & pRunDrive & "\Tools\Shortcuts\Explorers\"                  & "|" & ""                                         & "|" & pRunDrive & "\Tools\AutoHotKey\Tools\Explorers\ico\Explorers.ico"
    pFileInfo.Add "03.FourthExplorer.lnk"                     , pRunDrive & "\Tools\AutoHotKey\Tools\Explorers\Explorers.exe"                         & "|" & pRunDrive & "\Tools\Shortcuts\Explorers\"                  & "|" & "4"                                        & "|" & pRunDrive & "\Tools\AutoHotKey\Tools\Explorers\ico\Explorers.ico"

End Function

'***********************************************************************
'* 処理名   ： フォルダ作成処理                                        *
'* 引数     ： pPath  作成するフォルダパス（フルパス）                 *
'* 処理内容 ： 再帰的にフォルダを作成していきます                      *
'*             ドライブ → ドライブ\階層１→ドライブ\階層１\階層２\    *
'*             → ドライブ\階層１\階層２\対象フォルダ                  *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Function CreateNotExistFolder(ByVal pPath)

    '変数宣言・インスタンス作成
    Dim objFso       : set objFso   = WScript.CreateObject("Scripting.FileSystemObject")  'FileSystemObjectオブジェクト
    Dim driveName    : driveName    = Left(objFso.GetDriveName(pPath),2)                  'ドライブ名を取得
    Dim parentFolder : parentFolder = objFso.GetParentFolderName(pPath)                   '親フォルダー名を取得

    '対象のドライブが存在する時
    If objFso.DriveExists(driveName) Then

        set objDrive = objFso.GetDrive(driveName) 'Driveオブジェクトを作成

    Else

        Exit Function                             '処理を終了

    End If

    'ドライブの準備ができている時
    If objDrive.IsReady Then

        '拡張子文字列が取得出来た場合(ファイルの時)
        If Len(objFso.GetExtensionName(pPath)) > 0 Then

            '親フォルダーが存在しない時、対象パスから親フォルダー作成する（再帰的）
            If Not(objFso.FolderExists(parentFolder)) Then CreateNotExistFolder(parentFolder)

        Else

            '対象フォルダーが存在しない時
            If Not(objFso.FolderExists(pPath)) Then

                '親フォルダーを作成後、対象フォルダーを作成（再帰的）
                CreateNotExistFolder(parentFolder)
                objFso.CreateFolder(pPath)

            End If

        End If

    End If

    'オブジェクトの破棄
    Set objFso   = Nothing
    Set objDrive = Nothing

End Function
