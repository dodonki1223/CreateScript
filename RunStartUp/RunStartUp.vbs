'**************************************************************************************
'* プログラム名 ： スタートアップ処理スクリプト                                       *
'* 処理概要     ： スタートアップ時に実行するスクリプト。実行されたドライブからOrchis *
'*                 で使用するショートカットファイルのリンク先を作成し直す。           *
'*                 スタートアップ時に実行されて欲しいプログラムを一括で実行する。     *
'* メモ         ： このファイルをショートカットにしてコマンドライン引数を指定すること *
'*                 ★使用例★                                                         *
'*                   C:\Tools\CreateScript\RunStartUp\RunStartUp.vbs "House"          *
'*                   C:\Tools\CreateScript\RunStartUp\RunStartUp.vbs "USB"            *
'*                 ※実行する環境によりコマンドライン引数を変更する事                 *
'*                 URLファイルの作成方法                                              *
'*                   ファイル名を「○○○.url」形式にしショートカット先にURLを指定    *
'* 設定         ： このスクリプトのデフォルト設定はUSBで実行されます                  *
'**************************************************************************************

'--------------------------------------
' 設定
'--------------------------------------
'※実行区分 「House：家、USB：USB」
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

        pRunFile.Add "WheelAccele"        , pRunDrive & "\Tools\WheelAccele\WheelAccele.exe"
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

                Case "House"

                    pRunFile.Add "GoogleChrome"       , """" & pRunDrive & "\Program Files\Google\Chrome\Application\chrome.exe"""

                Case "USB"

                    pRunFile.Add "GoogleChrome"       , pRunDrive & "\Tools\GoogleChromePortable\GoogleChromePortable.exe"

            End Select

        End If

    End If

    pRunFile.Add "Clibor"             , pRunDrive & "\Tools\clibor\Clibor.exe"
    pRunFile.Add "AutoHotKeyTool"     , pRunDrive & "\Tools\AutoHotKey\AutoHotKeyTool.exe"
    pRunFile.Add "AkabeiMonitor"      , pRunDrive & "\Tools\AkabeiMonitor\akamoni.exe"

    Select Case runKbn

        Case "House"

            pRunFile.Add "BijinTokeiGadget"   , pRunDrive & "\Tools\BijinTokeiGadget\BijinTokeiGadget.exe"
            pRunFile.Add "BijoLinuxGadget"    , pRunDrive & "\Tools\BijoLinuxGadget\BijoLinuxGadget.exe"
            pRunFile.Add "T-Clock"            , pRunDrive & "\Tools\T-Clock\Clock64.exe"
            pRunFile.Add "Slack"              , """" & "%UserProfile%\AppData\Local\slack\slack.exe"""
            pRunFile.Add "GoogleDrive"        , """" & pRunDrive & "\Program Files\Google\Drive File Stream\59.0.3.0\GoogleDriveFS.exe"""

    End Select

    '実行ドライブ文字列を取得
    Dim driveStr : driveStr = Left(pRunDrive, 1)

    'ドライブごと起動するOrchisを変更する
    Select Case driveStr

        Case "C"

            Select Case runKbn

                Case "House"

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
    pFileInfo.Add "AkabeiMonitor.lnk"                         , pRunDrive & "\Tools\AkabeiMonitor\akamoni.exe"                                        & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "AutoHotKeyTool.lnk"                        , pRunDrive & "\Tools\AutoHotKey\AutoHotKeyTool.exe"                                    & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "BijinTokeiGadget.lnk"                      , pRunDrive & "\Tools\BijinTokeiGadget\BijinTokeiGadget.exe"                            & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "BijoLinuxGadget.lnk"                       , pRunDrive & "\Tools\BijoLinuxGadget\BijoLinuxGadget.exe"                              & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "Clibor.lnk"                                , pRunDrive & "\Tools\clibor\Clibor.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "MouseGestureL.lnk"                         , pRunDrive & "\Tools\MouseGestureL\MouseGestureL.exe"                                  & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "Orchis.lnk"                                , pOrchisDirectory                                                                      & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"                            & "|" & ""                                 & "|" & pRunDrive & "\Program Files\Orchis\orchis.exe"
    pFileInfo.Add "WheelAccele.lnk"                           , pRunDrive & "\Tools\WheelAccele\WheelAccele.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
    pFileInfo.Add "T-Clock.lnk"                               , pRunDrive & "\Tools\T-Clock\Clock64.exe"                                              & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"

    Select Case runKbn

        Case "House"

            pFileInfo.Add "GoogleDrive.lnk"                           , """" & pRunDrive & "\Program Files\Google\Drive File Stream\59.0.3.0\GoogleDriveFS.exe""" & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "Slack.lnk"                                 , """" & "%UserProfile%\AppData\Local\slack\slack.exe"""                                    & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"
            pFileInfo.Add "Logicool Options.lnk"                      , """" & pRunDrive & "\Program Files\Logicool\LogiOptions\LogiOptions.exe"""                & "|" & pRunDrive & "\Tools\Shortcuts\StartUp\"

    End Select

    '★OftenUse★
    pFileInfo.Add "FolderFileList.lnk"                        , pRunDrive & "\Tools\FolderFileList\FolderFileList.exe"                                & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
    pFileInfo.Add "FolderFileListDebug.lnk"                   , pRunDrive & "\Tools\FolderFileList\FolderFileList_Debug.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
    pFileInfo.Add "00_ImageForClipboard.lnk"                  , pRunDrive & "\Tools\ImageForClipboard\ImageForClipboard.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\ImageForClipboard\"
    pFileInfo.Add "01_ヘルプ.lnk"                             , pRunDrive & "\Tools\ImageForClipboard\ImageForClipboard.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\ImageForClipboard\"         & "|" & "/?"
    pFileInfo.Add "02_クリップボード内の画像を表示.lnk"       , pRunDrive & "\Tools\ImageForClipboard\ImageForClipboard.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\ImageForClipboard\"         & "|" & "/AutoClose 5 /ImageSize 40"
    pFileInfo.Add "03_クリップボード内の画像を保存.lnk"       , pRunDrive & "\Tools\ImageForClipboard\ImageForClipboard.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\ImageForClipboard\"         & "|" & "/AutoClose 2 /ImageSize 40 /AutoSave %UserProfile%\Downloads\ /Extension png"
    pFileInfo.Add "ReduceMemory.lnk"                          , pRunDrive & "\Tools\ReduceMemory\ReduceMemory.exe"                                    & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
    pFileInfo.Add "TeamViewer.lnk"                            , pRunDrive & "\Tools\TeamViewerPortable\TeamViewerPortable.exe"                        & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
    pFileInfo.Add "Visual Studio Code.lnk"                    , """" & "%UserProfile%\AppData\Local\Programs\Microsoft VS Code\Code.exe"""            & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"
    pFileInfo.Add "X-Finder.lnk"                              , pRunDrive & "\Tools\X-Finder\xf64.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"

    Select Case runKbn

        Case "House"

            pFileInfo.Add "GoogleChrome.lnk"                          , """" & pRunDrive & "\Program Files\Google\Chrome\Application\chrome.exe"""                & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"

        Case "USB"

            pFileInfo.Add "GoogleChrome.lnk"                          , pRunDrive & "\Tools\GoogleChromePortable\GoogleChromePortable.exe"                        & "|" & pRunDrive & "\Tools\Shortcuts\OftenUse\"

    End Select

    '★FileEdit★
    pFileInfo.Add "GIMP.lnk"                                  , pRunDrive & "\Tools\GIMPPortable\GIMPPortable.exe"                                    & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "Greenshot.lnk"                             , pRunDrive & "\Tools\Greenshot\Greenshot.exe"                                          & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "ImgBurn.lnk"                               , pRunDrive & "\Tools\ImgBurnPortable\ImgBurn.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "PDFTKBuilder.lnk"                          , pRunDrive & "\Tools\PDFTKBuilderPortable\PDFTKBuilderPortable.exe"                    & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "PSSTPSST.lnk"                              , pRunDrive & "\Tools\PSSTPSST\PSSTPSST.exe"                                            & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "ResourceHacker.lnk"                        , pRunDrive & "\Tools\ResourceHacker\ResourceHacker.exe"                                & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "Stirling.lnk"                              , pRunDrive & "\Tools\stir131\Stirling.exe"                                             & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "CDEx.lnk"                                  , pRunDrive & "\Tools\CDExPortable\CDExPortable.exe"                                    & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "Mp3tag.lnk"                                , pRunDrive & "\Tools\mp3tag\Mp3tag.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"
    pFileInfo.Add "Mp3Gain.lnk"                               , pRunDrive & "\Tools\wxMP3gainPortable\wxMP3gainPortable.exe"                          & "|" & pRunDrive & "\Tools\Shortcuts\FileEdit\"

    '★Player･Viewer★
    pFileInfo.Add "Calibre.lnk"                               , pRunDrive & "\Tools\CalibrePortable\calibre-portable.exe"                             & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"              & "|" & ""                                         & "|" & pRunDrive & "\Tools\CalibrePortable\calibre-portable.exe"         & "|" & pRunDrive & "\Tools\CalibrePortable"
    pFileInfo.Add "IconExplorer.lnk"                          , pRunDrive & "\Tools\IconExplorer\IconExplorer.exe"                                    & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"
    pFileInfo.Add "Kindle.lnk"                                , "%UserProfile%\AppData\Local\Amazon\Kindle\application\Kindle.exe"                    & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"
    pFileInfo.Add "MusicBee.lnk"                              , pRunDrive & "\Tools\MusicBee\MusicBee.exe"                                            & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"
    pFileInfo.Add "MangaMeeya.lnk"                            , pRunDrive & "\Tools\MangaMeeya_73\MangaMeeya.exe"                                     & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"
    pFileInfo.Add "PDF-XChangeViewer.lnk"                     , pRunDrive & "\Tools\PDF-XChangeViewerPortable\PDF-XChangeViewerPortable.exe"          & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"
    pFileInfo.Add "VLC Media Player.lnk"                      , pRunDrive & "\Tools\VLCPortable\VLCPortable.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"
    pFileInfo.Add "XnView.lnk"                                , pRunDrive & "\Tools\XnViewPortable\XnViewPortable.exe"                                & "|" & pRunDrive & "\Tools\Shortcuts\Player･Viewer\"

    '★Maintenance★
    pFileInfo.Add "Autoruns.lnk"                              , pRunDrive & "\Tools\AutorunsPortable\AutorunsPortable.exe"                            & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "CCleaner.lnk"                              , pRunDrive & "\Tools\CCleanerPortable\CCleaner64.exe"                                  & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "ChangeKey.lnk"                             , pRunDrive & "\Tools\ChangeKey_v150\ChgKey.exe"                                        & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "CPU-Z.lnk"                                 , pRunDrive & "\Tools\CPU-ZPortable\CPU-ZPortable.exe"                                  & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "CrystalDiskInfo.lnk"                       , pRunDrive & "\Tools\CrystalDiskInfoPortable\CrystalDiskInfoPortable.exe"              & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "CrystalDiskMark.lnk"                       , pRunDrive & "\Tools\CrystalDiskMarkPortable\CrystalDiskMarkPortable.exe"              & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "Defraggler.lnk"                            , pRunDrive & "\Tools\DefragglerPortable\Defraggler64.exe"                              & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "GPU-Z.lnk"                                 , pRunDrive & "\Tools\GPU-ZPortable\GPU-ZPortable.exe"                                  & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "IObitUninstaller.lnk"                      , pRunDrive & "\Tools\IObitUninstallerPortable\IObitUninstallerPortable.exe"            & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "ProcessExplorer.lnk"                       , pRunDrive & "\Tools\ProcessExplorerPortable\ProcessExplorerPortable.exe"              & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "ProcessMonitor.lnk"                        , pRunDrive & "\Tools\ProcessMonitorPortable\ProcessMonitorPortable.exe"                & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"
    pFileInfo.Add "SystemExplorer.lnk"                        , pRunDrive & "\Tools\SystemExplorerPortable\SystemExplorerPortable.exe"                & "|" & pRunDrive & "\Tools\Shortcuts\Maintenance\"

    '★MicrosoftOffice★
    pFileInfo.Add "Access.lnk"                                , pRunDrive & "\Tools\MicrosoftOffice\RunAccess.exe"                                    & "|" & pRunDrive & "\Tools\Shortcuts\MicrosoftOffice\"
    pFileInfo.Add "Excel.lnk"                                 , pRunDrive & "\Tools\MicrosoftOffice\RunExcel.exe"                                     & "|" & pRunDrive & "\Tools\Shortcuts\MicrosoftOffice\"
    pFileInfo.Add "PowerPoint.lnk"                            , pRunDrive & "\Tools\MicrosoftOffice\RunPowerPoint.exe"                                & "|" & pRunDrive & "\Tools\Shortcuts\MicrosoftOffice\"
    pFileInfo.Add "Word.lnk"                                  , pRunDrive & "\Tools\MicrosoftOffice\RunWord.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\MicrosoftOffice\"

    '★Development★
    pFileInfo.Add "A5SQL Mk-2.lnk"                            , pRunDrive & "\Tools\A5SQLMk-2\A5M2.exe"                                               & "|" & pRunDrive & "\Tools\Shortcuts\Development\"
    pFileInfo.Add "cmd.lnk"                                   , "%windir%\system32\cmd.exe"                                                           & "|" & pRunDrive & "\Tools\Shortcuts\Development\"                & "|" & ""                                         & "|" & "%windir%\system32\cmd.exe"                                       & "|" & "%windir%\system32"
    pFileInfo.Add "PowerShell.lnk"                            , "%windir%\System32\WindowsPowerShell\v1.0\powershell.exe"                             & "|" & pRunDrive & "\Tools\Shortcuts\Development\"                & "|" & ""                                         & "|" & "%windir%\System32\WindowsPowerShell\v1.0\powershell.exe"         & "|" & "%windir%\system32"
    pFileInfo.Add "WinMerge.lnk"                              , pRunDrive & "\Tools\WinMergePortable\WinMergePortable.exe"                            & "|" & pRunDrive & "\Tools\Shortcuts\Development\"

    Select Case runKbn

        Case "House"

            pFileInfo.Add "Docker Desktop.lnk"                        , """" & pRunDrive & "\Program Files\Docker\Docker\Docker Desktop.exe"""                & "|" & pRunDrive & "\Tools\Shortcuts\Development\"
            pFileInfo.Add "GitBash.lnk"                               , """" & pRunDrive & "\Program Files\Git\git-bash.exe"""                                & "|" & pRunDrive & "\Tools\Shortcuts\Development\"                & "|" & ""                                         & "|" & """" & pRunDrive & "\Program Files\Git\git-bash.exe"""            & "|" & "%UserProfile%"
            pFileInfo.Add "GitKraken.lnk"                             , "%UserProfile%\AppData\Local\gitkraken\Update.exe"                                    & "|" & pRunDrive & "\Tools\Shortcuts\Development\"                & "|" & "--processStart gitkraken.exe"
            pFileInfo.Add "Oracle VM VirtualBox.lnk"                  , """" & pRunDrive & "\Program Files\Oracle\VirtualBox\VirtualBox.exe"""                & "|" & pRunDrive & "\Tools\Shortcuts\Development\"
            ' pFileInfo.Add "Visual Studio 2017.lnk"                    , "%ProgramFiles(x86)%\Microsoft Visual Studio\2017\Community\Common7\IDE\devenv.exe"   & "|" & pRunDrive & "\Tools\Shortcuts\Development\"


    End Select

    '★OtherTool★
    pFileInfo.Add "7-Zip.lnk"                                 , pRunDrive & "\Tools\7-ZipPortable\7-ZipPortable.exe"                                  & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "DeInput.lnk"                               , pRunDrive & "\Tools\DeInput\DeInput.exe"                                              & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "FastCopy.lnk"                              , pRunDrive & "\Tools\FastCopyPortable\FastCopyPortable.exe"                            & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "FireFileCopy.lnk"                          , pRunDrive & "\Tools\FireFileCopy\FFC.exe"                                             & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "FitWin.lnk"                                , pRunDrive & "\Tools\fitwin\fitwin.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "Fスクリーンキーボード.lnk"                 , pRunDrive & "\Tools\fkey\fkey.exe"                                                    & "|" & pRunDrive & "\Tools\Shortcuts\Other\"                      & "|" & ""                                         & "|" & pRunDrive & "\Tools\fkey\fkey.exe"                                & "|" & pRunDrive & "\Tools\fkey"
    pFileInfo.Add "IObitUnlocker.lnk"                         , pRunDrive & "\Tools\IObitUnlockerPortable\IObitUnlockerPortable.exe"                  & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "TanZIP.lnk"                                , pRunDrive & "\Tools\TanZIP\TanZIP.exe"                                                & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "makeexe.lnk"                               , pRunDrive & "\Tools\makeexe\"                                                         & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "pointClip.lnk"                             , pRunDrive & "\Tools\PointClip\pointClip.exe"                                          & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "RoboCopyGUI.lnk"                           , pRunDrive & "\Tools\RoboCopyGUI\RoboCopyGUI.exe"                                      & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "RunSpp.lnk"                                , pRunDrive & "\Tools\SPP\RunSpp.bat"                                                   & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "StopWatchD.lnk"                            , pRunDrive & "\Tools\StopWatchD\StopWatchD.exe"                                        & "|" & pRunDrive & "\Tools\Shortcuts\Other\"
    pFileInfo.Add "VeraCrypt.lnk"                             , pRunDrive & "\Tools\VeraCrypt\VeraCrypt.exe"                                          & "|" & pRunDrive & "\Tools\Shortcuts\Other\"

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

    '★Windows（Applications）★
    pFileInfo.Add "Applications.lnk"                          , "%windir%\explorer.exe"                                                               & "|" & pRunDrive & "\Tools\Shortcuts\Windows\"                    & "|" & "shell:appsfolder"

    '★Windows（設定）★
    pFileInfo.Add "設定.lnk"                                  , "%windir%\explorer.exe"                                                               & "|" & pRunDrive & "\Tools\Shortcuts\Windows\設定\"               & "|" & "ms-settings:"                             & "|" & "%WinDir%\System32\imageres.dll, 109"
    pFileInfo.Add "マルチモニター.lnk"                        , "%windir%\system32\control.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\設定\"               & "|" & "desk.cpl"                                 & "|" & "%WinDir%\System32\imageres.dll, 186"
    pFileInfo.Add "個人設定.lnk"                              , "%windir%\system32\control.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\設定\"               & "|" & "/name Microsoft.Personalization"          & "|" & "%WinDir%\System32\shell32.dll, 141"
    pFileInfo.Add "Windows Update.lnk"                        , "%windir%\explorer.exe"                                                               & "|" & pRunDrive & "\Tools\Shortcuts\Windows\設定\"               & "|" & "ms-settings:windowsupdate"                & "|" & "%WinDir%\System32\shell32.dll, 46"
    pFileInfo.Add "システム.lnk"                              , "%windir%\system32\control.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\設定\"               & "|" & "/name Microsoft.System"                   & "|" & "%WinDir%\System32\shell32.dll, 272"

    '★Windows（アクセサリ）★
    pFileInfo.Add "Windows Media Player.lnk"                  , "%ProgramFiles(x86)%\Windows Media Player\wmplayer.exe"                               & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"
    pFileInfo.Add "コマンドプロンプト.lnk"                    , "%windir%\system32\cmd.exe"                                                           & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"         & "|" & ""                                         & "|" & "%windir%\system32\cmd.exe"                                       & "|" & "%windir%\system32"
    pFileInfo.Add "タスクマネージャー.lnk"                    , "%windir%\system32\taskmgr.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"
    pFileInfo.Add "ペイント.lnk"                              , "%UserProfile%\AppData\Local\Microsoft\WindowsApps\pbrush.exe"                        & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"
    pFileInfo.Add "メモ帳.lnk"                                , "%windir%\system32\notepad.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"
    pFileInfo.Add "リモートデスクトップ.lnk"                  , "%windir%\system32\mstsc.exe"                                                         & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"
    pFileInfo.Add "ワードパッド.lnk"                          , "%ProgramFiles%\Windows NT\Accessories\wordpad.exe"                                   & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"
    pFileInfo.Add "電卓.lnk"                                  , "%windir%\system32\calc.exe"                                                          & "|" & pRunDrive & "\Tools\Shortcuts\Windows\アクセサリ\"

    '★Windows（コントロールパネル）★
    pFileInfo.Add "コントロールパネル.lnk"                    , "%windir%\system32\control.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\"
    pFileInfo.Add "デバイスマネージャー.lnk"                  , "%windir%\system32\devmgmt.msc"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\"
    pFileInfo.Add "ネットワークと共有センター.lnk"            , "%windir%\system32\control.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\" & "|" & "/name Microsoft.NetworkAndSharingCenter"  & "|" & "%WinDir%\System32\shell32.dll, 276"
    pFileInfo.Add "フォルダーオプション.lnk"                  , "%windir%\system32\control.exe"                                                       & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\" & "|" & "folders"                                  & "|" & "%WinDir%\System32\shell32.dll, 110"
    pFileInfo.Add "プログラムの追加と削除.lnk"                , "%windir%\system32\appwiz.cpl"                                                        & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\" & "|" & ""                                         & "|" & "%WinDir%\System32\shell32.dll, 162"
    pFileInfo.Add "電源オプション.lnk"                        , "%windir%\system32\powercfg.cpl"                                                      & "|" & pRunDrive & "\Tools\Shortcuts\Windows\コントロールパネル\" & "|" & ""                                         & "|" & "%windir%\system32\powercfg.cpl, 0"

    '★Windows（その他）★
    pFileInfo.Add "DirectX診断ツール.lnk"                     , "%windir%\system32\dxdiag.exe"                                                        & "|" & pRunDrive & "\Tools\Shortcuts\Windows\その他\"
    pFileInfo.Add "Microsoft Edge.lnk"                        , "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"                           & "|" & pRunDrive & "\Tools\Shortcuts\Windows\その他\"
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
