'--------------------------------------
' 定数宣言
'--------------------------------------
Const PORTABLE_APPS_DIRECTORY = "\Tools\PortableApps\PortableApps\"

'--------------------------------------
' 変数宣言・インスタンス作成
'--------------------------------------
Dim objAppli : Set objAppli   = WScript.CreateObject("Shell.Application")          'WScript.Applicationオブジェクト
Dim objFso   : set objFso     = WScript.CreateObject("Scripting.FileSystemObject") 'FileSystemObject
Dim objShell : Set objShell   = WScript.CreateObject("WScript.Shell")              'WScript.Shellオブジェクト
Dim symLinks : Set symLinks   = WScript.CreateObject("Scripting.Dictionary")       'シンボリックリンク格納Dictionary

'--------------------------------------
' 管理者権限で実行させる
'--------------------------------------
' 2回目以降は runas というコマンドライン引数を渡して実行する
if Wscript.Arguments.Count = 0 then
    objAppli.ShellExecute "wscript.exe", WScript.ScriptFullName & " runas", "", "runas", 1
    Wscript.Quit
end if

Main()

'***********************************************************************
'* 処理名   ： メイン処理                                              *
'* 引数     ： なし                                                    *
'* 処理内容 ： メイン処理                                              *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub Main()

    '--------------------------------------
    ' 実行ドライブを取得
    '--------------------------------------
    Dim runDrive : runDrive = objFSo.GetDriveName(WScript.ScriptFullName)

    '--------------------------------------
    ' シンボリックリンク作成処理
    '--------------------------------------
    Call AddSymLinks(symLinks ,runDrive)

    For Each key In symLinks.Keys

        '作成情報を切り分ける（①PortableApps管理下のフォルダ名、②インストール先フォルダ）
        Dim arySymLinks : arySymLinks = Split(symLinks(key), "|")

        'SymLinkの作成先とリンク先のディレクトリを取得する
        Dim symLinkPath       : symLinkPath = runDrive & PORTABLE_APPS_DIRECTORY & arySymLinks(0)
        Dim symLinkTargetPath : symLinkTargetPath = arySymLinks(1)

        '作成先のSymLinkが既にあるとエラーになるため事前に削除して強制的に再作成させる
        if objFso.FolderExists(symLinkPath) then
            objShell.Run "cmd /c rmdir " & symLinkPath, 0, false
        end if

        'シンボリックリンク作成のコマンドを実行していく
        'USBなどのデフォルトのファイルシステムのFat系だとシンボリックリンクの作成ができないためNTFSにあらかじめフォーマットする必要がある
        objShell.Run "cmd /c mklink /d " & symLinkPath &  " " & symLinkTargetPath, 0, false

    Next

    '--------------------------------------
    ' オブジェクト破棄処理
    '--------------------------------------
    Set objShell = Nothing
    Set objAppli = Nothing
    Set objFso   = Nothing

End Sub


'***********************************************************************
'* 処理名   ： シンボリックリンク作成用のディレクトリ追加処理          *
'* 引数     ： pSymLinks        作成ディレクトリ格納Dictionary         *
'*             pRunDrive        実行ドライブパス                       *
'* 処理内容 ： シンボリックリンク作成用のディレクトリををDictionaryに  *
'*             追加する                                                *
'* 戻り値   ： pRunSymLinks                                            *
'***********************************************************************
Function AddSymLinks(ByRef pSymLinks,ByVal pRunDrive)

    '--------------------------------------
    ' シンボリックリンクを設定していく
    ' ※キー：アプリ名、項目：アプリパス
    '--------------------------------------
    pSymLinks.Add "7-ZipPortable"             , "7-ZipPortable"             & "|" & pRunDrive & "\Tools\7-ZipPortable"
    pSymLinks.Add "CDExPortable"              , "CDExPortable"              & "|" & pRunDrive & "\Tools\CDExPortable"
    pSymLinks.Add "CPU-ZPortable"             , "CPU-ZPortable"             & "|" & pRunDrive & "\Tools\CPU-ZPortable"
    pSymLinks.Add "CrystalDiskInfoPortable"   , "CrystalDiskInfoPortable"   & "|" & pRunDrive & "\Tools\CrystalDiskInfoPortable"
    pSymLinks.Add "CrystalDiskMarkPortable"   , "CrystalDiskMarkPortable"   & "|" & pRunDrive & "\Tools\CrystalDiskMarkPortable"
    pSymLinks.Add "FastCopyPortable"          , "FastCopyPortable"          & "|" & pRunDrive & "\Tools\FastCopyPortable"
    pSymLinks.Add "GIMPPortable"              , "GIMPPortable"              & "|" & pRunDrive & "\Tools\GIMPPortable"
    pSymLinks.Add "GoogleChromePortable"      , "GoogleChromePortable"      & "|" & pRunDrive & "\Tools\GoogleChromePortable"
    pSymLinks.Add "GPU-ZPortable"             , "GPU-ZPortable"             & "|" & pRunDrive & "\Tools\GPU-ZPortable"
    pSymLinks.Add "IObitUninstallerPortable"  , "IObitUninstallerPortable"  & "|" & pRunDrive & "\Tools\IObitUninstallerPortable"
    pSymLinks.Add "IObitUnlockerPortable"     , "IObitUnlockerPortable"     & "|" & pRunDrive & "\Tools\IObitUnlockerPortable"
    pSymLinks.Add "PDFTKBuilderPortable"      , "PDFTKBuilderPortable"      & "|" & pRunDrive & "\Tools\PDFTKBuilderPortable"
    pSymLinks.Add "PDF-XChangeViewerPortable" , "PDF-XChangeViewerPortable" & "|" & pRunDrive & "\Tools\PDF-XChangeViewerPortable"
    pSymLinks.Add "ProcessExplorerPortable"   , "ProcessExplorerPortable"   & "|" & pRunDrive & "\Tools\ProcessExplorerPortable"
    pSymLinks.Add "ProcessMonitorPortable"    , "ProcessMonitorPortable"    & "|" & pRunDrive & "\Tools\ProcessMonitorPortable"
    pSymLinks.Add "SystemExplorerPortable"    , "SystemExplorerPortable"    & "|" & pRunDrive & "\Tools\SystemExplorerPortable"
    pSymLinks.Add "TeamViewerPortable"        , "TeamViewerPortable"        & "|" & pRunDrive & "\Tools\TeamViewerPortable"
    pSymLinks.Add "VLCPortable"               , "VLCPortable"               & "|" & pRunDrive & "\Tools\VLCPortable"
    pSymLinks.Add "WinMergePortable"          , "WinMergePortable"          & "|" & pRunDrive & "\Tools\WinMergePortable"
    pSymLinks.Add "wxMP3gainPortable"         , "wxMP3gainPortable"         & "|" & pRunDrive & "\Tools\wxMP3gainPortable"
    pSymLinks.Add "XnViewPortable"            , "XnViewPortable"            & "|" & pRunDrive & "\Tools\XnViewPortable"

End Function
