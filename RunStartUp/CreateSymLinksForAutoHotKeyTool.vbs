'--------------------------------------
' 定数宣言
'--------------------------------------
Const AUTO_HOT_KEY_TOOL_TOOLS_DIRECTORY = "\Tools\AutoHotKey\Tools\"

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
        Dim symLinkPath       : symLinkPath = runDrive & AUTO_HOT_KEY_TOOL_TOOLS_DIRECTORY & arySymLinks(0)
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
    'Tools フォルダ内のシンボリックリンクの作成
    pSymLinks.Add "Explorers"         , "Explorers"               & "|" & pRunDrive & "\Tools\Explorers"
    pSymLinks.Add "FolderFileList"    , "FolderFileList"          & "|" & pRunDrive & "\Tools\FolderFileList"
    pSymLinks.Add "ImageForClipboard" , "ImageForClipboard"       & "|" & pRunDrive & "\Tools\ImageForClipboard"
    pSymLinks.Add "MyAutoHotKeySpy"    , "MyAutoHotKeySpy"        & "|" & pRunDrive & "\Tools\MyAutoHotKeySpy"
    pSymLinks.Add "ShutDownDialog"    , "ShutDownDialog"          & "|" & pRunDrive & "\Tools\ShutDownDialog"

    'SelfMadeMenuフォルダ内のシンボリックリンクの作成
    pSymLinks.Add "Components"        , "SelfMadeMenu\Components" & "|" & pRunDrive & "\Tools\AutoHotKey\Components"
    pSymLinks.Add "Icon"              , "SelfMadeMenu\Icon"       & "|" & pRunDrive & "\Tools\AutoHotKey\Icon"
    pSymLinks.Add "Tools"             , "SelfMadeMenu\Tools"      & "|" & pRunDrive & "\Tools\AutoHotKey\Tools"

End Function
