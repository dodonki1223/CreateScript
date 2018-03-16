'**************************************************************************************
'* プログラム名 ： デスクトップの壁紙変更スクリプト                                   *
'* 処理概要     ： レジストリを使用しデスクトップの壁紙情報を書き換える。             *
'*                 書き換え後、デスクトップの壁紙の情報をメッセージボックスにて表示、 *
'*                 壁紙のパスをクリップボードに保存します。                           *
'* メモ         ：                                                                    *
'**************************************************************************************

Option Explicit

'--------------------------------------
' 定数
'--------------------------------------
'変更する壁紙のパス設定
Dim cWallPaperPath : cWallPaperPath = "設定したい壁紙のフルパス"

'レジストリを操作するためのWMIメソッドを呼び出すために使用するルートキー
Dim cHKEY_CURRENT_USER : cHKEY_CURRENT_USER = &H80000001

'レジストリに作成するキー情報
Dim cCreateRegistryKey : cCreateRegistryKey = "Control Panel\Desktop"

'レジストリに値をセットする名前
Dim cSetRegistryValueName : cSetRegistryValueName = "Wallpaper"

'レジストリの内容を画面に即時反映させるための処理を実行する回数
Dim cRunRegistryReflectCount : cRunRegistryReflectCount = 10

'レジストリの内容を画面に即時反映させるための命令
Dim cRegistryReflectOrder : cRegistryReflectOrder    = "RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters "

'--------------------------------------
' 変数宣言・インスタンス作成
'--------------------------------------
Dim objShell   : Set objShell   = WScript.CreateObject("WScript.Shell")   'WScript.Shellオブジェクト

'メイン処理を実行
Main()

'***********************************************************************
'* 処理名   ： メイン処理                                              *
'* 引数     ： なし                                                    *
'* 処理内容 ： メイン処理                                              *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub Main()

    '--------------------------------------
    ' コンピュータ名を取得
    '--------------------------------------
    'コンピュータ情報からコンピュータ名を取得
    Dim mComputerInfo : Set mComputerInfo = GetComputerInfo()
    Dim mComputerName : mComputerName     = mComputerInfo.Name

    '--------------------------------------
    ' レジストリに壁紙の情報をセット
    '--------------------------------------
    'レジストリを取得
    Dim mObjRegistry  : Set mObjRegistry = GetObject("winmgmts:\\" & mComputerName & "\root\default:StdRegProv")
    
    '壁紙をセットする用のレジストリを作成する
    mObjRegistry.CreateKey cHKEY_CURRENT_USER, cCreateRegistryKey

    'レジストリにデスクトップの壁紙情報セットする
    mObjRegistry.SetStringValue cHKEY_CURRENT_USER, cCreateRegistryKey, cSetRegistryValueName, cWallPaperPath
    
    Set mObjRegistry = Nothing

    '--------------------------------------
    ' レジストリの内容を画面に反映させる
    '--------------------------------------
    '複数回実行しないと反映されないため何回も実行する
    Dim i
    For i = 1 To cRunRegistryReflectCount

        objShell.Run cRegistryReflectOrder, 1, true
         WScript.Sleep(1000)                         '１秒待つ

    Next

    '--------------------------------------
    ' ユーザー名を取得
    '--------------------------------------
    'WScript.Networkオブジェクトを作成
    Dim mObjNetWork : Set mObjNetWork = WScript.CreateObject("WScript.Network")
    
    'ログインユーザー名を作成（コンピュータ名 ＋ \ ＋ ユーザー名）
    Dim mLoginUserName : mLoginUserName = mComputerName & "\" & mObjNetWork.UserName
    
    Set mObjNetWork= Nothing
    
    '--------------------------------------
    ' 壁紙の情報を画面に表示
    '--------------------------------------
    'ローカルコンピュータに接続 
    '※ネットワークのパソコンにアクセスする時はユーザー名とパスワードをセットすること
    '  サンプルコード：http://www.wmifun.net/step/vb6_06.html
    Dim mLocator : Set mLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
    Dim mService : Set mService = mLocator.ConnectServer
    
    'デスクトップ情報を取得
    Dim mDesktopInfos : Set mDesktopInfos = mService.ExecQuery("Select * From Win32_Desktop")

    'ユーザーごと繰り返される
    Dim mDesktopInfo, mWallPaperPath
    For Each mDesktopInfo In mDesktopInfos

        'ユーザー名がログインしているユーザー名と一致した時
        If mLoginUserName = mDesktopInfo.Name Then

            '壁紙のパスを取得
            mWallPaperPath = mDesktopInfo.Wallpaper
            Exit For
        
        End If
    
    Next

    '--------------------------------------
    ' 壁紙のパスをクリップボードにコピー
    '--------------------------------------
    Dim mClipBoardCopyText : mClipBoardCopyText = "cmd.exe /c ""echo " & mWallPaperPath & "| clip"""
    objShell.Run mClipBoardCopyText, 0

    '壁紙の情報を表示
    MsgBox "設定されている壁紙のパスは以下の通りです" & _
           vbCrLf & _
           vbCrLf & _
           mWallPaperPath & _
           vbCrLf & _
           vbCrLf & _
           "※壁紙のパスはクリップボードにコピーされました"

    '--------------------------------------
    ' オブジェクトの破棄処理
    '--------------------------------------
    Set objShell = Nothing    

End Sub

'***********************************************************************
'* 処理名   ： コンピュータの情報を取得                                *
'* 引数     ： なし                                                    *
'* 処理内容 ： WMIを使用しコンピュータの情報を取得する                 *
'*               Domain             ：ドメイン(ワークグループ)         *
'*               Manufacturer       ：メーカー                         *
'*               Model              ：型番                             *
'*               Name               ：コンピュータ名                   *
'*               PrimaryOwnerName   ：所有者(初期セットアップ時に作成  *
'*                                    したアカウント名)                *
'*               TotalPhysicalMemory：メモリ数                         *
'* 戻り値   ： コンピュータ情報を保持したObject                        *
'***********************************************************************
Function GetComputerInfo()

   'WMIからコンピュータの情報を取得する
    Dim mObjWMIService : Set mObjWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\." & "\root\cimv2")
    Dim mSystemInfos   : Set mSystemInfos  = mObjWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    
    'コンピュータの情報をセット
    Dim mSystemInfo
    For Each mSystemInfo In mSystemInfos
    
        Set GetComputerInfo = mSystemInfo
        
    Next

End Function
