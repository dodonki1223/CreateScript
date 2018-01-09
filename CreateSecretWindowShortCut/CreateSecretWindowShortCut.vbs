'**************************************************************************************
'* プログラム名 ： シークレットウィンドウ（Chrome）で実行するショートカットを作成     *
'* 処理概要     ： Chromeで特定のURLをシークレットウィンドウで実行するショートカット  *
'*                 を作成する                                                         *
'*                 ※Chromeのパスはレジストリから取得するため、環境によってはChromeの *
'*                   パスが違う可能性があるため注意すること                           *
'* メモ         ： 既定のブラウザがChromeであること前提で作成しているので既定のブラウ *
'*                 ザがChromeで無い場合は正しく無いものが作成されます                 *
'* 設定         ：                                                                    *
'**************************************************************************************

Option Explicit

'*****************************************
'* 変数                                  *
'*****************************************
Dim mObjShell    : Set mObjShell  = WScript.CreateObject("WScript.Shell")       
Dim mFileInfo    : Set mFileInfo  = WScript.CreateObject("Scripting.Dictionary")

'*****************************************
'* 定数                                  *
'*****************************************
'MsgBox･InputBoxに表示する文言を設定
Dim cMsgTitle                  : cMsgTitle                  = "シークレットウィンドウで実行「lnk」ファイル作成"
Dim cMsgInputFileName          : cMsgInputFileName          = "ファイル名を入力してください。" & VbCrLf & "※ファイル名は「○○○.lnk」○○○の部分を指定してください。ファイルはデスクトップに作成されます。"
Dim cMsgInputURL               : cMsgInputURL               = "URLを入力してください。"
Dim cMsgIncorrectFileNameError : cMsgIncorrectFileNameError = "ファイル名が未入力またはファイル名として正しくありません" & VbCrLf & "処理を終了します..."
Dim cMsgIncorrectURLError      : cMsgIncorrectURLError      = "URLが未入力またはURLとして正しくありません" & VbCrLf & "処理を終了します..."

'ファイルの出力先パス
Dim cDesktopPath : cDesktopPath = mObjShell.SpecialFolders("Desktop") & "\"

'既定のブラウザのパス ※既定のブラウザがChromeであること
Dim cDefaultBrowserExePath : cDefaultBrowserExePath = GetDefaultBrowserPath()


Main()

'***********************************************************************
'* 処理名   ： メイン処理                                              *
'* 引数     ： なし                                                    *
'* 処理内容 ： メイン処理                                              *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub Main()

    '------------------------------------
    ' ファイル名を取得                   
    '------------------------------------
    'ユーザーに対話しファイル名を取得
    Dim mFileName : mFileName = InputBox(cMsgInputFileName, cMsgTitle)
    
    'ファイル名が正しいかどうかを取得
    Dim mIsCorrectFileName : mIsCorrectFileName = IsCorrectRegExpMatch("(\\|\/|\:|\*|\?|\""|\<|\>|\|)", mFileName)

    '未入力またはファイル名が正しくない場合は処理を終了
    If mIsCorrectFileName  = True Or mFileName = "" Then
    
        MsgBox cMsgIncorrectFileNameError, vbOKOnly, cMsgTitle
        Wscript.Quit()
        
    End If
    
    'ファイル名の末尾に「.lnk」を追加
    mFileName = mFileName & ".lnk"

    '------------------------------------
    ' URLを取得                          
    '------------------------------------
    'ユーザーに対話（URL）
    Dim mURL : mURL = InputBox(cMsgInputURL, cMsgTitle)

    'URLが正しいかどうかを取得
    Dim mIsCorrectURL : mIsCorrectURL = IsCorrectRegExpMatch("^(https*|ftp)://[-_!~';:@&=,%#/a-zA-Z0-9\$\*\+\?\.\(\)]+$", mURL)

    '未入力またはURLが正しくない場合は処理を終了
    If mIsCorrectURL = False Or mURL = "" Then
    
        MsgBox cMsgIncorrectURLError, vbOKOnly, cMsgTitle
        Wscript.Quit()
        
    End If

    '------------------------------------
    ' シークレットウィンドウ用のコマンドライン引数を作成
    '------------------------------------
    '--incognito + "表示したいURL"
    Dim mCommandLineArg : mCommandLineArg = "--incognito " & """" & mURL & """"

    '------------------------------------
    ' ショートカット作成処理             
    '------------------------------------
    'ファイル情報を設定('キー：ファイル名、作成情報：ファイルパス|出力先フォルダ)
                  'ファイル名  'ショートカット先               'ファイルの出力先         'コマンドライン引数    'アイコン
    mFileInfo.Add mFileName ,   cDefaultBrowserExePath & "|" &  cDesktopPath      & "|" & mCommandLineArg & "|" & cDefaultBrowserExePath
    
    'ショートカット作成処理
    Dim key
    For Each key In mFileInfo.Keys

        '作成情報を切り分ける（①ショートカットパス、②出力フォルダ、③コマンドライン引数、④アイコン）
        Dim mAryFileInfo : mAryFileInfo = Split(mFileInfo(key),"|")
    
        'lnkファイルのファイル名からショートカットを作成するフルパスを取得
        Dim mShortCutFileName     : mShortCutFileName     = key
        Dim mShortCutFileFullPath : mShortCutFileFullPath = mAryFileInfo(1) & mShortCutFileName
    
        'ショートカットを作成ファイルの保存先フォルダが無かった場合はフォルダを作成する
        CreateNotExistFolder(mShortCutFileFullPath)
    
        'ショートカットオブジェクトを作成し出力先パス、コマンドライン引数、アイコンを指定
        Dim mShortCut : Set mShortCut = mObjShell.CreateShortcut(mShortCutFileFullPath) 'ショートカットオブジェクトを作成
        mShortCut.TargetPath                                    = mAryFileInfo(0)       'ショートカット先       ※Chromeのパスをセット
        If UBound(mAryFileInfo) > 1 Then mShortCut.Arguments    = mAryFileInfo(2)       'コマンドライン引数設定
        If UBound(mAryFileInfo) > 2 Then mShortCut.IconLocation = mAryFileInfo(3)       'アイコン情報を設定
    
        'ショートカットを作成
        mShortCut.Save
    
    Next

    'オブジェクトの破棄
    Set mObjShell  = Nothing
    Set mFileInfo  = Nothing

End Sub

'***********************************************************************
'* 処理名   ： 既定のブラウザのパスを取得                              *
'* 引数     ： なし                                                    *
'* 処理内容 ： レジストリから既定のブラウザのパス文字列を取得し返す    *
'* 戻り値   ： 既定のブラウザのパス                                    *
'***********************************************************************
Function GetDefaultBrowserPath()

    '既定のブラウザで開くExeのパスをレジストリから探すキー
    '※対象のキーがいくつか存在するが今回は「１」から探すように作成する
    '  １：HKEY_CLASSES_ROOT\http\shell\open\command\
    '  ２：HKEY_CLASSES_ROOT\https\shell\open\command\
    '  ３：HKEY_CURRENT_USER\Software\Classes\http\shell\open\command\
    '  ４：HKEY_CURRENT_USER\Software\Classes\https\shell\open\command\
    Dim cRegRunHttpKey : cRegRunHttpKey = "HKEY_CLASSES_ROOT\HTTP\shell\open\command\"

    '「WScript.Shell」のオブジェクトを作成
    Dim mObjShell : Set mObjShell = WScript.CreateObject("WScript.Shell")       

    'レジストリに設定されている値を取得する
    '※「"ブラウザパス" -- "%1"」の形式で取得される
    Dim mDefaultBrowserValue   : mDefaultBrowserValue = mObjShell.RegRead(cRegRunHttpKey)

    'レジストリに設定されている値からExeのパスを取得する
    '※「"ブラウザパス" -- "%1"」の形式から「"ブラウザパス"」のみを抜き出す
    Dim mDefaultBrowserExePath : mDefaultBrowserExePath = Left(mDefaultBrowserValue, InStr(mDefaultBrowserValue, ".exe") + 4)

    '作成したオブジェクトを破棄
    Set mObjShell = Nothing

    '返り値を設定（ダブルクオーテーションを削除）
    GetDefaultBrowserPath = Replace(mDefaultBrowserExePath, """", "")

End Function

'***********************************************************************
'* 処理名   ： 正規表現に一致するか                                    *
'* 引数     ： pPattern 正規表現パターン                               *
'*             pString  対象文字列                                     *
'* 処理内容 ： パターンに一致（正規表現でチェック）するかどうか        *
'* 戻り値   ： パターンに一致：True、パターンに不一致：False           *
'***********************************************************************
Function IsCorrectRegExpMatch(pPattern, pString)

    '----------------------------------
    ' 正規表現オブジェクトを作成       
    '----------------------------------
    Dim mRegExp : Set mRegExp = New RegExp
    mRegExp.Pattern    = pPattern '正規表現のパターンを設定
    mRegExp.IgnoreCase = True     '大文字・小文字を区別しないように設定
    mRegExp.Global     = True     '文字列を全体を検索するように設定
    
    '----------------------------------
    ' 正規表現に一致するか結果を取得
    '----------------------------------
    '正規表現パターンに一致した時
    If mRegExp.test(pString) Then ' 検索をテストします。
    
        IsCorrectRegExpMatch = True
        
    Else
    
        IsCorrectRegExpMatch = False
        
    End If
    
    '----------------------------------
    ' 作成したオブジェクトを破棄
    '----------------------------------
    Set mRegExp = Nothing
    
End Function

'***********************************************************************
'* 処理名   ： フォルダ作成処理                                        *
'* 引数     ： pPath 対象パス                                          *
'* 処理内容 ： 対象パスに存在しないパスがいたら作成する                *
'*             ※フォルダの作成は再帰的に行います                      *
'*               ドライブ → ドライブ\階層１→ドライブ\階層１\階層２\  *
'*               → ドライブ\階層１\階層２\対象フォルダ                *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Function CreateNotExistFolder(pPath)
    
    Dim mObjFso       : Set mObjFso   = WScript.CreateObject("Scripting.FileSystemObject")
    Dim mDriveName    : mDriveName    = Left(mObjFso.GetDriveName(pPath),2)                'ドライブ名を取得
    Dim mParentFolder : mParentFolder = mObjFso.GetParentFolderName(pPath)                 '親フォルダー名を取得
    
    '対象のドライブが存在する時
    If mObjFso.DriveExists(mDriveName) Then
    
        'Driveオブジェクトを作成
        Dim mObjDrive : Set mObjDrive = mObjFso.GetDrive(mDriveName) 
        
    Else
    
        Exit Function
        
    End If
    
    'ドライブの準備ができている時
    If mObjDrive.IsReady Then
    
        '拡張子文字列が取得出来た場合(ファイルの時)
        If Len(mObjFso.GetExtensionName(pPath)) > 0 Then 
        
            '親フォルダーが存在しない時、対象パスから親フォルダー作成する（再帰的）
            If Not(mObjFso.FolderExists(mParentFolder)) Then CreateNotExistFolder(mParentFolder)
            
        Else
        
            '対象フォルダーが存在しない時
            If Not(mObjFso.FolderExists(pPath)) Then
            
                '親フォルダーを作成後、対象フォルダーを作成（再帰的）
                CreateNotExistFolder(mParentFolder)
                mObjFso.CreateFolder(pPath)
                
            End If
            
        End If
        
    End If
    
    'オブジェクトの破棄
    Set mObjFso   = Nothing
    Set mObjDrive = Nothing

End Function