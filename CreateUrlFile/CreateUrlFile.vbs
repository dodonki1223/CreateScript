'**************************************************************************************
'* プログラム名 ： urlファイル作成スクリプト                                          *
'* 処理概要     ： ファイル名と対象URLを入力して「ファイル名.url」ファイル作成する    *
'* メモ         ：                                                                    *
'* 設定         ：                                                                    *
'**************************************************************************************

Option Explicit

'*****************************************
'* 定数                                  *
'*****************************************
'InputBoxに表示する文言を設定
Dim mTitle           : mTitle           = "URLファイルの作成"
Dim mFileNameMessage : mFileNameMessage = "ファイル名を入力してください。" & VbCrLf & VbCrLf & "※ファイル名は「○○○.url」○○○の部分を指定してください。ファイルはデスクトップに作成されます。"
Dim mURLMessage      : mURLMessage      = "URLを入力してください。"

'入力内容が正しくない時に表示するメッセージ
Dim mIncorrectFileNameError : mIncorrectFileNameError = "ファイル名が未入力またはファイル名として正しくありません" & VbCrLf & VbCrLf & "処理を終了します..."
Dim mIncorrectURLError : mIncorrectURLError = "URLが未入力またはURLとして正しくありません" & VbCrLf & VbCrLf & "処理を終了します..."

'*****************************************
'* 変数                                  *
'*****************************************
Dim mObjShell    : Set mObjShell  = WScript.CreateObject("WScript.Shell")        
Dim mFileInfo    : Set mFileInfo  = WScript.CreateObject("Scripting.Dictionary")

'ファイルの出力先パスを取得
Dim mDesktopPath : mDesktopPath = mObjShell.SpecialFolders("Desktop") & "\"    

Main()

'***********************************************************************
'* 処理名   ： メイン処理                                              *
'* 引数     ： なし                                                    *
'* 処理内容 ： メイン処理                                              *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub Main()

    '*****************************************
    '* ファイル名取得                        *
    '*****************************************
    'ユーザーに対話しファイル名を取得
    Dim mFileName : mFileName = InputBox(mFileNameMessage, mTitle)
    
    '*****************************************
    '* ファイル名入力チェック                *
    '* ※未入力＆ファイル名妥当性チェック    *
    '*****************************************
    'ファイル名妥当性チェック用RegExp設定
    '※ファイル名に使えない文字の一覧:\ / : * ? " < > |
    '  エスケープ文字は「\」、区切り文字は「|」です
    Dim mObjFileNameRegExp : Set mObjFileNameRegExp = New RegExp
    mObjFileNameRegExp.pattern = "(\\|\/|\:|\*|\?|\""|\<|\>|\|)"
    mObjFileNameRegExp.Global  = True

    'ファイル名、妥当性チェック結果を取得
    Dim mIsFileName : mIsFileName = mObjFileNameRegExp.Test(mFileName)

    '未入力またはファイル名が正しくない場合は処理を終了
    If mIsFileName = True Or mFileName = "" Then
    
        MsgBox mIncorrectFileNameError
        Wscript.Quit()
        
    End If
    
    'ファイル名の末尾に「.url」を追加
    mFileName = mFileName & ".url"

    '*****************************************
    '* URL取得                        *
    '*****************************************
    'ユーザーに対話（URL）
    Dim mURL : mURL = InputBox(mURLMessage, mTitle)

    '*****************************************
    '* URL入力チェック                       *
    '* ※未入力＆URL妥当性チェック           *
    '*****************************************
    'URL妥当性チェック用RegExp設定
    Dim mObjURLRegExp : Set mObjURLRegExp = New RegExp
    mObjURLRegExp.pattern = "^(https*|ftp)://[-_!~';:@&=,%#/a-zA-Z0-9\$\*\+\?\.\(\)]+$"
    mObjFileNameRegExp.Global  = True

    'URL妥当性チェック結果を取得
    Dim mIsUrl : mIsUrl = mObjURLRegExp.Test(mURL)

    '未入力またはURLが正しくない場合は処理を終了
    If mIsUrl = False Or mURL = "" Then
    
        MsgBox mIncorrectURLError
        Wscript.Quit()
        
    End If

    '*****************************************
    '* ショートカット作成処理                *
    '*****************************************
    'ファイル情報を設定('キー：ファイル名、作成情報：ファイルパス|出力先フォルダ)
                  'ファイル名  'ショートカット先        'ファイルの出力先
    mFileInfo.Add mFileName ,  mURL              & "|" & mDesktopPath 
    
    'ショートカット作成処理
    Dim key
    For Each key In mFileInfo.Keys

        '作成情報を切り分ける（①ショートカットパス、②出力フォルダ）
        Dim mAryFileInfo : mAryFileInfo = Split(mFileInfo(key),"|")
    
        'urlファイルのファイル名からショートカットを作成するディレクトリを取得
        Dim FileName : FileName = key
        Dim Path     : Path     = mAryFileInfo(1) & FileName
    
        '作成先ディレクトリのフォルダが無かった場合はフォルダを作成する
        CreateNotExistFolder(Path)
    
        'ショートカットオブジェクトを作成し出力先パス、コマンドライン引数
        Dim mShortCut : Set mShortCut = mObjShell.CreateShortcut(Path) 'ショートカットオブジェクトを作成
        mShortCut.TargetPath = mAryFileInfo(0)                         'ショートカット先
    
        'ショートカットを作成
        mShortCut.Save
    
    Next

    'オブジェクトの破棄
    Set mObjShell  = Nothing
    Set mFileInfo  = Nothing

End Sub

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