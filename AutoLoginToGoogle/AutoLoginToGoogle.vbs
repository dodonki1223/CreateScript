'**************************************************************************************
'* プログラム名 ： Googleアカウント自動ログインスクリプト                             *
'* 処理概要     ： InternetExplorerを使い、Googleアカウントに自動ログインを行います   *
'*                   ①IEを起動                                                       *
'*                   ②Googleアカウントをログアウト                                   *
'*                       ※Googleにログインしていた場合はログアウトします             *
'*                   ③Googleアカウントにログイン                                     *
'*                   ④ログイン失敗処理                                               *
'*                       ※ログインに失敗した時はメッセージを表示させ処理を終了します *
'*                   ⑤対象のURLに遷移                                                *
'* メモ         ： ※2017/08/25現在の仕様で作成                                       *
'*                 起動中のIEがあった場合、アクティブタブで処理が行われ、現在開いて   *
'*                 いるページ情報が失われてしまう。対応策としてアクティブタブを新し   *
'*                 いタブで開き直し、元々のアクティブタブでGoogleの自動ログイン処理   *
'*                 を行う                                                             *
'* 設定         ： Googleアカウントのメールアドレスとパスワードの情報を設定すること   *
'*                 ※設定をしないと正しく動きません！！                               *
'**************************************************************************************
Option Explicit

'*****************************************
'* 定数                                  *
'*****************************************
'URL
Dim GoogleTopURL     : GoogleTopURL     = "https://www.google.co.jp/"
Dim LoginURL         : LoginURL         = "https://www.google.com/accounts/Login?hl=ja&continue=http://www.google.co.jp/"
Dim TransferURL      : TransferURL      = "https://drive.google.com/drive/my-drive"                                       'ログイン後の表示ページはGoogleドライブ

'Googleアカウント情報
'※Googleのアカウント情報を設定すること。設定しないと動きません。
Dim LoginMailAddress          : LoginMailAddress          = "Googleアカウントメールアドレス 例：sample@sanmple.co.jp "
Dim LoginPassword             : LoginPassword             = "Googleアカウントのパスワード   例：PassW0rd             "

'パスワードの入力バージョン
'※処理を変更する場合はLoginPassword変数の処理を上記と入れ替えること
'Dim InputLoginPasswordTitle   : InputLoginPasswordTitle   = "Googleアカウントログインパスワード入力"
'Dim InputLoginPasswordMessage : InputLoginPasswordMessage = "Googleアカウントのログインパスワードを入力してください"
'Dim LoginPassword             : LoginPassword             = InputBox(InputLoginPasswordMessage, InputLoginPasswordTitle)

'メッセージボックス
Dim MsgBoxTitle               : MsgBoxTitle               = "Google自動ログイン"
Dim MsgBoxTransferContent     : MsgBoxTransferContent     = "Googleドライブに移動します" & vbCrLf & "よろしいですか？"
Dim MsgBoxLoginFailureContent : MsgBoxLoginFailureContent = "パスワードが違います！！" & VbCrLf & "処理を終了します..."

'ログイン処理表示区分 True:処理を表示、False:処理を表示しない
'※デフォルトは処理を表示する
Dim RunningVisible   : RunningVisible   = True

Main()

'***********************************************************************
'* 処理名   ： メイン処理                                              *
'* 引数     ： なし                                                    *
'* 処理内容 ： メイン処理                                              *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub Main()

    '*****************************************
    '* IE起動処理                            *
    '*****************************************
    '起動中のIEを取得
    Dim mObjIE : Set mObjIE = GetRunningIEObject()
    
    'IEが既に起動中かのフラグ
    Dim mIsRunningIE : mIsRunningIE = True
    
    '起動中のIEが存在しなかったら
    If mObjIE Is Nothing Then 
    
        Set mObjIE = WScript.CreateObject("InternetExplorer.Application") 'IEオブジェクトを作成する
        mIsRunningIE = False                                              '起動中フラグをOFF
        mObjIE.Visible = RunningVisible
        
    End If
    
    'IEが起動中だったら現在のタブを新しいタブで開く
    If mIsRunningIE = True Then TransferPage mObjIE, mObjIE.LocationURL, "1"
    
    '*****************************************
    '* Googleアカウントログアウト処理        *
    '*****************************************
    'GoogleのTopページへ遷移
    TransferPage mObjIE, GoogleTopURL, "0"
    
    'ページ内に「ログアウト」という文字列を含んだ「a」タグがある時
    '※ログアウトボタンが存在したら
    If IsExistTag(mObjIE, "tag", "a", "ログアウト") = True Then
    
        'ログアウトボタンのクリック処理を実行
        Dim mLogOutButton : Set mLogOutButton = mObjIE.Document.getElementById("gb_71")
        mLogOutButton.click()  
        WaitIe mObjIE, 1000   'ページが読み込み終わるまで待つ
        
    End If
    
    '*****************************************
    '* Googleアカウントログイン処理          *
    '*****************************************
    'Googleのログインページへ遷移
    TransferPage mObjIE, LoginURL, "0"
    
    'ページ内に「別のアカウントを使用」という文字列を含んだ「p」タグがある時
    '※アカウントが複数の時でアカウント選択画面の時の処理
    If IsExistTag(mObjIE, "tag", "p", "別のアカウントを使用") = True Then
    
        '別のアカウントを使用のクリック処理
        Dim mOtherAccount : Set mOtherAccount = mObjIE.Document.getElementById("identifierLink")
        mOtherAccount.click()
        WaitIe mObjIE, 1000 'ページが読み込み終わるまで待つ
        
    End If
    
    'ログインユーザーを設定（メールアドレスの設定）
    Dim mLoginUser : Set mLoginUser = mObjIE.Document.getElementsByName("identifier")(0)
    mLoginUser.Value = LoginMailAddress
    
    'ページ内に「次へ」という文字列を含んだ「RveJvd snByac」Classを持つタグがある時
    '※Googleアカウントを持っていない時の処理
    If IsExistTag(mObjIE, "class", "RveJvd snByac", "次へ") = True Then
    
        '次へボタンのクリック処理
        Dim mNextButton : Set mNextButton = mObjIE.Document.getElementsByClassName("ZFr60d CeoRYc")(0)
        mNextButton.click()
    
        'ページが切り替わるまで繰り返す
        Dim mIsPageChange : mIsPageChange = False 'ページ切り替えフラグ
        While mIsPageChange = False
        
            'ページ内に「パスワードを入力」という文字列を含んだ「AxOyFc snByac」Classを持つタグがある時
            '※パスワードタグが存在したらページが切り替わったこととする
            If IsExistTag(mObjIE, "class", "AxOyFc snByac", "パスワードを入力") = True Then
            
                mIsPageChange = True
                
            End If
            
            WScript.Sleep(1000)  '１秒待つ
            
        Wend
        
    End If
    
    'ログインパスワードを設定
    Dim mLoginPassword : Set mLoginPassword = mObjIE.Document.getElementsByName("password")(0)
    mLoginPassword.Value = LoginPassword
    
    'ログインボタンのクリック処理
    Dim mSignInButton : Set mSignInButton = mObjIE.Document.getElementsByClassName("ZFr60d CeoRYc")(0)
    mSignInButton.click() 
    WaitIe mObjIE, 1000  'ページが読み込み終わるまで待つ
    
    'IEが起動中ではなくかつログイン処理表示区分がFalseの時、IEを可視化
    If mIsRunningIE = False And RunningVisible = False Then mObjIE.Visible = True

    '*****************************************
    '* ログイン失敗時処理                    *
    '*****************************************
    'ページ内に「パスワードが正しくありません。もう一度お試しください。」という文字列を含んだ「dEOOab RxsGPe」Classを持つタグがある時
    If IsExistTag(mObjIE, "class", "dEOOab RxsGPe", "パスワードが正しくありません。もう一度お試しください。") = True Then
    
        'メッセージを表示して処理を終了
        MsgBox MsgBoxLoginFailureContent , vbOKOnly, MsgBoxTitle
        Wscript.Quit()
        
    End If

    '*****************************************
    '* 対象URLへの遷移処理                   *
    '*****************************************
    '対象URLへ遷移するかユーザーに対話
    Dim mMsgResult : mMsgResult = MsgBox(MsgBoxTransferContent, vbOKCancel, MsgBoxTitle)
    
    'キャンセルが押された時は処理を終了
    If mMsgResult = vbCancel Then Wscript.Quit()
    
    '対象URLへ遷移
    TransferPage mObjIE, TransferURL, "0"

End Sub

'***********************************************************************
'* 処理名   ： ページ遷移処理                                          *
'* 引数     ： pObjIE       IEオブジェクト                             *
'*             pTransferURL 遷移先URL                                  *
'*             pNewTabKbn   新しいタブで開くかの区分                   *
'*               0  現在のタブで開く                                   *
'*               1  新しいタブで開く                                   *
'* 処理内容 ： ページ遷移処理を行い、終了するまで待機                  *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub TransferPage(pObjIE, pTransferURL, pNewTabKbn)
    
    'ページの遷移処理（新しいタブで開くかの区分により処理を分岐）
      Select Case pNewTabKbn
      
        Case "0"   '現在のタブで開く
        
              pObjIE.Navigate(pTransferURL)
              
        Case "1"   '新しいタブで開く
        
              pObjIE.Navigate2 pTransferURL, 2048
              
          Case Else   'どれにも当てはまらない（現在のタブで開くと同じとする）
              
              pObjIE.Navigate(pTransferURL)
              
      End Select
    
    'ページが読み込み終わるまで待つ
    WaitIe pObjIE, 1000  
    
End Sub

'***********************************************************************
'* 処理名   ： IE待機                                                  *
'* 引数     ： pObjIE     IEオブジェクト                               *
'*             pWaitTime  待機時間                                     *
'*               目安、1000で1秒です                                   *
'* 処理内容 ： IEオブジェクトが起動中またはページが読込中の場合は待機  *
'*             時間分待つ処理を繰り返す。IEオブジェクトが起動しページ  *
'*             の読み込み中で無くなったら処理を終了                    *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub WaitIe(pObjIE, pWaitTime)

    '1秒待機
    WScript.Sleep(1000)

    '「IEが起動し終わる」か「ページが読み込み終わる」まで繰り返す
    While pObjIE.ReadyState <> 4 Or pObjIE.Busy = True
    
        WScript.Sleep(pWaitTime) '待機時間分待つ
        
    Wend
    
End Sub

'***********************************************************************
'* 処理名   ： 起動中IEオブジェクト取得                                *
'* 引数     ： なし                                                    *
'* 処理内容 ： 現在起動中のウインドウを取得し、IEが起動中ならそのオブ  *
'*             ジェクトを返し、起動中でなかったらNothingを返す         *
'* 戻り値   ： IEオブジェクト / Nothing                                *
'***********************************************************************
Function GetRunningIEObject()

    Dim mObjShell : Set mObjShell = CreateObject("Shell.Application")
    Dim mObjFso   : Set mObjFso   = CreateObject("Scripting.FileSystemObject")
    Dim mObjWindow

    '戻り値に初期値を設定
    Set GetRunningIEObject = Nothing

    '起動中のウインドウ数分繰り返す
    For Each mObjWindow In mObjShell.Windows

        'ウインドウのFullNameを小文字に変換した文字列が「iexplore.exe」の時
        '※ウインドウがIEの時
        If LCase(mObjFso.GetFileName(mObjWindow.FullName)) = "iexplore.exe" Then

            'IEオブジェクトを戻り値にセット
            Set GetRunningIEObject = mObjWindow
            Exit For

        End If

    Next

    Set mObjShell = Nothing
    Set mObjFso   = Nothing
    
End Function

'***********************************************************************
'* 処理名   ： タグの存在チェック                                      *
'* 引数     ： pObjIE       IEオブジェクト                             *
'*             pMethodType  タグの検索するタイプ                       *
'*               ※name,class,tagのいずれか                            *
'*             pElementName タグを検索する文字列                       *
'*             pKeywords    タグ内から検索する文字列                   *
'* 処理内容 ： 表示中のページから「検索するタイプ」、「検索する文字列」*
'*             を使って取得出来たタグのうち「タグ内から検索する文字列」*
'*             を含むタグがあるかチェックする                          *
'*             ※処理内容が意味不明になってしまった（泣）              *
'* 戻り値   ： True:存在する / False:存在しない                        *
'***********************************************************************
Function IsExistTag(pObjIE, pMethodType, pElementName, pKeywords)

    Dim mObjDoc, mMyDoc
  
    '戻り値に初期値をセット
    IsExistTag = False

    'タグの検索するタイプによりオブジェクトの作成方法を変える
      Select Case pMethodType
      
        Case "name"
        
              Set mObjDoc = pObjIE.document.getElementsByName(pElementName)
              
        Case "class"
        
              Set mObjDoc = pObjIE.document.getElementsByClassName(pElementName)
              
        Case "tag"
        
              Set mObjDoc = pObjIE.document.getElementsByTagName(pElementName)
              
      End Select

    '見つかったオブジェクト数分繰り返す
    For Each mMyDoc In mObjDoc
    
        'オブジェクト内に「タグ内から検索する文字列」があった時
        If InStr(mMyDoc.outerHTML, pKeywords) > 0 Then
        
            IsExistTag = True
            Exit For
            
        End If
        
    Next
    
End Function