'**************************************************************************************
'* プログラム名 ： 時刻同期スクリプト                                                 *
'* 処理概要     ： TimeServerから現在日付・時刻情報を取得しパソコンの日付と時刻の設定 *
'*                 を行います                                                         *
'*                   ①管理者として実行                                               *
'*                   ②TimeServerにアクセスし現在日付・時刻情報を取得                 *
'*                   ③TimeServerから取得した現在日付・時刻情報から経過日数、時刻情報 *
'*                     を取得（1900/1/1からの経過時間用の情報を作成する）             *
'*                   ④日付・時刻設定をコマンドで行う                                 *
'* メモ         ： 参考URL：http://35huyu.seesaa.net/article/382978581.html           *
'* 設定         ：                                                                    *
'**************************************************************************************
Option Explicit

'*****************************************
'* 定数                                  *
'*****************************************
'タイムサーバーアドレス
Dim cTimeSverAddress : cTimeSverAddress = "https://ntp-a1.nict.go.jp/cgi-bin/ntp"

'HTTPステータスコード２００（リクエストの成功）
Dim cHttpStatusCode200 : cHttpStatusCode200 = 200

'時間関連の定数
Dim cMinuteSecond         : cMinuteSecond         = 60                            '１分間の秒数
Dim cHourSecond           : cHourSecond           = cMinuteSecond * cMinuteSecond '１時間分の秒数 ※６０秒 * ６０分
Dim cDaySecond            : cDaySecond            = cHourSecond * 24              '１日分の秒数   ※１時間分の秒数 * ２４時間
Dim cTimeDifferenceForGMT : cTimeDifferenceForGMT = 9                             '日本／東京とグリニッジ標準時間との時差
Dim cCalculationMinDay    : cCalculationMinDay    = "1900/1/1"                    '計算最小日付

'メッセージボックス関連
Dim cMsgBoxTitle                     : cMsgBoxTitle                     = "時刻同期スクリプト"
Dim cMsgBoxTimeServerConnectionError : cMsgBoxTimeServerConnectionError = "タイムサーバーに接続できませんでした。" & vbCrLf & "処理を終了します。"
Dim cMsgBoxProcessSucess             : cMsgBoxProcessSucess             = "時刻の同期に成功しました。"

Main()

'***********************************************************************
'* 処理名   ： メイン処理                                              *
'* 引数     ： なし                                                    *
'* 処理内容 ： メイン処理                                              *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub Main()

    '----------------------------------------
    ' 管理者として実行
    '----------------------------------------
    RunAdministrator()

    '----------------------------------------
    ' TimeServerから現在日付・時刻情報を取得
    '----------------------------------------
    '「WinHttpRequest」COMオブジェクトを作成
    Dim mWinHttpRequest : Set mWinHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")

    'HTTPリクエストの送信
    mWinHttpRequest.Open "GET", cTimeSverAddress, false
    mWinHttpRequest.Send

    'HTTPステータスコードが２００以外の時
    If mWinHttpRequest.Status <> 200 Then
    
        'タイムサーバーに接続できませんエラーを表示
        MsgBox cMsgBoxTimeServerConnectionError, vbOKOnly, cMsgBoxTitle
        Exit Sub
        
    End If
    
    'レスポンスボディの内容を取得する ※html文で取得される
    Dim mResponseText : mResponseText = mWinHttpRequest.ResponseText

    'レスポンスボディの内容から1900/1/1から現在までの秒数を取得（グリニッジ標準時間で取得される）
    Dim mSecondToNowDate : mSecondToNowDate = GetTextForRegExpMatch("[0-9]*\.[0-9]*", mResponseText)
    
    '1900/1/1にグリニッジ標準時間と東京／日本の時差分を追加する
    mSecondToNowDate = Int(mSecondToNowDate + (cTimeDifferenceForGMT * cHourSecond))
    
    '----------------------------------------
    ' 経過日数、時刻情報を取得
    '----------------------------------------
    '現在までの経過日数、時刻情報を取得する 
    '※「mDays」日：「mHour」時間：「mMinute」分：「mSecond」秒 1900/1/1から経過しています
    
    '「mDays」日を取得
    Dim mDays                  : mDays                  = Int(mSecondToNowDate / cDaySecond)
    Dim mDaySecondToNowDate    : mDaySecondToNowDate    = mDays * cDaySecond
    
    '「mHour」時間を取得する
    Dim mHour                  : mHour                  = Int((mSecondToNowDate - mDaySecondToNowDate) / cHourSecond)
    Dim mHourSecondToNowDate   : mHourSecondToNowDate   = mHour * cHourSecond
    
    '「mMinute」分を取得する
    Dim mMinute                : mMinute                = Int((mSecondToNowDate - ( mDaySecondToNowDate + mHourSecondToNowDate ) ) / cMinuteSecond)
    Dim mMinuteSecondToNowDate : mMinuteSecondToNowDate = mMinute * cMinuteSecond
    
    '「mSecond」秒を取得する
    Dim mSecond                : mSecond = Int(mSecondToNowDate - ( mDaySecondToNowDate + mHourSecondToNowDate + mMinuteSecondToNowDate ))
    
    '現在日付、現在時刻を取得
    Dim mDate : mDate = DateAdd("d", CCur(mDays), cCalculationMinDay)
    Dim mTime : mTime = mHour & ":" & mMinute & ":" & mSecond

    '----------------------------------------
    ' 日付・時刻設定
    '----------------------------------------
    Dim mObjShell : Set mObjShell = WScript.CreateObject("WScript.Shell") 
    
    '日付設定コマンドを実行
    Dim mDateSetCommand : mDateSetCommand = "cmd /c date " & mDate
    mObjShell.Exec(mDateSetCommand)
    
    '時刻設定コマンドを実行
    Dim mTimeSetCommand : mTimeSetCommand = "cmd /c time " & mTime
    mObjShell.Exec(mTimeSetCommand)

    '「WScript.Shell」オブジェクトの破棄処理
    Set mObjShell = Nothing

    '----------------------------------------
    ' 終了処理
    '----------------------------------------
    '処理の完了メッセージを表示
    MsgBox "「" & mDate & " " & mTime & "」で" & VbCrLf & cMsgBoxProcessSucess, vbOKOnly, cMsgBoxTitle

End Sub


'***********************************************************************
'* 処理名   ： 管理者権限でスクリプトを実行                            *
'* 引数     ： なし                                                    *
'* 処理内容 ： WSHバージョンが5.7以上のときWMI情報からOS情報を取得し、 *
'*             OSのバージョンがVISTA以降の時は管理者権限に昇格する     *
'*             ※Windows Vista 以上の環境では、ユーザーアカウント制御  *
'*               (UAC)のため、VBScriptを「管理者として実行」できません *
'* 参考URL  ：https://www.server-world.info/query?os=Other&p=vbs&f=1   *
'* 戻り値   ： なし                                                    *
'***********************************************************************
Sub RunAdministrator()

    'コマンドライン引数が０件 かつ WSHのバージョンが「5.7」以上の時
    Do While WScript.Arguments.Count = 0 And WScript.Version >= 5.7

        'WMIからOS情報を取得
        Dim WMI : Set WMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
        Dim OS  : Set OS = WMI.ExecQuery("SELECT *FROM Win32_OperatingSystem")
        
        'OS情報分繰り返す
        Dim mValue
        For Each mValue in OS
        
            'OSのバージョンが「6.0」より小さい時は処理を終了
            If Left(mValue.Version, 3) < 6.0 Then Exit Sub
            
        Next

        'このスクリプトを管理者で実行
        Dim mObjApp : Set mObjApp = CreateObject("Shell.Application")
        mObjApp.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ uac", "", "runas"

        WScript.Quit
        
    Loop

End Sub

'***********************************************************************
'* 処理名   ： 正規表現に一致する文字列を取得                          *
'* 引数     ： pPattern 正規表現パターン                               *
'*             pString  対象文字列                                     *
'* 処理内容 ： パターンに一致（正規表現でチェック）した文字列を取得    *
'* 戻り値   ： パターンに一致した文字列                                *
'***********************************************************************
Function GetTextForRegExpMatch(pPattern, pString)

    '----------------------------------
    ' 正規表現オブジェクトを作成       
    '----------------------------------
    Dim mRegExp : Set mRegExp = New RegExp
    mRegExp.Pattern    = pPattern '正規表現のパターンを設定
    mRegExp.IgnoreCase = True     '大文字・小文字を区別しないように設定
    mRegExp.Global     = True     '文字列を全体を検索するように設定
    
    '----------------------------------
    ' 正規表現から一致する文字列を取得 
    '----------------------------------
    '文字列の検索を実行
    Dim mMtchesText : Set mMtchesText = mRegExp.Execute(pString)
    
    Dim mReturnString, mMatchText
    
    'mMtchesText（Matchesコレクション）分繰り返す
    For Each mMatchText in mMtchesText
    
        mReturnString = mReturnString & mMatchText.Value
    
    Next
    
    '----------------------------------
    ' 作成したオブジェクトを破棄
    '----------------------------------
    Set mRegExp     = Nothing
    Set mMtchesText = Nothing
    
    '返り値を設定
    GetTextForRegExpMatch = mReturnString
    
End Function