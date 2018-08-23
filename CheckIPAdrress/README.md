# IPアドレスチェックスクリプト

## ipconfigコマンドで表示されるIPv4アドレスのIPアドレスをメッセージボックスで表示する。

### 下記コマンドを使用してIPアドレスを取得している

- @for /F "delims=: tokens=2" %a in ('ipconfig ^| findstr "IP"') do @echo %a
