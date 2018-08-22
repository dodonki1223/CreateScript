# Hyper-Vの機能を有効・無効にするバッチ

## Hyper-Vの現在の状態を表示しユーザーに対話でHyper-Vの機能の有効・無効にするかを選択させる。設定後、再起動をすることでHyper-Vの機能が有効・無効になります

### ファイルの説明

- 01_ToEnableHyper-V.bat
    - 初めに実行するbatファイル
    - Hyper-Vの状態の確認やHyper-Vの機能を有効・無効にするには管理者で実行する必要がある
    - このBatファイルで「Hyper-Vの状態確認」、「Hyper-Vの有効」、「Hyper-Vの無効」を管理者で実行する 
- 02_ToEnableHyper-V-Command.bat
    - Hyper-Vの機能を有効化するためのファイル
- 03_ToDisableHyper-V-Command.bat
    - Hyper-Vの機能を無効化するためのファイル
- 04_CheckHyper-V-Status.bat
    - Hyper-Vの現在の状態を確認するためのファイル
    - 有効なのか無効なのか確認できる

