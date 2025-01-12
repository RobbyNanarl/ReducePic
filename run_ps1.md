### PowerShellスクリプトの実行について

このツールはPowerShellスクリプトで記述されています。
Windowsの標準的な環境では、PowerShellスクリプトは直接実行できない設定になっており、実行するにはあらかじめ実行ポリシーを変更して権限を与える必要があります。  
このツールに同梱のショートカットには、実行中のみ実行権限を一時的に与えてこのツールを起動できる指定を含めています。  
もし何らかの理由でこのショートカットで起動できない場合、次のいずれかのような手段が考えられます。

1. Windowsショートカットを作り直す  
`ReducePic.ps1`のショートカットを同じ場所に作成  
ショートカットのプロパティを開き次のとおり設定  
リンク先　`powershell.exe -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File ReducePic.ps1`  
作業フォルダー　消す（空欄）　※そのままでも問題はない  
実行時の大きさ　最小化  
（`powershell.exe`は入力後、`C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe` のように自動でその環境でのフルパスになります。）

2. PowerShellコンソールから手動で実行  
`Set-ExecutionPolicy RemoteSigned -Scope Process`　（→ [Y] はい）  
`.\ReducePic.ps1`  
※この指定方法では、PowerShellコンソールを閉じると実行ポリシーは元に戻ります。
