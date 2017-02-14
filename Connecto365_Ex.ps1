<#---------------------------------------------------------------------------------
◆Exchange Online (Office 365)接続スクリプト
･指定した実行アカウントで自動的にExchange Online (Office 365)に接続します。
･実行アカウントのパスワードは、暗号化してcre.passに保存します。
･cre.passは実行ファイルと同じディレクトリに格納します。
---------------------------------------------------------------------------------#>

#～～～～～～～～～～ 初期設定 ～～～～～～～～～～
#実行アカウント
$Account = ""

#実行ファイルなどの格納先パスを取得
$FileDirectory = Split-Path $myInvocation.MyCommand.Path -Parent

#パスワードファイルのフルパスを作成
$PathPassword = $FileDirectory + "\cre.pass"

#～～～～～～～～～～ メイン処理 ～～～～～～～～～～
#必要なファイルが存在するか確認
if (Test-Path $PathPassword){
    Write-Output "cre.passファイルを確認"

    #O365へ接続
    $Password = Get-Content $PathPassword | ConvertTo-SecureString
    $LiveCred = New-Object System.Management.Automation.PSCredential $Account,$Password
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber
    Connect-MsolService -Credential $LiveCred
}else{
    Write-Output "cre.passファイルを確認できず"
}
