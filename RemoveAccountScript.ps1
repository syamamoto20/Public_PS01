<#---------------------------------------------------------------------------------
◆アカウント削除スクリプト
･指定したOUで、下記条件に合致したアカウントを削除します。
  -EnabledがFalse
  -whenChangedがスクリプト実行日より60日前の日付
･削除するアカウントの関連情報と削除した処理ログをCSVファイルに出力します。
･CSVファイルは、実行ファイルと同じディレクトリにLogフォルダを作成し、その中に出力します。
---------------------------------------------------------------------------------#>

#～～～～～～～～～～ 初期設定 ～～～～～～～～～～
#削除対象のOU
$TargetSearchBase = "OU="

#60日前の日付を設定する
$TargetDaysAgo = ((Get-Date).AddDays(-60)).ToString("yyyy/MM/dd")


#～～～～～～～～～～ 準備処理 ～～～～～～～～～～
#例外処理用
$ErrorActionPreference = "Stop"

#処理件数用
$OkNo = 0
$NgNo = 0

#実行ファイルの格納先を取得する
$FileDirectory = Split-Path $myInvocation.MyCommand.Path -Parent

#ログファイルとバックアップCSV用のフルパスを作成
$now = Get-Date -format "yyyyMMdd_HHmmss"
$PathLog = $FileDirectory + "\Log\Log_" + $now.ToString() + ".csv"
$PathProcessLog = $FileDirectory + "\ProcessLog\PL_" + $now.ToString() + ".csv"

#ログヘッダー
"開始時間,UPN,DistinguishedName,結果,備考" | Out-File $PathProcessLog -Append


#～～～～～～～～～～ メイン処理 ～～～～～～～～～～
Write-Output "処理開始"

Import-Module ActiveDirectory

#削除対象アカウントを取得する
#条件：EnabledがFalse、whenChangedがスクリプト実行日より60日前の日付であること
$TargetUsers = Get-ADUser -Filter {Enabled -eq "False"} -SearchBase $TargetSearchBase -Properties whenChanged | Sort-Object SamAccountName | Where-Object {$_.whenChanged -lt $TargetDaysAgo}

#削除対象アカウントが該当する場合
if($TargetUsers){

    #削除対象アカウントのGet-ADUser情報をCSV出力する
    $TargetUsers | Export-CSV $PathLog -Encoding Default

    foreach($TargetUser in $TargetUsers){
    
        #エラーメッセージ、エラーフラグの初期化
        $MessError = ""

        #対象アカウントを削除する
        try{
            Remove-ADUser -Identity $TargetUser.DistinguishedName -Confirm:$false
        }catch{
            #エクセプション時のエラーメッセージ設定
            $MessError = "[RemoveErr]" + $error[0]
        }

        if([String]::IsNullOrEmpty($MessError)){
            #正常ログ
            $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
            $now.ToString() + ",`"" + $TargetUser.UserPrincipalName + "`",`"" + $TargetUser.DistinguishedName + "`",○,削除完了" | Out-File $PathProcessLog -Append
            $OkNo++
        }else{
            #エラーログ
            $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
            $now.ToString() + ",`"" + $TargetUser.UserPrincipalName + "`",`"" + $TargetUser.DistinguishedName + "`",×," + $MessError | Out-File $PathProcessLog -Append
            $NgNo++
        }
    }
#削除対象アカウントが該当しない場合、ログのみ出力する
}else{
    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
    $now.ToString() + ",-,-,-,[処理なし]削除対象、該当なし" | Out-File $PathProcessLog -Append
}

#ログ出力
$StrError = "失敗 " + $NgNo.ToString() + "件"
$now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
$now.ToString() + ",-,-,-,成功 " + $OkNo.ToString() + "件 / " + $StrError + "の処理が完了" | Out-File $PathProcessLog -Append

Write-Output "処理完了"
exit
