<#---------------------------------------------------------------------------------
◆セキュリティーグループに参加しているアカウント情報取得スクリプト
･アカウント情報を取得したいセキュリティーグループのdistinguishedNameを指定します。
･実行ファイルと同じディレクトリにCSVファイルでログ出力します。
･CSVファイルに出力する情報は、下記の通りです。
  -sAMAccountName
  -proxyAddresses
  -UserPrincipalName
  -Surname
  -givenName
---------------------------------------------------------------------------------#>

#--------------- 初期設定 ---------------
#SGからアカウントを抽出するDNを指定
$targetDN = ""


#--------------- 処理内容 ---------------
#実行ファイルなどの格納先パスを指定する
$FileDirectory = Split-Path $myInvocation.MyCommand.Path -Parent

#ログファイルのフルパスを作成
$now = Get-Date -format "yyyyMMdd_HHmmss"
$PathLog = $FileDirectory + "\Log_" + $now.ToString() + ".csv"

#対象のSGを取得する、取得できない場合は処理を終了する
try{
    $infoSgMembers = Get-ADGroupMember -Identity $targetDN
}catch{
    #エラーログ出力
    $error[0] | Out-File $PathLog -Append
    exit
}

#対象のSGに所属するアカウントの情報を取得する
if($infoSgMembers){

    #ログヘッダー
    "SAN,PA,UPN,姓,名" | Out-File $PathLog -Append

    foreach($infoSgMember in $infoSgMembers){
        $userInfo = Get-ADUser -Filter {SamAccountName -eq $infoSgMember.SamAccountName} -Properties proxyAddresses | Select sAMAccountName,proxyAddresses,UserPrincipalName,Surname,givenName
        $outLog = $userInfo.sAMAccountName + "," + $userInfo.proxyAddresses + "," + $userInfo.UserPrincipalName + "," + $userInfo.Surname + "," + $userInfo.givenName
        $outLog | Out-File $PathLog -Append    
    }
}else{
    #エラーログ出力
    "ADに存在しないSGです。" | Out-File $PathLog -Append
     exit
}
