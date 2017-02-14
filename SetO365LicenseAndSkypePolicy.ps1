<#---------------------------------------------------------------------------------
◆O365 ライセンス付与とSkype Onlineのポリシー設定
･指定のCSVファイルから優先メールアドレスを読み込み、O365 ライセンス付与と
 Skype Onlineのポリシーを設定します。
･付与するO365 ライセンスとSkype Onlineのポリシーを指定します。
---------------------------------------------------------------------------------#>

#～～～～～～～～～～ 初期パラメーター ～～～～～～～～～～
#対象のO365ライセンスとSkypeのポリシー
$addlicenses = ""
$PolicyName = ""

#CSVのファイル名
$CsvFile = ""

#～～～～～～～～～～ メイン処理 ～～～～～～～～～～
$LiveCred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber
Connect-MsolService -Credential $LiveCred
Start-Sleep -Seconds 3

Write-Output "◆ライセンスの付与、開始"

<#---------------【 O365 ライセンスの付与 】---------------
◆優先メールアドレスからUPNを取得し、O365 ライセンスを付与する
---------------------------------------------------------#>
Import-Csv $PathCsv | foreach {

    $PrimarySmtpAddress = ""   
    $UPN = ""        
    $chkUPN = ""
    $MessError02 = ""
     
    $PrimarySmtpAddress = $_.PrimarySmtpAddress
    $PrimarySmtpAddress = $PrimarySmtpAddress.Trim()

    try{

        #優先メールアドレスからUPNを取得する
        $UPN = (Get-Mailbox -Identity $PrimarySmtpAddress).UserPrincipalName

        #UPNが存在するか確認する
        $chkUPN = Get-Msoluser -UserPrincipalName $UPN
        if($chkUPN){

            #ライセンスを付与する
            Set-Msoluser -UserPrincipalName $UPN -UsageLocation JP
            Set-MsolUserLicense -UserPrincipalName $UPN -addlicenses $addlicenses

            #Skypeの設定用にUPNを格納する
            [Array]$ArrayUPNs += $UPN
                
            Start-Sleep -Seconds 3
        }
    }catch{

        #エクセプション時のエラーメッセージ設定
        $MessError02 = $error[0]
        $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"       
        $now.ToString() + ",-,"""+ $PrimarySmtpAddress +""",-,-,×," + $MessError02 | Out-File $PathLog -Append        
        $NgNo++
    }        

    #正常ログ出力
    if([String]::IsNullOrEmpty($MessError02)){

        $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
        $now.ToString() + ",-,"""+ $PrimarySmtpAddress +""",-,-,○,ライセンス付与、完了" | Out-File $PathLog -Append        
        $OkNo02++
    }
}

<#---------------【 Skype Onlineの設定 】---------------
◆Skype Onlineのポリシーを設定する
---------------------------------------------------------#>
# Skype Online接続
Import-Module LyncOnlineConnector
$Lyncsession = New-CsOnlineSession -Credential $LiveCred
Import-PSSession $Lyncsession -AllowClobber

Write-Output "◆ライセンスの付与、完了"
Start-Sleep -Seconds 60
Write-Output "◆SkypeのPolicy設定、開始"

foreach($ArrayUPN in $ArrayUPNs){

    $MessError03 = ""

    try{
        Grant-CsconferencingPolicy -Identity $ArrayUPN -PolicyName $PolicyName
        Start-Sleep -Seconds 3
    }catch{

        #エクセプション時のエラーメッセージ設定
        $MessError03 = $error[0]
        $now.ToString() + ",-,""" + $ArrayUPN + """,-,-,×," + $MessError03 | Out-File $PathLog -Append
        $NgNo++
    }

    #正常ログ出力
    if([String]::IsNullOrEmpty($MessError03)){

        $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
        $now.ToString() + ",-,""" + $ArrayUPN + """,-,-,○,Skypeポリシー設定、完了" | Out-File $PathLog -Append
        $OkNo03++
    }
}
Write-Output "◆SkypeのPolicy設定、完了"
exit
