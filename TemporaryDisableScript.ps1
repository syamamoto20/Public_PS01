<#---------------------------------------------------------------------------------
◆アカウント一時無効化スクリプト
･CSVファイルから読み込んだUserPrincipalNameを対象に下記処理を行います。
  -アカウントの無効化
  -アカウントのアドレス帳表示を非表示
  -Domain Users以外のセキュリティーグループを削除  
  -アカウントを指定OUに移動
･実行ファイルと同じディレクトリにCSVファイルでログ出力します。
---------------------------------------------------------------------------------#>

#～～～～～～～～～～ 初期設定 ～～～～～～～～～～
#入力ユーザリストCSVを指定
$OuFile = "TemporaryDisableList.csv"

#退避先のOU名を指定
$PathOu = ""


#～～～～～～～～～～ 準備処理 ～～～～～～～～～～
#例外処理用
$ErrorActionPreference = "Stop"

#処理件数用
$OkNo = 0
$NgNo = 0
$NgEtcNo = 0

#実行ファイルなどの格納先パスを指定
$FileDirectory = Split-Path $myInvocation.MyCommand.Path -Parent

#CSVファイルのフルパスを作成
$PathCsv = $FileDirectory + "\" + $OuFile

#ログファイルのフルパスを作成
$now = Get-Date -format "yyyyMMdd_HHmmss"
$PathLog = $FileDirectory + "\Log_" + $now.ToString() + ".csv"

#CSVファイルが存在するか確認、ない場合は処理を終了する
if (Test-Path $PathCsv){
    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
    $now.ToString() + ",-,-,-,CSVファイルあり、処理開始" | Out-File $PathLog -Append

    #処理開始のメッセージ
    Write-Output "CSVファイルあり、処理開始"
}else{
    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
    $now.ToString() + ",-,-,-,CSVファイルなし、処理終了" | Out-File $PathLog -Append
    exit
}

#ログヘッダー
"開始時間,UPN,DistinguishedName,結果,備考" | Out-File $PathLog -Append

Import-Module ActiveDirectory

#～～～～～～～～～～ メイン処理 ～～～～～～～～～～
Import-Csv $PathCsv | foreach {
    #UserPrincipalNameで対象ユーザーの情報を取得する
    $Upn = $_.UserPrincipalName
    $UserInfo = Get-ADUser -Filter {UserPrincipalName -eq $Upn}

    if($UserInfo){
        #エラーメッセージ、エラーフラグの初期化
        $MessError = ""

        try{
            #対象のアカウントの無効化
            Get-ADUser -Filter {UserPrincipalName -eq $UserInfo.UserPrincipalName} | Set-ADUser -Enabled $False

            #対象のアカウントのアドレス帳を非表示にする（msExchHideFromAddressListsを"True"にする）
            Get-ADUser -Filter {UserPrincipalName -eq $UserInfo.UserPrincipalName} | Set-ADUser -replace @{msExchHideFromAddressLists=$True}

            #対象のアカウントのSG取得           
            $UserSgs = Get-ADPrincipalGroupMembership -Identity $UserInfo.DistinguishedName
            
            #Domain Users以外のSGを削除  
            if($UserSgs){
                foreach($UserSg in $UserSgs){
                    if($UserSg.name -ne "Domain Users"){
                        Remove-ADGroupMember -Identity $UserSg.name -Members $UserInfo.DistinguishedName -Confirm:$False
                    }            
                } 
            }

            #対象のアカウントを指定OUに移動する
            Move-ADObject -Identity $UserInfo.DistinguishedName -TargetPath $PathOu
        }catch{
            #エクセプション時のエラーメッセージ設定
            $MessError = $error[0]
            $NgNo++
        }

        #ログ出力
        if([String]::IsNullOrEmpty($MessError)){
            #正常ログ
            $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
            $now.ToString() + "," + $Upn.ToString() + ",""" + $UserInfo.DistinguishedName + """,○,削除対応済み" | Out-File $PathLog -Append

            #正常処理件数をカウントする
            $OkNo++
        }else{
            #エラーログ
            $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
            $now.ToString() + "," + $Upn.ToString() + ",-,×," + $MessError | Out-File $PathLog -Append
        }
    #AD情報が存在しない場合、エラー処理
    }else{
        #エラーログ出力
        $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
        $now.ToString() + "," + $Upn.ToString() + ",-,×,ADに存在しないユーザー" | Out-File $PathLog -Append

        #エラー処理件数をカウントする
        $NgEtcNo++
    }
}

#ログ出力
$now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
$StrError = "エラー : " + $NgNo.ToString() + "件 該当なし : " + $NgEtcNo.ToString() + "件"
$now.ToString() + ",-,-,-,成功 " + $OkNo.ToString() + "件 / " + $StrError + "の処理完了" | Out-File $PathLog -Append

#処理完了のメッセージ
Write-Output "処理が完了しました。"
exit
