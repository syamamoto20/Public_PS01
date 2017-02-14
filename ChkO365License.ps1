<#---------------------------------------------------------------------------------
◆Office 365ライセンス設定状況確認スクリプト
･Office 365のライセンスで、サービスプランの設定状況を確認します。
･対象はライセンスが有効である全てのアカウントになります。
･O365に接続するアカウント、確認したいサービスプランのライセンスを指定します。
･実行ファイルと同じディレクトリにCSVファイルでログ出力します。
---------------------------------------------------------------------------------#>

<# ---------- 初期設定 ---------- #>
#O365に接続するアカウント情報
$loginUserName = "" 
$password = ""

#設定を確認するサービスプランのライセンス
$O365_BUSINESS_PREMIUM = ":O365_BUSINESS_PREMIUM"
$POWERAPPS_VIRAL = ":POWERAPPS_VIRAL"


<# ---------- 準備処理 ---------- #>
#処理件数用
$OkNo = 0
$NgNo = 0
$ChkNo = 0

#例外処理用
$ErrorActionPreference = "Stop"


<# ---------- Function ---------- #>
#【ConnectMsolService】 テナントに接続する
function ConnectMsolService(){
    $cred = New-Object System.Management.Automation.PSCredential($loginUserName, $(ConvertTo-SecureString $password -asplaintext -force))
    $msolCred = Get-Credential -Credential $cred
    Connect-MsolService -Credential $msolCred -EV err

    if ($err -ne 0){
        return $false
    }
    return $true
}

#【ContainsTargetService】 サービスプランの設定有無を確認する
function ContainsTargetService($license){
    $StrLog = ""
    $ChkNo = 0

    foreach ($target in $targetServiceInfo){
        foreach ($serviceStatus in $license.ServiceStatus){
            foreach ($servicePlan in $serviceStatus.ServicePlan){
                if ($servicePlan.ServiceName -eq $target){

                    if($serviceStatus.ProvisioningStatus -eq "Disabled"){                        
                        #サービスプランが無効化の場合
                        $StrLog = $StrLog + $serviceStatus.ProvisioningStatus + ","
                        $ChkNo++
                    }else{
                        #サービスプランが有効化の場合
                        $StrLog = $StrLog + $serviceStatus.ProvisioningStatus + ","
                    }              
                }
            }
        }
    }
    return $StrLog,$ChkNo
}


<# ---------- メイン処理 ---------- #>
Write-Output "確認開始"

#実行ファイルなどの格納先パスを指定する
$FileDirectory = Split-Path $myInvocation.MyCommand.Path -Parent

#ログファイルのパスを作成する
$now = Get-Date -format "yyyyMMdd_HHmmss"
$PathLog = $FileDirectory + "\Log_" + $now.ToString() + ".csv"

#テナントに接続する
$ret = ConnectMsolService $loginUserName $password

#テナントに接続に失敗した場合、処理を終了する
if(!$ret){
    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
    $now.ToString() + " Office 365 の接続に失敗、処理を中断" | Out-File $PathLog -Append
    exit
}

#ログヘッダーを出力する
"Time,UPN,POWERAPPS,FLOW,TEAMS,結果,備考" | Out-File $PathLog -Append

try{
    #ライセンスが有効である全てのユーザーを取得する
    $users = Get-MsolUser -All | where {$_.isLicensed -eq "True"} | Sort-Object UserPrincipalName
}catch{
    #エクセプション時、処理を終了する
    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
    $now.ToString() + ",-,-,-,-,-," + $error[0] | Out-File $PathLog -Append
    exit
}

if($users){
    foreach ($user in $users){
        foreach ($license in $user.Licenses){           
         
            #O365_BUSINESS_PREMIUM、POWERAPPS_VIRALによって処理を分ける
            switch($license.AccountSkuId){

                #ライセンスがO365_BUSINESS_PREMIUMの場合
                $O365_BUSINESS_PREMIUM{

                    #確認するサービスプラン名を設定する
                    [string[]]$TargetServiceInfo = "POWERAPPS_O365_P1","FLOW_O365_P1","TEAMS1"

                    #対象のサービスプラン名が無効化か確認する
                    [string[]]$StrResult = ContainsTargetService($license)

                    #ログ出力処理
                    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
                    if($StrResult[1] -eq 3){
                        #正常ログ
                        $now.ToString() + ",`"" + $user.UserPrincipalName + "`"," + $StrResult[0] + "○,処理完了" | Out-File $PathLog -Append
                        $OkNo++
                    }else{
                        #エラーログ
                        $now.ToString() + ",`"" + $user.UserPrincipalName + "`"," + $StrResult[0] + "×,処理未完了" | Out-File $PathLog -Append
                        $NgNo++
                    }
                }

                #ライセンスがPOWERAPPS_VIRALの場合
                $POWERAPPS_VIRAL{

                    #確認するサービスプラン名を設定する                                        
                    [string[]]$TargetServiceInfo = "POWERAPPS_P2_VIRAL","FLOW_P2_VIRAL"   

                    #対象のサービスプラン名が無効化か確認する
                    [string[]]$StrResult = ContainsTargetService($license)
                    
                    #ログ出力処理
                    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
                    if($StrResult[1] -eq 2){
                        #正常ログ
                        $now.ToString() + ",`"" + $user.UserPrincipalName + "`"," + $StrResult[0] + "-,○,処理完了" | Out-File $PathLog -Append
                        $OkNo++
                    }else{
                        #エラーログ
                        $now.ToString() + ",`"" + $user.UserPrincipalName + "`"," + $StrResult[0] + "-,×,処理未完了" | Out-File $PathLog -Append
                        $NgNo++
                    }  	
                }
            }
        }
    }
}else{
    #エラーログ
    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
    $now.ToString() + ",-,-,-,-,-,アカウントが取得できません" | Out-File $PathLog -Append
    exit
}

#ログ出力
$now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
$now.ToString() + ",-,-,-,-,-,完了 " + $OkNo.ToString() + "件 / 未完了 " + $NgNo.ToString()  + "件" | Out-File $PathLog -Append
Write-Output "確認完了"
exit
