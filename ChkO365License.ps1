<#---------------------------------------------------------------------------------
��Office 365���C�Z���X�ݒ�󋵊m�F�X�N���v�g
�Office 365�̃��C�Z���X�ŁA�T�[�r�X�v�����̐ݒ�󋵂��m�F���܂��B
��Ώۂ̓��C�Z���X���L���ł���S�ẴA�J�E���g�ɂȂ�܂��B
�O365�ɐڑ�����A�J�E���g�A�m�F�������T�[�r�X�v�����̃��C�Z���X���w�肵�܂��B
����s�t�@�C���Ɠ����f�B���N�g����CSV�t�@�C���Ń��O�o�͂��܂��B
---------------------------------------------------------------------------------#>

<# ---------- �����ݒ� ---------- #>
#O365�ɐڑ�����A�J�E���g���
$loginUserName = "" 
$password = ""

#�ݒ���m�F����T�[�r�X�v�����̃��C�Z���X
$O365_BUSINESS_PREMIUM = ":O365_BUSINESS_PREMIUM"
$POWERAPPS_VIRAL = ":POWERAPPS_VIRAL"


<# ---------- �������� ---------- #>
#���������p
$OkNo = 0
$NgNo = 0
$ChkNo = 0

#��O�����p
$ErrorActionPreference = "Stop"


<# ---------- Function ---------- #>
#�yConnectMsolService�z �e�i���g�ɐڑ�����
function ConnectMsolService(){
    $cred = New-Object System.Management.Automation.PSCredential($loginUserName, $(ConvertTo-SecureString $password -asplaintext -force))
    $msolCred = Get-Credential -Credential $cred
    Connect-MsolService -Credential $msolCred -EV err

    if ($err -ne 0){
        return $false
    }
    return $true
}

#�yContainsTargetService�z �T�[�r�X�v�����̐ݒ�L�����m�F����
function ContainsTargetService($license){
    $StrLog = ""
    $ChkNo = 0

    foreach ($target in $targetServiceInfo){
        foreach ($serviceStatus in $license.ServiceStatus){
            foreach ($servicePlan in $serviceStatus.ServicePlan){
                if ($servicePlan.ServiceName -eq $target){

                    if($serviceStatus.ProvisioningStatus -eq "Disabled"){                        
                        #�T�[�r�X�v�������������̏ꍇ
                        $StrLog = $StrLog + $serviceStatus.ProvisioningStatus + ","
                        $ChkNo++
                    }else{
                        #�T�[�r�X�v�������L�����̏ꍇ
                        $StrLog = $StrLog + $serviceStatus.ProvisioningStatus + ","
                    }              
                }
            }
        }
    }
    return $StrLog,$ChkNo
}


<# ---------- ���C������ ---------- #>
Write-Output "�m�F�J�n"

#���s�t�@�C���Ȃǂ̊i�[��p�X���w�肷��
$FileDirectory = Split-Path $myInvocation.MyCommand.Path -Parent

#���O�t�@�C���̃p�X���쐬����
$now = Get-Date -format "yyyyMMdd_HHmmss"
$PathLog = $FileDirectory + "\Log_" + $now.ToString() + ".csv"

#�e�i���g�ɐڑ�����
$ret = ConnectMsolService $loginUserName $password

#�e�i���g�ɐڑ��Ɏ��s�����ꍇ�A�������I������
if(!$ret){
    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
    $now.ToString() + " Office 365 �̐ڑ��Ɏ��s�A�����𒆒f" | Out-File $PathLog -Append
    exit
}

#���O�w�b�_�[���o�͂���
"Time,UPN,POWERAPPS,FLOW,TEAMS,����,���l" | Out-File $PathLog -Append

try{
    #���C�Z���X���L���ł���S�Ẵ��[�U�[���擾����
    $users = Get-MsolUser -All | where {$_.isLicensed -eq "True"} | Sort-Object UserPrincipalName
}catch{
    #�G�N�Z�v�V�������A�������I������
    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
    $now.ToString() + ",-,-,-,-,-," + $error[0] | Out-File $PathLog -Append
    exit
}

if($users){
    foreach ($user in $users){
        foreach ($license in $user.Licenses){           
         
            #O365_BUSINESS_PREMIUM�APOWERAPPS_VIRAL�ɂ���ď����𕪂���
            switch($license.AccountSkuId){

                #���C�Z���X��O365_BUSINESS_PREMIUM�̏ꍇ
                $O365_BUSINESS_PREMIUM{

                    #�m�F����T�[�r�X�v��������ݒ肷��
                    [string[]]$TargetServiceInfo = "POWERAPPS_O365_P1","FLOW_O365_P1","TEAMS1"

                    #�Ώۂ̃T�[�r�X�v�����������������m�F����
                    [string[]]$StrResult = ContainsTargetService($license)

                    #���O�o�͏���
                    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
                    if($StrResult[1] -eq 3){
                        #���탍�O
                        $now.ToString() + ",`"" + $user.UserPrincipalName + "`"," + $StrResult[0] + "��,��������" | Out-File $PathLog -Append
                        $OkNo++
                    }else{
                        #�G���[���O
                        $now.ToString() + ",`"" + $user.UserPrincipalName + "`"," + $StrResult[0] + "�~,����������" | Out-File $PathLog -Append
                        $NgNo++
                    }
                }

                #���C�Z���X��POWERAPPS_VIRAL�̏ꍇ
                $POWERAPPS_VIRAL{

                    #�m�F����T�[�r�X�v��������ݒ肷��                                        
                    [string[]]$TargetServiceInfo = "POWERAPPS_P2_VIRAL","FLOW_P2_VIRAL"   

                    #�Ώۂ̃T�[�r�X�v�����������������m�F����
                    [string[]]$StrResult = ContainsTargetService($license)
                    
                    #���O�o�͏���
                    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
                    if($StrResult[1] -eq 2){
                        #���탍�O
                        $now.ToString() + ",`"" + $user.UserPrincipalName + "`"," + $StrResult[0] + "-,��,��������" | Out-File $PathLog -Append
                        $OkNo++
                    }else{
                        #�G���[���O
                        $now.ToString() + ",`"" + $user.UserPrincipalName + "`"," + $StrResult[0] + "-,�~,����������" | Out-File $PathLog -Append
                        $NgNo++
                    }  	
                }
            }
        }
    }
}else{
    #�G���[���O
    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
    $now.ToString() + ",-,-,-,-,-,�A�J�E���g���擾�ł��܂���" | Out-File $PathLog -Append
    exit
}

#���O�o��
$now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
$now.ToString() + ",-,-,-,-,-,���� " + $OkNo.ToString() + "�� / ������ " + $NgNo.ToString()  + "��" | Out-File $PathLog -Append
Write-Output "�m�F����"
exit
