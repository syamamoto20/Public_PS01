<#---------------------------------------------------------------------------------
���A�J�E���g�ꎞ�������X�N���v�g
�CSV�t�@�C������ǂݍ���UserPrincipalName��Ώۂɉ��L�������s���܂��B
  -�A�J�E���g�̖�����
  -�A�J�E���g�̃A�h���X���\�����\��
  -Domain Users�ȊO�̃Z�L�����e�B�[�O���[�v���폜  
  -�A�J�E���g���w��OU�Ɉړ�
����s�t�@�C���Ɠ����f�B���N�g����CSV�t�@�C���Ń��O�o�͂��܂��B
---------------------------------------------------------------------------------#>

#�`�`�`�`�`�`�`�`�`�` �����ݒ� �`�`�`�`�`�`�`�`�`�`
#���̓��[�U���X�gCSV���w��
$OuFile = "TemporaryDisableList.csv"

#�ޔ���OU�����w��
$PathOu = ""


#�`�`�`�`�`�`�`�`�`�` �������� �`�`�`�`�`�`�`�`�`�`
#��O�����p
$ErrorActionPreference = "Stop"

#���������p
$OkNo = 0
$NgNo = 0
$NgEtcNo = 0

#���s�t�@�C���Ȃǂ̊i�[��p�X���w��
$FileDirectory = Split-Path $myInvocation.MyCommand.Path -Parent

#CSV�t�@�C���̃t���p�X���쐬
$PathCsv = $FileDirectory + "\" + $OuFile

#���O�t�@�C���̃t���p�X���쐬
$now = Get-Date -format "yyyyMMdd_HHmmss"
$PathLog = $FileDirectory + "\Log_" + $now.ToString() + ".csv"

#CSV�t�@�C�������݂��邩�m�F�A�Ȃ��ꍇ�͏������I������
if (Test-Path $PathCsv){
    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
    $now.ToString() + ",-,-,-,CSV�t�@�C������A�����J�n" | Out-File $PathLog -Append

    #�����J�n�̃��b�Z�[�W
    Write-Output "CSV�t�@�C������A�����J�n"
}else{
    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
    $now.ToString() + ",-,-,-,CSV�t�@�C���Ȃ��A�����I��" | Out-File $PathLog -Append
    exit
}

#���O�w�b�_�[
"�J�n����,UPN,DistinguishedName,����,���l" | Out-File $PathLog -Append

Import-Module ActiveDirectory

#�`�`�`�`�`�`�`�`�`�` ���C������ �`�`�`�`�`�`�`�`�`�`
Import-Csv $PathCsv | foreach {
    #UserPrincipalName�őΏۃ��[�U�[�̏����擾����
    $Upn = $_.UserPrincipalName
    $UserInfo = Get-ADUser -Filter {UserPrincipalName -eq $Upn}

    if($UserInfo){
        #�G���[���b�Z�[�W�A�G���[�t���O�̏�����
        $MessError = ""

        try{
            #�Ώۂ̃A�J�E���g�̖�����
            Get-ADUser -Filter {UserPrincipalName -eq $UserInfo.UserPrincipalName} | Set-ADUser -Enabled $False

            #�Ώۂ̃A�J�E���g�̃A�h���X�����\���ɂ���imsExchHideFromAddressLists��"True"�ɂ���j
            Get-ADUser -Filter {UserPrincipalName -eq $UserInfo.UserPrincipalName} | Set-ADUser -replace @{msExchHideFromAddressLists=$True}

            #�Ώۂ̃A�J�E���g��SG�擾           
            $UserSgs = Get-ADPrincipalGroupMembership -Identity $UserInfo.DistinguishedName
            
            #Domain Users�ȊO��SG���폜  
            if($UserSgs){
                foreach($UserSg in $UserSgs){
                    if($UserSg.name -ne "Domain Users"){
                        Remove-ADGroupMember -Identity $UserSg.name -Members $UserInfo.DistinguishedName -Confirm:$False
                    }            
                } 
            }

            #�Ώۂ̃A�J�E���g���w��OU�Ɉړ�����
            Move-ADObject -Identity $UserInfo.DistinguishedName -TargetPath $PathOu
        }catch{
            #�G�N�Z�v�V�������̃G���[���b�Z�[�W�ݒ�
            $MessError = $error[0]
            $NgNo++
        }

        #���O�o��
        if([String]::IsNullOrEmpty($MessError)){
            #���탍�O
            $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
            $now.ToString() + "," + $Upn.ToString() + ",""" + $UserInfo.DistinguishedName + """,��,�폜�Ή��ς�" | Out-File $PathLog -Append

            #���폈���������J�E���g����
            $OkNo++
        }else{
            #�G���[���O
            $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
            $now.ToString() + "," + $Upn.ToString() + ",-,�~," + $MessError | Out-File $PathLog -Append
        }
    #AD��񂪑��݂��Ȃ��ꍇ�A�G���[����
    }else{
        #�G���[���O�o��
        $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
        $now.ToString() + "," + $Upn.ToString() + ",-,�~,AD�ɑ��݂��Ȃ����[�U�[" | Out-File $PathLog -Append

        #�G���[�����������J�E���g����
        $NgEtcNo++
    }
}

#���O�o��
$now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
$StrError = "�G���[ : " + $NgNo.ToString() + "�� �Y���Ȃ� : " + $NgEtcNo.ToString() + "��"
$now.ToString() + ",-,-,-,���� " + $OkNo.ToString() + "�� / " + $StrError + "�̏�������" | Out-File $PathLog -Append

#���������̃��b�Z�[�W
Write-Output "�������������܂����B"
exit
