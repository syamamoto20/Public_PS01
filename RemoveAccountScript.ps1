<#---------------------------------------------------------------------------------
���A�J�E���g�폜�X�N���v�g
��w�肵��OU�ŁA���L�����ɍ��v�����A�J�E���g���폜���܂��B
  -Enabled��False
  -whenChanged���X�N���v�g���s�����60���O�̓��t
��폜����A�J�E���g�̊֘A���ƍ폜�����������O��CSV�t�@�C���ɏo�͂��܂��B
�CSV�t�@�C���́A���s�t�@�C���Ɠ����f�B���N�g����Log�t�H���_���쐬���A���̒��ɏo�͂��܂��B
---------------------------------------------------------------------------------#>

#�`�`�`�`�`�`�`�`�`�` �����ݒ� �`�`�`�`�`�`�`�`�`�`
#�폜�Ώۂ�OU
$TargetSearchBase = "OU="

#60���O�̓��t��ݒ肷��
$TargetDaysAgo = ((Get-Date).AddDays(-60)).ToString("yyyy/MM/dd")


#�`�`�`�`�`�`�`�`�`�` �������� �`�`�`�`�`�`�`�`�`�`
#��O�����p
$ErrorActionPreference = "Stop"

#���������p
$OkNo = 0
$NgNo = 0

#���s�t�@�C���̊i�[����擾����
$FileDirectory = Split-Path $myInvocation.MyCommand.Path -Parent

#���O�t�@�C���ƃo�b�N�A�b�vCSV�p�̃t���p�X���쐬
$now = Get-Date -format "yyyyMMdd_HHmmss"
$PathLog = $FileDirectory + "\Log\Log_" + $now.ToString() + ".csv"
$PathProcessLog = $FileDirectory + "\ProcessLog\PL_" + $now.ToString() + ".csv"

#���O�w�b�_�[
"�J�n����,UPN,DistinguishedName,����,���l" | Out-File $PathProcessLog -Append


#�`�`�`�`�`�`�`�`�`�` ���C������ �`�`�`�`�`�`�`�`�`�`
Write-Output "�����J�n"

Import-Module ActiveDirectory

#�폜�ΏۃA�J�E���g���擾����
#�����FEnabled��False�AwhenChanged���X�N���v�g���s�����60���O�̓��t�ł��邱��
$TargetUsers = Get-ADUser -Filter {Enabled -eq "False"} -SearchBase $TargetSearchBase -Properties whenChanged | Sort-Object SamAccountName | Where-Object {$_.whenChanged -lt $TargetDaysAgo}

#�폜�ΏۃA�J�E���g���Y������ꍇ
if($TargetUsers){

    #�폜�ΏۃA�J�E���g��Get-ADUser����CSV�o�͂���
    $TargetUsers | Export-CSV $PathLog -Encoding Default

    foreach($TargetUser in $TargetUsers){
    
        #�G���[���b�Z�[�W�A�G���[�t���O�̏�����
        $MessError = ""

        #�ΏۃA�J�E���g���폜����
        try{
            Remove-ADUser -Identity $TargetUser.DistinguishedName -Confirm:$false
        }catch{
            #�G�N�Z�v�V�������̃G���[���b�Z�[�W�ݒ�
            $MessError = "[RemoveErr]" + $error[0]
        }

        if([String]::IsNullOrEmpty($MessError)){
            #���탍�O
            $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
            $now.ToString() + ",`"" + $TargetUser.UserPrincipalName + "`",`"" + $TargetUser.DistinguishedName + "`",��,�폜����" | Out-File $PathProcessLog -Append
            $OkNo++
        }else{
            #�G���[���O
            $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
            $now.ToString() + ",`"" + $TargetUser.UserPrincipalName + "`",`"" + $TargetUser.DistinguishedName + "`",�~," + $MessError | Out-File $PathProcessLog -Append
            $NgNo++
        }
    }
#�폜�ΏۃA�J�E���g���Y�����Ȃ��ꍇ�A���O�̂ݏo�͂���
}else{
    $now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
    $now.ToString() + ",-,-,-,[�����Ȃ�]�폜�ΏہA�Y���Ȃ�" | Out-File $PathProcessLog -Append
}

#���O�o��
$StrError = "���s " + $NgNo.ToString() + "��"
$now = Get-Date -format "yyyy/MM/dd HH:mm:ss"
$now.ToString() + ",-,-,-,���� " + $OkNo.ToString() + "�� / " + $StrError + "�̏���������" | Out-File $PathProcessLog -Append

Write-Output "��������"
exit
