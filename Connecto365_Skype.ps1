<#---------------------------------------------------------------------------------
��Skype Online (Office 365)�ڑ��X�N���v�g
��w�肵�����s�A�J�E���g�Ŏ����I��Skype Online (Office 365)�ɐڑ����܂��B
����s�A�J�E���g�̃p�X���[�h�́A�Í�������cre.pass�ɕۑ����܂��B
�cre.pass�͎��s�t�@�C���Ɠ����f�B���N�g���Ɋi�[���܂��B
---------------------------------------------------------------------------------#>

#�`�`�`�`�`�`�`�`�`�` �����ݒ� �`�`�`�`�`�`�`�`�`�`
#���s�A�J�E���g
$Account = ""

#���s�t�@�C���Ȃǂ̊i�[��p�X���擾
$FileDirectory = Split-Path $myInvocation.MyCommand.Path -Parent

#�p�X���[�h�t�@�C���̃t���p�X���쐬
$PathPassword = $FileDirectory + "\cre.pass"

#�`�`�`�`�`�`�`�`�`�` ���C������ �`�`�`�`�`�`�`�`�`�`
#�K�v�ȃt�@�C�������݂��邩�m�F
if (Test-Path $PathPassword){
    Write-Output "cre.pass�t�@�C�����m�F"

    #O365�֐ڑ�
    $Password = Get-Content $PathPassword | ConvertTo-SecureString
    $LiveCred = New-Object System.Management.Automation.PSCredential $Account,$Password
    Import-Module LyncOnlineConnector
    $Lyncsession = New-CsOnlineSession -Credential $LiveCred
    Import-PSSession $Lyncsession
}else{
    Write-Output "cre.pass�t�@�C�����m�F�ł���"
}
