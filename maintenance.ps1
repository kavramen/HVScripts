function findowner{
    #script allows to associate users with VMs. Search is based on events.
    $vmmsoper = Get-WinEvent -LogName "Microsoft-Windows-Hyper-V-VMMS-Operational"  | where {$_.id -eq 27311}
    class DiskRecord
    {
        [string]$Disk
        [string]$User
        [string]$UserName
        [string]$VmId
        [string]$VMname
        DiskRecord($Disk, $User, $UserName, $VmId, $VMname) {
        $this.Disk = $Disk
        $this.User = $User
        $this.UserName = $UserName
        $this.VmId = $VmId
        $this.VMname = $VMname
        }
    }
    $vmsonhost = get-vm
    foreach ($event in $vmmsoper)
    {
        $msg = $event.message.substring(33)
        $message = $msg.substring(0,$msg.length-2)
        foreach ($vm in $vmsonhost)
        {
            if ($vm.HardDrives[0].Path -eq $message)
            {
                $cvmid = $vm.vmid
                $cvmname = $vm.name
            }
            else
            {
                $cvmid=''
                $cvmname=''
            }
        }
        $usrname = Get-ADUser -Server spbsupdc02.support2.veeam.local -Identity $event.userid.value
        $obj = New-Object DiskRecord($message,$event.userid.value, $usrname.name, $cvmid, $cvmname)
        $obj | Export-Csv -Path "C:\scripting\vmlist.csv" -Append -NoTypeInformation
    }
    $file = Import-Csv -Path "C:\scripting\vmlist.csv" | sort disk,user -Unique
    $file | Export-Csv -Path "C:\scripting\vmlist.csv" -NoTypeInformation
}    
function findoldvms{
    $vms = Get-VM | where {$_.state -eq 'off'}
    $oldvms = @()
    $vmstodelete = @()
    $file = Import-Csv -Path "C:\scripting\vmlist.csv"
    foreach ($vm in $vms) {
        #check if we have drive attached to find lastrun time
        if (!$vm.HardDrives[0].Path){
            $config = dir $vm.configurationlocation -filter "$($vm.id).xml" -recurse
            $lastwritetime = $config.LastWriteTime
        }
        else{
            $diskFile = Get-Item -Path $vm.HardDrives[0].Path
            $lastwritetime = $diskFile.LastWriteTime
        }
        #deleting vms older than 21 day
        if ($lastwritetime -lt ((get-date).AddDays(-21))){
            $vmstodelete += $vm.name
            Get-VMHardDiskDrive -VM $vm | Foreach { Remove-item -path $_.Path -Recurse -Force -Confirm:$False}
            Remove-VM $vm -Force -Confirm:$False
        }
        elseif ($lastwritetime -lt ((get-date).AddDays(+1))){
            foreach ($string in $file){
                if ($string.vmid -eq $vm.VMId){
                    $string | Export-Csv -Path "C:\scripting\notify.csv" -NoTypeInformation -Append
                }
            }
        }
    }
    $vms = Get-VM | where {$_.state -ne 'off'}
    foreach ($vm in $vms) {
        #finding VMs which are running more than 30 days and deleting them
        if ($vm.uptime.days -gt 30){
            Get-VMHardDiskDrive -VM $vm | Foreach { Remove-item -path $_.Path -Recurse -Force -Confirm:$False}
            Remove-VM $vm -Force -Confirm:$False
        }
    }
}
function FindUserFromSupport2inAMUST{
param(
[parameter(Mandatory=$true,ValueFromPipeline=$true)]
[Microsoft.ActiveDirectory.Management.ADuser]$support2User,
[parameter(Mandatory=$true)]
[PScredential]$AMUSTcreds
)
$samAcc = $support2User.SamAccountName
$UserFullName = $support2User.Name

if($result = Get-ADUser -Filter {Name -eq $support2User.Name} -Properties emailaddress -Credential $AMUSTcreds -Server "amust.local" ){
       return $result}

if($result = Get-ADUser -Filter {SamAccountName -eq $support2User.SamAccountName} -Properties emailaddress -Credential $AMUSTcreds -Server "amust.local" ){
       return $result}



###-SearchBase "OU=Support,OU=Amust,OU=Employees,OU=Accounts,DC=amust,DC=local"
}
function SendEmailFromOutlook{
param(
[parameter(Mandatory = $false)]
$HTMLBodyFragment,
[parameter(Mandatory = $true)]
[string]$recipient,
[parameter(Mandatory = $false)]
[string]$customTextAtTheBegining,
[parameter(Mandatory = $false)]
[string]$subject
)

$htmlBody = @"

<html>

<body>

<div>
<p>
$customTextAtTheBegining
</p>
$HTMLBodyFragment
</div>

<p>
==============================================
</p>

<p>
POSH sup lab scripts
</p>

</body>

</html>
"@

$o = New-Object -com Outlook.Application

$mail = $o.CreateItem(0)

$mail.subject = $subject
$mail.HTMLBody = $htmlBody
$mail.To = $recipient
$mail.Send()
}
function sendnotification{
    $credentials = Import-Clixml "C:\scripting\amustcred.xml"
    $notify = Import-Csv -Path "C:\scripting\notify.csv"
    foreach ($string in $notify){
        $user = get-aduser -server spbsupdc02.support2.veeam.local -Identity $string.user
        $amustuser = FindUserFromSupport2inAMUST -support2User $user -AMUSTcreds $credentials
        $vmname = $string.vmname
        $userstonotify = $amustuser.EmailAddress
        sendnotification -CustomTextAtTheBegining "Please note that your vm $vmname will be deleted soon. To avoid this - please restart the VM." -Subject "Warning for vm $vmname" -recipient $userstonotify
    }
    Remove-item -path "C:\scripting\notify.csv"
}
findowner
findoldvms
sendnotification