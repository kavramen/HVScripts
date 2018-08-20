#script allows to associate users with VMs. Search is based on events.
function findowner{
    param(
    [parameter(Mandatory = $true)]
    $sup2admcred,
    [parameter(Mandatory = $true)]
    $hvhost
    )
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
    #creating session for DC and hv host
    $dcsession = New-PSSession -ComputerName spbsupdc02.support2.veeam.local -Credential $sup2admcred
    $hvsession = New-PSSession -ComputerName $hvhost -Credential $sup2admcred
    $vmmsoper = Get-WinEvent -LogName "Microsoft-Windows-Hyper-V-VMMS-Operational" -ComputerName hv2012r2n1.main.support2.veeam.local -Credential $sup2admcred | where {$_.id -eq 27311}
    $vmsonhost = Invoke-Command -ScriptBlock {get-vm} -Session $hvsession
    foreach ($event in $vmmsoper)
    {
        #getting disk path from event
        $msg = $event.message.substring(33)
        $message = $msg.substring(0,$msg.length-2)
        foreach ($vm in $vmsonhost)
        {
            #getting vm disk path
            $vmdrive = Invoke-Command -Session $hvsession -ScriptBlock {get-vhd -vmid $args[0]} -ArgumentList $vm.VMId
            #comparing if some vm has disk mentioned in event
            if ($vmdrive.Path -eq $message)
            {
                $vm.HardDrives[0].Path
                $cvmid = $vm.vmid
                $cvmname = $vm.name
                break
            }
            else
            {
                $cvmid=''
                $cvmname=''
            }
        }
        $usrname = Invoke-Command -Session $dcsession -ScriptBlock {Get-ADUser -Server spbsupdc02.support2.veeam.local -Identity $args[0]} -Args $event.userid.value
        $obj = New-Object DiskRecord($message,$event.userid.value, $usrname.name, $cvmid, $cvmname)
        $obj | Export-Csv -Path "C:\scripting\vmlist_$hvhost.csv" -Append -NoTypeInformation
    }
    $file = Import-Csv -Path "C:\scripting\vmlist_$hvhost.csv" | sort disk,user -Unique
    $file | Export-Csv -Path "C:\scripting\vmlist_$hvhost.csv" -NoTypeInformation
    Remove-PSSession -Session $dcsession
    Remove-PSSession -Session $hvsession
}
#Script check when VM was used last time and if it's not forgotten    
function findoldvms{
    param(
    [parameter(Mandatory = $true)]
    $sup2admcred,
    [parameter(Mandatory = $true)]
    $hvhost,
    [parameter(Mandatory = $true)]
    $vmlist
    )
    #creating session for hv host
    $hvsession = New-PSSession -ComputerName $hvhost -Credential $sup2admcred
    $vms = Invoke-Command -ScriptBlock {Get-VM | where {$_.state -eq 'off'}} -Session $hvsession
    $oldvms = @()
    $vmstodelete = @()
    foreach ($vm in $vms) {
        #getting vm disk path
        $vmdrive = Invoke-Command -Session $hvsession -ScriptBlock {get-vhd -vmid $args[0]} -ArgumentList $vm.VMId
        #check if we have drive attached to find lastrun time
        if (!$vmdrive.Path){
            $config =  Invoke-Command -Session $hvsession -ScriptBlock {param ($vmconfig, $vmid) Get-ChildItem -Path $vmconfig -filter "$vmid.xml" -recurse} -ArgumentList $vm.configurationlocation, $vm.id
            $lastwritetime = $config.LastWriteTime
        }
        else{
            $diskFile = Invoke-Command -Session $hvsession -ScriptBlock {Get-Item -Path $args[0]} -ArgumentList $vmdrive.Path
            $lastwritetime = $diskFile.LastWriteTime
        }
        #finding VMs which were not modified for more than 21 day and deleting them
        if ($lastwritetime -lt ((get-date).AddDays(-21))){
            Invoke-Command -Session $hvsession -ScriptBlock {Get-VMHardDiskDrive -VM $args[0] | Foreach { Remove-item -path $_.Path -Recurse -Force -Confirm:$False}} -ArgumentList $vm
            Invoke-Command -Session $hvsession -ScriptBlock {Remove-VM $args[0] -Force -Confirm:$False} -ArgumentList $vm
            break
        }
        #finding VMs which were not modified for more than 17 days and less than 21 day and notifying owner
        elseif ($lastwritetime -lt ((get-date).AddDays(-17)) -AND $lastwritetime -gt ((get-date).AddDays(-21))){
            foreach ($string in $vmlist){
                if ($string.vmid -eq $vm.VMId){
                    $string | Export-Csv -Path "C:\scripting\notify.csv" -NoTypeInformation -Append
                }
            }
        }
    }
    $vms = Invoke-Command -Session $hvsession -ScriptBlock {Get-VM | where {$_.state -ne 'off'}}
    foreach ($vm in $vms) {
        #finding VMs which are running more than 30 days and deleting them
        if ($vm.uptime.days -gt 30){
            Invoke-Command -Session $hvsession -ScriptBlock {Get-VMHardDiskDrive -VM $args[0] | Foreach { Remove-item -path $_.Path -Recurse -Force -Confirm:$False}} -ArgumentList $vm
            Invoke-Command -Session $hvsession -ScriptBlock {Remove-VM $args[0] -Force -Confirm:$False} -ArgumentList $vm
            break
        }
        #finding VMs which are running more than 27 days and less than 30 and notifying owner
        elseif($vm.uptime.days -gt 27 -AND $vm.uptime.days -lt 30){
            foreach ($string in $vmlist){
                if ($string.vmid -eq $vm.VMId){
                    $string | Export-Csv -Path "C:\scripting\notify.csv" -NoTypeInformation -Append
                }
            }   
        }
    }
}
#Script is finding amust 2 users based on support2 UserName
function FindUserFromSupport2inAMUST{
param(
    [parameter(Mandatory=$true)]
    $username,
    [parameter(Mandatory=$true)]
    [PScredential]$AMUSTcreds
)
    $amustuser = Get-ADUser -Filter {Name -eq $UserName} -Properties emailaddress -Credential $AMUSTcreds -Server "amust.local"
    $result = $amustuser.EmailAddress
    return $result
###-SearchBase "OU=Support,OU=Amust,OU=Employees,OU=Accounts,DC=amust,DC=local"
}
#function to send e-mail through Outlook client
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

$o = New-Object -ComObject Outlook.Application

$mail = $o.CreateItem(0)

$mail.subject = $subject
$mail.HTMLBody = $htmlBody
$mail.To = $recipient
$mail.Send()
}
#function is sending notification to users
function sendnotification{
param(
    [parameter(Mandatory=$true)]
    $file,
    [parameter(Mandatory=$true)]
    $hvhost,
    [parameter(Mandatory=$true)]
    [PScredential]$AMUSTcreds
    )
    #Sending notification for each user in file
    foreach ($string in $notify){
        $user = $string.UserName
        $amustuseremail = FindUserFromSupport2inAMUST -username $user -AMUSTcreds $AMUSTcreds
        $vmname = $string.vmname
        SendEmailFromOutlook -CustomTextAtTheBegining "Please note that your vm: $vmname on host: $hvhost will be deleted soon. To avoid this - please restart the VM." -Subject "Warning for vm $vmname" -recipient $amustuseremail
    }
}
#loading credentials
$sup2admcred = Import-Clixml -Path C:\scripting\sup2admcred.xml
$amustcred = Import-Clixml -Path C:\scripting\amustcred.xml
#getting list of hosts to connect, you can just import a csv with host fqdn.
$hosts = "hv2012r2n1.main.support2.veeam.local","hv2012r2n2.main.support2.veeam.local"
foreach($hvhost in $hosts){
    findowner -sup2admcred $sup2admcred -hvhost $hvhost
#loading vmlist.csv
    Try
    {
    $vmlist = Import-Csv -Path "C:\scripting\vmlist_$hvhost.csv"
    }
    Catch
    {
        Write-Host ("vmlist_$hvhost.csv couldn't be found")
    }
    findoldvms -vmlist $vmlist -sup2admcred $sup2admcred -hvhost $hvhost
#loading notify.csv
    Try
    {
        $notify = Import-Csv -Path "C:\scripting\notify.csv"
        sendnotification -file $notify -host $hvhost -AMUSTcreds $amustcred
        #deleting file after notifications are sent
        Remove-item -path "C:\scripting\notify.csv"
    }
    Catch
    {
        Write-Host ("There are no users to warn")
    }
}


