#script allows to associate users with VMs. Search is based on events.
function findowner{
    param(
        [parameter(Mandatory = $true)]
        $sup2admcred,
        [parameter(Mandatory = $true)]
        $chvhost
    )
    class DiskRecord{
        [string]$Disk
        [string]$User
        [string]$UserName
        [string]$VmId
        [string]$VMname
        [string]$hvhost
        DiskRecord($Disk, $User, $UserName, $VmId, $VMname, $HVhost){
            $this.Disk = $Disk
            $this.User = $User
            $this.UserName = $UserName
            $this.VmId = $VmId
            $this.VMname = $VMname
            $this.hvhost = $hvhost
        }
    }
    #creating session for DC and hv host
    Write-Host ("Searching for VM owners on host $chvhost") -BackgroundColor Yellow -ForegroundColor Black
    Try{
        $dcsession = New-PSSession -ComputerName spbsupdc02.support2.veeam.local -Credential $sup2admcred
        $hvsession = New-PSSession -ComputerName $chvhost -Credential $sup2admcred -ErrorAction Stop
    }
    catch{
        Write-host("Failed to create PSSession") -BackgroundColor Red
        break
    }
    #collecting events about successful disk creation for last 7 days.
    $vmmsoper = Invoke-Command -Session $hvsession -ScriptBlock {param ($pc,$cred) Get-WinEvent -LogName "Microsoft-Windows-Hyper-V-VMMS-Operational" -ComputerName $pc -Credential $cred | where {$_.timecreated -ge (Get-Date).AddDays(-70)} | where {$_.id -eq 27311}} -ArgumentList $chvhost,$sup2admcred
    $vmsonhost = Invoke-Command -ScriptBlock {get-vm} -Session $hvsession
    if ($vmsonhost -ne $null){
        $tocsv = @()
        foreach ($event in $vmmsoper){
            #getting disk path from event
            $msg = $event.message.substring(33)
            $message = $msg.substring(0,$msg.length-2)
            foreach ($vm in $vmsonhost){
                #getting vm disk path
                $vmdrive = Invoke-Command -Session $hvsession -ScriptBlock {get-vhd -vmid $args[0]} -ArgumentList $vm.VMId
                #Checking if VM is on snapshot
                if($vmdrive.ParentPath -eq $null){
                    $disklocation = $vmdrive.path
                }
                else{
                    $disklocation = $vmdrive.ParentPath
                }
                #comparing if some vm has disk mentioned in event
                if ($disklocation -eq $message){
                    $vm.HardDrives[0].Path
                    $cvmid = $vm.vmid
                    $cvmname = $vm.name
                    #trying to find user by SID, if not found skiping
                    Try{
                        if($event.userid -eq "S-1-5-18"){
                            $usrname.Name = "System"    
                        }
                        Else{
                            $usrname = Invoke-Command -Session $dcsession -ScriptBlock {param($sid) Get-ADUser -Server support2.veeam.local -Identity $sid} -Args $event.userid
                        }
                    }
                    Catch{
                        break
                    }
                    $obj = New-Object DiskRecord($message,$event.userid, $usrname.name, $cvmid, $cvmname, $chvhost)
                    $tocsv += $obj
                    break
                }
            }
        }
        #checking if we need to create file or add to an existing file.
        if(Test-path "C:\scripting\vmlist_$chvhost.csv"){
            $tocsv | ConvertTo-Csv -NoTypeInformation | add-Content -Path "C:\scripting\vmlist_$chvhost.csv" 
        }
        else{
            $tocsv | ConvertTo-Csv -NoTypeInformation | set-Content -Path "C:\scripting\vmlist_$chvhost.csv" 
        }
        $file = Import-Csv -Path "C:\scripting\vmlist_$chvhost.csv" | sort disk,user -Unique
        $file | ConvertTo-Csv -NoTypeInformation | set-Content -Path "C:\scripting\vmlist_$chvhost.csv" 
    }
    else{
        Write-Host("There are no VMs on host $hvhost") -BackgroundColor Red
    }
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
    Write-Host ("Searching for forgotten VMs  on host $hvhost") -BackgroundColor Yellow -ForegroundColor Black
    $hvsession = New-PSSession -ComputerName $hvhost -Credential $sup2admcred
    $vms = Invoke-Command -ScriptBlock {Get-VM | where {$_.state -eq 'off'}} -Session $hvsession
    $oldvms = @()
    $tonotify = @()
    foreach ($vm in $vms){
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
                    if($string.username -eq "System"){
                        break
                    }
                    else{
                        $tonotify += $string
                        break 
                    }
                }
            }
        }
    }
    $vms = Invoke-Command -Session $hvsession -ScriptBlock {Get-VM | where {$_.state -ne 'off'}}
    foreach ($vm in $vms){
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
                    if($string.username -eq "System"){
                        break
                    }
                    else{
                        $tonotify += $string
                        break 
                    }
                }
            }   
        }
    }
    #checking if need to create file or add info to it
    if(Test-Path C:\scripting\notify.csv){
        $tonotify | ConvertTo-Csv -NoTypeInformation | add-Content -Path "C:\scripting\notify.csv" 
    }
    else{
        $tonotify | ConvertTo-Csv -NoTypeInformation | Set-Content -Path "C:\scripting\notify.csv" 
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
        [PScredential]$AMUSTcreds
    )
    class UserVMsToNotify{
        [string]$Vmname
        [string]$hvhost
        UserVMsToNotify($vmname, $hvhost){
            $this.vmname = $VMname
            $this.hvhost = $hvhost
        }
    }
    $vmtable = @()
    $users = $file | select UserName -Unique
    foreach ($user in $users){
        $uservms = $file | where {$_.username -match $user.UserName}
        foreach ($uservm in $uservms){
            $obj = New-Object UserVMsToNotify($uservm.vmname, $uservm.hvhost)
            $vmtable += $obj
        }
        try{
            $amustuseremail = FindUserFromSupport2inAMUST -username $user.UserName -AMUSTcreds $AMUSTcreds
            #SendEmailFromOutlook -CustomTextAtTheBegining "This is a final warning" -HTMLBodyFragment ($vmtable | ConvertTo-Html -Fragment) -Subject "Your VMs will be deleted soon" -recipient $amustuseremail
        }
        catch{
            $username = $user.UserName
            Write-Host ("Failed to find user e-mail for user $username") -BackgroundColor Red 
        }
        $vmtable = $null
    }
}
#Get-Credential | Export-Clixml -Path C:\scripting\amustcred.xml
#loading credentials
$sup2admcred = Import-Clixml -Path C:\scripting\sup2admcred.xml
$amustcred = Import-Clixml -Path C:\scripting\amustcred.xml
#getting list of hosts to connect, you can just import a csv with host fqdn.
$hosts = "hv2012r2n2.main.support2.veeam.local","hv2012r2n1.main.support2.veeam.local"
foreach($hvhost in $hosts){
    findowner -sup2admcred $sup2admcred -chvhost $hvhost
#loading vmlist.csv
    Try{
        $vmlist = Import-Csv -Path "C:\scripting\vmlist_$hvhost.csv"
        findoldvms -vmlist $vmlist -sup2admcred $sup2admcred -hvhost $hvhost
    }
    Catch{
        Write-Host ("vmlist_$hvhost.csv couldn't be found")
    }
#loading notify.csv
}
Try{
    $notify = Import-Csv -Path "C:\scripting\notify.csv"
    sendnotification -file $notify -AMUSTcreds $amustcred
    #deleting file after notifications are sent
    $dat = (get-date -UFormat "%y.%m.%d.%H.%m").ToString()
    Rename-item -path "C:\scripting\notify.csv" -NewName "report_$dat.csv"
}
Catch{
    Write-Host ("There are no users to warn") -BackgroundColor Green -ForegroundColor Black
}