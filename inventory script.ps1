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
    elseif ($lastwritetime -lt ((get-date).AddDays(-14))){
        foreach ($string in $file){
            if ($string.vmid -eq $vm.VMId){
                $string | Export-Csv -Path "C:\scripting\notify.csv" -NoTypeInformation -Append
            }
        }
    }
}
$vms = Get-VM | where {$_.state -ne 'off'}
$forgottenvms = @()
foreach ($vm in $vms) {
    #finding VMs which are running more than 30 days and deleting them
    if ($vm.uptime.days -gt 30){
        $forgottenvms += $vm.Name
        Get-VMHardDiskDrive -VM $vm | Foreach { Remove-item -path $_.Path -Recurse -Force -Confirm:$False}
        Remove-VM $vm -Force -Confirm:$False
    }
}


