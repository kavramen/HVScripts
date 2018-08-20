#script allows to associate users with VMs. Search is based on events.
$vmmsoper = Get-WinEvent -LogName "Microsoft-Windows-Hyper-V-VMMS-Operational"  | where {$_.id -eq 27311}
class DiskRecord
{
    [string]$Disk
    [string]$User
    [string]$VmId

    DiskRecord($Disk, $User, $VmId) {
       $this.Disk = $Disk
       $this.User = $User
       $this.VmId = $VmId
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
            $s = $vm.vmid
        }
        else
        {
            $s=''
        }
    }
    $obj = New-Object DiskRecord($message,$vmmsoper.userid.accountdomainsid[1].value, $s)
    $obj | Export-Csv -Path "C:\scripting\vmlist.csv" -Append -NoTypeInformation
}
$file = Import-Csv -Path "C:\scripting\vmlist.csv" | sort disk,user -Unique
$file | Export-Csv -Path "C:\scripting\vmlist.csv" -NoTypeInformation
