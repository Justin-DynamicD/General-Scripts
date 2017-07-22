write "Adding all required Roles and Features..."
Try
    {
    Add-WindowsFeature Hyper-V, Multipath-IO -includemanagementtools -ErrorAction STOP -WarningAction STOP
    }
Catch
    {
    write "please take the appropriate action then re-run this script"
    }

write "sorting and renaming all network adapters..."
$AllAdapters = (Get-NetAdapter)
$All10GAdapters = ($AllAdapters|where{$_.LinkSpeed -eq "10 Gbps" -and $_.InterfaceDesription -notmatch 'Hyper-V*'}|sort-object MacAddress)
$All1GAdapters = ($AllAdapters|where{$_.LinkSpeed -eq "1 Gbps" -and $_.InterfaceDesription -notmatch 'Hyper-V*'}|sort-object MacAddress)
$i = 0
$All10GAdapters | ForEach-Object {
    Rename-NetAdapter -Name $_.Name -NewName "lbfo_10g_p$i"
    $i++
    }
$i = 0
$All1GAdapters | ForEach-Object {
    Rename-NetAdapter -Name $_.Name -NewName "lbfo_1g_p$i"
    $i++
    }

write "Creating LBFO teams and all virtual switches and adapters..."
If (!(Get-NetLBFOTeam lbfo_10g))
    {
    New-NetLbfoTeam -name lbfo_10g -TeamMembers lbfo_10g_p* -LoadBalancingAlgorithm Dynamic -TeamingMode SwitchIndependent -TeamNicName lbfo_10g_team
    New-NetLbfoTeam -name lbfo_1g -TeamMembers lbfo_1g_p* -LoadBalancingAlgorithm Dynamic -TeamingMode SwitchIndependent -TeamNicName lbfo_1g_team
    New-VMSwitch vs_10g -MinimumBandwidthMode Weight -NetAdapterName lbfo_10g_team -AllowManagementOS 0
    New-VMSwitch vs_1g -MinimumBandwidthMode Weight -NetAdapterName lbfo_1g_team -AllowManagementOS 0

    add-vmnetworkadapter -ManagementOS -name vnic_mgmt -Switchname vs_1g
    add-vmnetworkadapter -ManagementOS -name vnic_csv10g -Switchname vs_10g
    add-vmnetworkadapter -ManagementOS -name vnic_csv1g -Switchname vs_1g
    add-vmnetworkadapter -ManagementOS -name vnic_lm -Switchname vs_10g

    Set-vmnetworkadaptervlan -managementOS -VMNetworkAdapterName vnic_mgmt -Access -VlanId 115
    Set-vmnetworkadaptervlan -managementOS -VMNetworkAdapterName vnic_csv10g -Access -VlanId 161
    Set-vmnetworkadaptervlan -managementOS -VMNetworkAdapterName vnic_csv1g -Access -VlanId 164
    Set-vmnetworkadaptervlan -managementOS -VMNetworkAdapterName vnic_lm -Access -VlanId 160

    Set-vmnetworkadapter -ManagementOS -name vnic_mgmt -MinimumBandwidthWeight 5
    Set-vmnetworkadapter -ManagementOS -name vnic_csv10g -MinimumBandwidthWeight 40
    Set-vmnetworkadapter -ManagementOS -name vnic_lm -MinimumBandwidthWeight 20
    Set-VMSwitch vs_10g -DefaultFlowMinimumBandwidthWeight 10
    Set-VMSwitch vs_1g -DefaultFlowMinimumBandwidthWeight 30
    }
write "Checking IP Address for remaining tasks..."
$IPAddress="10.8.115.21"
$Gateway="10.8.115.1"
$Prefix=24
$DNS="10.8.16.24","10.8.16.25"
Get-NetAdapter "VEthernet (vnic_mgmt)"| NEW-NetIPAddress -IPAddress $IPAddress -Defaultgateway $Gateway -Prefixlength $Prefix
Set-DNSClientServerAddress -InterfaceAlias "VEthernet (vnic_mgmt)" -ServerAddresses $DNS

write "Completed!"