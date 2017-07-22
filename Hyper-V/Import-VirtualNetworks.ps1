
function Import-VirtualNetworks
{
    param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][String]$VLAN,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][String]$Network,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][String]$Purpose,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][String]$Site
    )

    If(!(Get-SCLogicalNetwork $Purpose)){
        $logicalNetwork = New-SCLogicalNetwork -Name $Purpose -LogicalNetworkDefinitionIsolation $false -EnableNetworkVirtualization $false -UseGRE $false -IsPVLAN $false
        }
    Else
        {
        $logicalNetwork = Get-SCLogicalNetwork $Purpose
        }

    If(!(Get-SCLogicalNetworkDefinition -Name $Site+'-'+$Purpose)){
        $allSubnetVlan = @()
        }
    Else
        {
        $allSubnetVlan += @((Get-SCLogicalNetworkDefinition -Name $Site+'-'+$Purpose).subnetvlans)
        }
        $allHostGroups = @()
        switch ($Site) {
            snv {$allHostGroups += Get-SCVMHostGroup -ID "4a881951-14a6-4969-b261-2bf8f1414979"}
            lax {$allHostGroups += Get-SCVMHostGroup -ID "9abe020d-e378-43ba-a1b1-09e56f528c2f"}
            mar {$allHostGroups += Get-SCVMHostGroup -ID "9df90a04-7ca9-45d9-850c-4674c2e27ca5"}
            cos {$allHostGroups += Get-SCVMHostGroup -ID "5136cd63-7bab-4c1c-beb5-b1dc2f1b739f"}
            }
        
        $allSubnetVlan += New-SCSubnetVLan -Subnet $Network -VLanID $VLAN
        New-SCLogicalNetworkDefinition -Name $Site+'-'+$Purpose -LogicalNetwork $logicalNetwork -VMHostGroup $allHostGroups -SubnetVLan $allSubnetVlan -RunAsynchronously
}

Import-csv blahfile.csv | Foreach-Object {Import-VirtualNetworks -VLAN $_.VLAN -Network $_.Network -Purpose $_.Purpose -Site $_.Site}