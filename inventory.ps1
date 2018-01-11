Function GetSubscription
{
    param([string] $Id)
    $list = Get-AzureRmSubscription -SubscriptionId $Id | Select SubscriptionId, Name, TenantId
    return $list
}

Function GetResourceGroups
{   
    $list = Get-AzureRmResourceGroup | Select ResourceGroupName, Location, ResourceId
    return $list
}

Function GetStorageAccounts
{
    $list = @()
    
    $storageAccounts = Get-AzureRmStorageAccount
    Foreach($sa in $storageAccounts)
    {
        $v = [PSCustomObject] @{
            Name = ($sa).StorageAccountName
            Location = ($sa).PrimaryLocation
            ResourceGroupName =($sa).ResourceGroupName
            Id = ($sa).Id
            Tier = ($sa).SKU.Tier
            Replication = ($sa).SKU.Name
        }

        $list += $v
    }

    return $list
}

Function GetVirtualMachines
{
    $list = @()
    
    $virtualMachines = Get-AzureRMVM
    Foreach($vm in $virtualMachines)
    {
        $resource = Get-AzureRmResource -ResourceId $vm.Id
        $v = [PSCustomObject] @{
            Name = ($resource).Name
            Location = ($resource).Location
            ResourceGroupName =($resource).ResourceGroupName
            VmSize = ($resource).Properties.HardwareProfile.VmSize
            OsType = ($resource).Properties.StorageProfile.OsDisk.OsType
            OsPublisher = ($resource).Properties.StorageProfile.ImageReference.Publisher
            OsOffer = ($resource).Properties.StorageProfile.ImageReference.Offer
            OsSKU =  ($resource).Properties.StorageProfile.ImageReference.Sku
            PrivateIP = (Get-AzureRmNetworkInterface | ? {$_.ID -match $resource.Name}).IpConfigurations.PrivateIpAddress
            Subnet = ""
            VNet = (Get-AzureRmVirtualNetwork | ? {$_.Subnets.ID -match (Get-AzureRmNetworkInterface | ? {$_.ID -match $resource.Name}).IpConfigurations.subnet.id}).Name
            AvailabilitySet = ""
            DomainJoined = ""
            OU = ""
            SCOM = ""
            Shavlik = ""
            BackupVault = ""
        }

        $list += $v
    }

    return $list | Sort-Object Name
}

Function GetSubnets
{
    #Subnets (name, range, available addresses, nsg)
    $list = @()
    
    $vnets = Get-AzureRmVirtualNetwork
    Foreach($vnet in $vnets){
        Foreach($subnet in $vnet.Subnets)
        {
            $sn = Get-AzureRmVirtualNetworkSubnetConfig -VirtualNetwork $vnet -Name $subnet.Name

            $v = [PSCustomObject] @{
                VNetName = ($vnet).Name
                Location = ($vnet).Location
                ResourceGroupName =($vnet).ResourceGroupName
                SubnetName = ($subnet).Name
                SubnetId = ($subnet).Id
                SubnetRange = ($sn).AddressPrefix
                NsgName = (Get-AzureRmResource -ExpandProperties -ResourceId ($sn).NetworkSecurityGroup.Id).Name
                NsgId = ($sn).NetworkSecurityGroup.Id
            }

            $list += $v
        }
    }

    return $list
}

Function GetAvailabilitySets
{
    $list = @()
    $rgs = Get-AzureRmResourceGroup

    Foreach($rg in $rgs)
    {
        $sets = Get-AzureRmAvailabilitySet -ResourceGroupName $rg.ResourceGroupName

        Foreach($as in $sets)
        {
            #Availability Sets (name, rg, location, members)
            $members = @()
            Foreach($vm in $as.VirtualMachinesReferences)
            {
                $members += (Get-AzureRmResource -ResourceId $vm.Id).Name
            }

            $v = [PSCustomObject] @{
                Name = ($as).Name
                Location = ($as).Location
                ResourceGroupName =($as).ResourceGroupName
                Members = ($members -join ",")
            }          
        }

        $list += $v
    }

    return $list
}

Function GetManagedDisks
{
    $list = @()
    
    $virtualMachines = Get-AzureRMVM

    Foreach($vm in $virtualMachines)
    {
        $resource = Get-AzureRmResource -ResourceId $vm.Id

        $disks = @()
        $disks += ($resource).Properties.StorageProfile.osDisk
        $disks += ($resource).Properties.StorageProfile.dataDisks


        Foreach($d in $disks)
        {
            $v = [PSCustomObject] @{
                Name = ($d).Name
                Location = ($resource).Location
                ResourceGroupName =($resource).ResourceGroupName
                AttchedTo = ($resource).Name
                Size = ($d).diskSizeGB
                Caching = ($d).caching
                Lun = ($d).lun           
            }

            $list += $v        
        }
        
    }

    return $list | Sort-Object Name
}

Function GetNics
{
    $nics = Get-AzureRmNetworkInterface

    Foreach($nic in $nics)
    {
        $v = [PSCustomObject] @{
            Name = ($nic).Name
            ResourceGroupName =($nic).ResourceGroupName
            IP = ($nic).IpConfigurations[0].PrivateIpAddress
        }

        $list += $v   
        
    }
    return $list  
}


Function GenerateReport
{
    param([string] $ReportPath, [string] $SubscriptionId, [string] $TenantId)

    Set-AzureRmContext -SubscriptionId $subscriptionId -TenantId $tenantId
    
    $subscription = GetSubscription -Id $subscriptionId
    $subscription | Export-XLSX -Path $reportPath -WorksheetName "Subscription" -ReplaceSheet -Table -Autofit 

    Write-Output "Exporting RG"
    $resourceGroups = GetResourceGroups
    $resourceGroups | Export-XLSX -Path $reportPath -WorksheetName "Resource Groups" -ReplaceSheet -Table -Autofit 

    Write-Output "Exporting Storage Accounts"
    $storageAccounts = GetStorageAccounts
    $storageAccounts | Export-XLSX -Path $reportPath -WorksheetName "Storage Accounts" -ReplaceSheet -Table -Autofit 

    Write-Output "Exporting Subnets"
    $subnets = GetSubnets
    $subnets | Export-XLSX -Path $reportPath -WorksheetName "Subnets" -ReplaceSheet -Table -Autofit 

    Write-Output "Exporting AS"
    $availabilitySets = GetAvailabilitySets
    $availabilitySets | Export-XLSX -Path $reportPath -WorksheetName "Availability Sets" -ReplaceSheet -Table -Autofit 

    Write-Output "Exporting VMs"
    $virtualMachines = GetVirtualMachines
    $virtualMachines | Export-XLSX -Path $reportPath -WorksheetName "Virtual Machines" -ReplaceSheet -Table -Autofit 

    Write-Output "Exporting Managed Disks"
    $managedDisks = GetManagedDisks
    $managedDisks | Export-XLSX -Path $reportPath -WorksheetName "Managed Disks" -ReplaceSheet -Table -Autofit 

    Write-Output "Exporting NICs"
    $nics = GetNics
    $nics | Export-XLSX -Path $reportPath -WorksheetName "NICs" -ReplaceSheet -Table -Autofit 

    ## Other Resources - Not required by FICT
    # Load Balancers
    # NSG
    # Recovery Services Vault

}

# Connect  
#Login-AzureRmAccount
#Install-Module PSExcel

#DEVTEST 
GenerateReport -ReportPath "C:\Inventory.xlsx" -SubscriptionId "" -TenantId ""

Write-Output "Finished"
    
       
