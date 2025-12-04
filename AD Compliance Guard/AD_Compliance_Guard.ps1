#Enter the Active Directory OU to search in the variable
$AllComputers = Get-ADComputer -LDAPFilter "(name=*)" -SearchBase "OU=sample,DC=yoink4cm,DC=com" -SearchScope Subtree

$Output = [System.Collections.Generic.List[PSObject]]::new() 
# using a generic list is more efficient with big collections of objects ( vs array)

foreach ($Computer in $AllComputers)
    {
     $Owner = (Get-ADComputer $Computer -Properties NTSecurityDescriptor).NTSecurityDescriptor.owner
     #Building a PSObject with all the properties need
     $obj = [PSCustomObject]@{ComputerName = $Computer.Name
               Owner= $owner
               }
    # Adding $Obj to $Output using the .add method
    $Output.add($Obj)
    }
$Output | export-csv ComputerDomainJoinReport.csv -notypeinformation