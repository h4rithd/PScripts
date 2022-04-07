$DomainObject = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$PrimaryDomainController = ($DomainObject.PdcRoleOwner).Name
$DistinguishedName = "DC=$($DomainObject.Name.Replace('.', ',DC='))"
$StringSearch = "LDAP://"
$StringSearch += $PrimaryDomainController + "/"
$StringSearch +=  $DistinguishedName
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]$StringSearch)                   
$objDomain = New-Object System.DirectoryServices.DirectoryEntry
$objSearcher.SearchRoot =  $objDomain   

$objSearcher.filter="samAccountType=805306368"
# Search by Domain Admins --> $objSearcher.filter="memberof=CN=Domain Admins,CN=Users,DC=<DOMAIN>,DC=com"
# For All Computers --> $objSearcher.filter="objectcategory=CN=Computer,CN=Schema,CN=Configuration,DC=<DOMAIN>,DC=com"
# Search by OS --> $objSearcher.filter="operatingsystem=*windows 10*"

$SearchResult = $objSearcher.FindAll()
Foreach($object in $SearchResult)
{
    Foreach($properties in $object.Properties)
    {
        Write-Host "------------------------------------------------"
        $properties
        #$properties.name # if you want single value
    }
}
Write-Host "------------------------------------------------"
