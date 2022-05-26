$device = Read-Host "Enter device name"
$resourceID = (Get-CMDevice -Name $device -Resource | Select-Object ResourceID).ResourceID
Add-CMUserCollectionDirectMembershipRule -CollectionId "HG10114B" -ResourceId $resourceID