<#
	.SYNOPSIS
		Hörmann set fileserver permissions
	.DESCRIPTION
		This script is used to check a given permissiontable (Excel-File) for permission-entries, 
        scan the responsible AD-groups for needed changes, write a report of the needed changes.
        The Script can also alter the groups and permissions.
			
	    The script is testet with Windows 2012.
		
		The user which executes this script must have sufficient rights (Domain Admin for example)
	.PARAMETER logfile
		The filename and path of the logfile. 
	.EXAMPLE
		PS C:\> .\set-fileserverpermissions.ps1 -ExcelFile <PathToFile> -AddGroupmembers
	.EXAMPLE
		PS C:\> .\set-fileserverpermissions.ps1 -logfile C:\logs\add-guestaccount.log
	.NOTES
        Version 0.1
        2015-03-26: R. Weickenmeier                      
		            

#>
# ToDo: Fehlende Gruppen automatisch anlegen lassen -> Benötigt OU-Angabe


#--------------------------------------------------------------------------------------------------------------------------
# Parameter Defintion:
#
#     $ExcelFile     Path to the Excel file with folder permissions. Will be converted to an CSV file in the same folder.
#
#                    Column 2 and 3 in this Excel file contains group and user names. Column 2 is normally used for ReadOnly 
#                    permissions and column 3 is normally used for ReadWrite permissions. Both columns are processed separately.
#
#                    If an entry with prefix "HGROUP\" is found, then this entry is assumed to be a permissions group 
#                    for a folder. All entries till the next permission group are assumed to be existing user or group  
#                    objects in the active directory. These users and groups will be added to the permission group.
#
#                    The entry "HGROUP\PermissionEnd" defines the end of the permission definitions and is of course not 
#                    a valid permission group in the active directory.
#
#     $Reportfile    x
#
#     $Logfile       x
#
#     $AddGroupmembers         Add users to groups, if their membership is definded in the $ExcelFile. 
#                              Without these flag set to TRUE, missing memberships are only reported.
#
#     $RemoveGroupmembers      Remove users from groups, if their membership is not definded in the $ExcelFile. 
#                              Without these flag set to TRUE, wrong memberships are only reported.
#
#     $CreateNotfoundGroups    Create groups that are defined in the $ExcelFile, but do not exist in the active directory. 
#                              Without these flag set to TRUE, missing groups are only reported.
#
#     $AddListGroupmembers     Add all groups to the list group (zzd-....-....-list), that are defined in the $ExcelFile.
#                              Without these flag set to TRUE, missing groups are only reported.
#
#     $RemoveListGroupmembers  Remove all groups from the list group (zzd-....-....-list), that are not in the $ExcelFile.
#                              Without these flag set to TRUE, wrong memberships are only reported.
#
#--------------------------------------------------------------------------------------------------------------------------
[CmdletBinding(SupportsShouldProcess=$True)]
param (
    [Parameter(Position=0, Mandatory=$false)]
        [String] $ExcelFile="T:\gb-coalville\zPermissions\Department-Data_Permissions_Hoermann_UK.xlsx",
	[Parameter(Position=1, Mandatory=$false)]
		[String] $Reportfile="C:\HGROUP\LOGS\$(get-date -uformat "%Y%m%d-%H%M")-Fileserver_Permission-change-report-$($env:USERNAME).log",
    [Parameter(Position=2, Mandatory=$false)][Alias('Log')]
        [String] $Logfile="C:\HGROUP\LOGS\$(get-date -uformat "%Y%m%d-%H%M")-Fileserver_Permission-change_errors-$($env:USERNAME).log",
    [Parameter(Position=3, Mandatory=$false)]
		[Switch] $AddGroupmembers,
    [Parameter(Position=4, Mandatory=$false)]
		[Switch] $RemoveGroupmembers,
    [Parameter(Position=5, Mandatory=$false)]
        [Switch] $CreateNotfoundGroups,
    [Parameter(Position=6, Mandatory=$false)]
		[Switch] $AddListGroupmembers,
    [Parameter(Position=7, Mandatory=$false)]
		[Switch] $RemoveListGroupmembers
)

import-module activedirectory

#--------------------------------------------------------------------------------------------------------------------------
# Function to convert the excel input file into an csv file
#--------------------------------------------------------------------------------------------------------------------------
function ConvertFrom-XLx {
  param ([parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
         [string]$path, 
         [switch]$PassThru
        )
 
  begin { $objExcel = New-Object -ComObject Excel.Application }
Process { if ((test-path $path) -and ( $path -match ".xl\w*$")) {
                    $path = (resolve-path -Path $path).path 
                $savePath = $path -replace ".xl\w*$",".csv"
              $objworkbook=$objExcel.Workbooks.Open( $path)
              $objworkbook.SaveAs($savePath,6) # 6 is the code for .CSV 
              $objworkbook.Close($false) 
              #
              # if ($PassThru) {Import-Csv -Path $savePath -delimiter ';'} Else {return $savepath}
              # 
              #    - When users culture is en-US, then converting from excel ends up with an delimiter "," in the csv file.
              #    - When users culture is de-DE, then converting from excel ends up with an delimiter ";" in the csv file.
              #    - To handle that properly the -UseCulture parameter can be used.
              #
              if ($PassThru) {Import-Csv -Path $savePath -UseCulture} Else {return $savepath}
          }
          else {Write-Host "$path : not found"} 
        } 
   end  { $objExcel.Quit() }
}


#--------------------------------------------------------------------------------------------------------------------------
# 1. Initialize Logfile & Reportfile
#--------------------------------------------------------------------------------------------------------------------------
Write-Verbose "Initialising Logfile" 
IF (!(Test-Path $logfile)) {
    New-Item -ItemType File -Path $logfile -Force | Out-Null # Creates empty logfile and all folders (if needed)
	if (!$PSBoundParameters['Whatif']) {
        IF (!(test-path $logfile)) {write-error "Logfile could not be created. Exiting"; exit 1}
    }
} 
Add-Content -Path $Logfile -Value "Error Logfile from $(get-date)"
Write-Verbose "Initialising Reportfile" 
IF (!(Test-Path $Reportfile)) {
    New-Item -ItemType File -Path $Reportfile -Force | Out-Null # Creates empty logfile and all folders (if needed)
	if (!$PSBoundParameters['Whatif']) {
        IF (!(test-path $Reportfile)) {write-error "Reportfile could not be created. Exiting"; exit 1}
    }
} 
Add-Content -Path $Reportfile -Value "Fileserver Permission Report from $(get-date)"


#--------------------------------------------------------------------------------------------------------------------------
# 2. Check if CSV is already present and delete it
#--------------------------------------------------------------------------------------------------------------------------
If (Test-Path($ExcelFile -replace ".xl\w*$",".csv")) {
    Write-Verbose "CSV-File $($ExcelFile -replace ".xl\w*$",".csv") was found. Deleting..."
    Remove-Item $($ExcelFile -replace ".xl\w*$",".csv")
}


#--------------------------------------------------------------------------------------------------------------------------
# 3. Convert Excel File to CSV, import data and check if the excel file is created with a supported version.
#--------------------------------------------------------------------------------------------------------------------------
#
# Old import:
#
#    AllContent = import-csv -path $CSVfile -Delimiter ';' -encoding Default -header "Folder","ReadOnly","ReadWrite","Surname" ,"GivenName"
# 
#    - When users culture is en-US, then converting from excel ends up with an delimiter "," in the csv file.
#    - When users culture is de-DE, then converting from excel ends up with an delimiter ";" in the csv file.
#    - To handle that properly the -UseCulture parameter can be used.
#
$CSVfile = ConvertFrom-XLx $ExcelFile 
$AllContent = import-csv -path $CSVfile -UseCulture -encoding Default -header "Folder","ReadOnly","ReadWrite","Surname" ,"GivenName"
$UserOU = @($allContent | Where-Object {$_.folder -eq "AD-OU (users of site):"})[0].ReadOnly
$GroupsOU = @($allContent | Where-Object {$_.folder -eq "AD-OU (permission groups):"})[0].ReadOnly
$ListGroup = @($allContent | Where-Object {$_.folder -eq "List-Group:"})[0].ReadOnly
$ExcelFileFormat = @($allContent | Where-Object {$_.folder -eq "ExcelFileFormat:"})[0].ReadOnly

$permissions = $allContent | Where-Object {$_.ReadOnly -ne "" -or $_.ReadWrite -ne ""}

$msg="ExcelFile: $ExcelFile"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
$msg="UserOUpath: $UserOU"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
$msg="OUpath: $GroupsOU"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
$msg="ListGroup: $ListGroup"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
$msg="ExcelFileFormat: $ExcelFileFormat"; Write-Host $msg; Add-Content -path $Reportfile -value $msg

If (-NOT ($ExcelFileFormat -like "v2.*")) {
    $msg="ExcelFile `"$ExcelFile`" has to be from format version 2.0 or higher."; Write-Host -f red $msg; Add-Content -path $Reportfile -value $msg
    Remove-Item $CSVfile
    exit
}


#--------------------------------------------------------------------------------------------------------------------------
# 4. Fill the list $AllGroupsInCSV with all AD groups that are references in the CSV file (Excel file)
#--------------------------------------------------------------------------------------------------------------------------
$AllGroupsInCSV = @()
$AllGroupsInCSV_temp = $allContent | Where-Object {$_.ReadOnly -like "HGROUP*" -and $_.ReadWrite -like "HGROUP*" -and $_.ReadOnly -ne "HGROUP\PermissionEnd"}

ForEach ($Group in $AllGroupsInCSV_temp) {
    $AllGroupsInCSV += $Group.ReadOnly.Split("\")[1]
    $AllGroupsInCSV += $Group.ReadWrite.Split("\")[1]
    #
    # Groups defined in $ExcelFile are all starting with "HGROUP\". This prefix has to be deleted for further usage.
    #
}
# 
# Remove duplicates
#
$AllGroupsInCSV = @($AllGroupsInCSV | Select-Object -Unique)



#--------------------------------------------------------------------------------------------------------------------------
# 5. Fill the list $AllGroupsInTargetOU with all AD groups the target-OU $GroupsOU contains.
#--------------------------------------------------------------------------------------------------------------------------
Try {
    $AllGroupsInTargetOU=@(Get-ADGroup -Filter * -SearchBase $GroupsOU | Select-Object @{Name="group";Expression={$_.name}})
} Catch {
    $msg="Fehler beim abrufen der AD-Gruppen in '$GroupsOU'."; Write-Host -f red $msg; Add-Content -path $Reportfile -value $msg
}

#--------------------------------------------------------------------------------------------------------------------------
# 6. Check if all groups in the target-OU $GroupsOU are referenced in $ExcelFile.
#--------------------------------------------------------------------------------------------------------------------------
If ($AllGroupsInTargetOU.Count -gt 0) {
    $DifferentGroups = @(Compare-Object $AllGroupsInTargetOU ($AllGroupsInCSV | Select-Object @{Name="group";Expression={$_}}) -property group | ?{$_.SideIndicator -eq "<="} | Sort group )
    If ($DifferentGroups.Count -gt 0) {
        $msg="Groups that are present in AD, but not listed in Excel-File:"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
        ForEach ($Group in $DifferentGroups) {
            $msg="  $($Group.group)"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
        }
    }
}

#--------------------------------------------------------------------------------------------------------------------------
# 7. Create missing AD groups in target-OU $GroupsOU if group is referenced in $ExcelFile.
#--------------------------------------------------------------------------------------------------------------------------
If ($CreateNotfoundGroups) {
    Write-Host "Checking for missing AD groups..."
    Foreach ($Group in $AllGroupsInCSV) {
        Try {
            Get-ADGroupMember $Group | Out-Null # Throws an exception if the AD-Group cannot be found...
            Write-Host "Found AD group $group."
        } Catch { 
            $msg="AD-Group '$($Group)' not found. Creating AD-Group..."; Write-Host $msg; Add-Content -path $Reportfile -value $msg
            New-ADGroup -Name $Group -SamAccountName $Group -GroupCategory Security -GroupScope Global -DisplayName $Group -Path $GroupsOU
        }
    }
    Write-Host "Checking succedded"
}


#--------------------------------------------------------------------------------------------------------------------------
# 8. Check permissions defined in $ExcelFile and change AD groups in target-OU $GroupsOU if necessary
#
#    Column 2 and 3 in the $ExcelFile contains group and user names. Column 2 is normally used for ReadOnly permissions
#    and column 3 is normally used for ReadWrite permissions. Both columns are processed separately.
#
#    If an entry with prefix "HGROUP\" is found, then this entry is assumed to define a permissions group for a folder. 
#    All entries till the next permission group are assumed to be existing user or group objects in the active directory.
#    These users and groups will be added to the permission group.
#
#    The entry "HGROUP\PermissionEnd" defines the end of the permission definition and is of course not a valid 
#    permission group in the active directory.
#
#--------------------------------------------------------------------------------------------------------------------------

$FirstPermissionGroup = $True
$CurrentReadOnlyGroupCSV = @()
$CurrentReadWriteGroupCSV = @()

ForEach ($Entry in $permissions) {

    #
    # If the prefix "HGROUP\" is found, then a new permission group follows. Otherwise add the user or group
    # to the current permission group.
    #
    If (($Entry.ReadOnly -like "HGROUP\*") -and ($Entry.ReadWrite -like "HGROUP\*")) {

        If ($FirstPermissionGroup) { 
            #
            # The first time a permission group is found there is nothing to do. Only intialize the new groups after this If-Else-Statement 
            # and store all following users and groups for this permission group.
            #
            $FirstPermissionGroup = $false
        }
        Else {
            #
            # The second time a permission group is found and further, a complete permission group from $ExcelFile is read and can be compared against the AD-Group.
            # Start comparing the ReadOnly-Permission-Group.
            #
            $msg="Checking group $CurrentReadOnlyGroupName..."; Write-Host $msg; Add-Content -path $Reportfile -value $msg
            Try {
                $CurrentReadOnlyGroupAD = @(Get-ADGroupMember $CurrentReadOnlyGroupName)
            } 
            Catch {
                $msg="AD-Group '$CurrentReadOnlyGroupName' not found. Check ExcelFile and AD-groups. If everything is correct, please start this script with flag '-CreateNotFoundGroups'."; Write-Warning $msg; Add-Content -path $Reportfile -value $msg
                break
            }

            $DifferentMembers = @(Compare-Object $CurrentReadOnlyGroupAD $CurrentReadOnlyGroupCSV -property SamAccountName)

            If ($DifferentMembers.Count -gt 0) {
                $UsersToRemove = @($DifferentMembers |?{$_.SideIndicator -eq "<="})
                $UsersToAdd = @($DifferentMembers |?{$_.SideIndicator -eq "=>"})    
                If ($UsersToAdd.Count -gt 0) {           
                    If ($AddGroupmembers) {   
                        # Add missing group members
                        $msg="  These users will be added:"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
                        $UsersToAdd | %{
                            $msg="    Adding $($_.SamAccountName) to AD Group..."; Write-Host -f green $msg; Add-Content -path $Reportfile -value $msg
                            Try {
                                Add-ADGroupMember -Identity $CurrentReadOnlyGroupName -Members $_.SamAccountName
                            } Catch {$msg="    Error adding User to AD Group. $($Error[0])"; Write-Error $msg; $Error.Clear; Add-Content -path $Reportfile -value $msg}
                        }
                    } Else {
                        # Report Only
                        $msg="  These users must be added:"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
                        $UsersToAdd | %{$msg="    $($_.SamAccountName)"; Write-Host -f green $msg; Add-Content -path $Reportfile -value $msg}
                    }
                }
                If ($UsersToRemove.Count -gt 0) {
                    If ($RemoveGroupmembers) {
                        # Remove group members not in permission list
                        $msg="  These users will be removed:"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
                        $UsersToRemove | %{
                            $msg="    Removing $($_.SamAccountName) from AD Group"; Write-Host -f red $msg; Add-Content -path $Reportfile -value $msg
                            Try {
                                Remove-ADGroupMember -Identity $CurrentReadOnlyGroupName -Members $_.SamAccountName -Confirm:$false
                            } Catch {$msg="    Error removing AD User from Group. $($Error[0])"; Write-Error $msg; $Error.Clear; Add-Content -path $Reportfile -value $msg}
                        }
                    } Else {
                        # Report Only
                        $msg="  These users must be removed:"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
                        $UsersToRemove | %{$msg="    $($_.SamAccountName)"; Write-Host -f red $msg; Add-Content -path $Reportfile -value $msg}
                    }

                }
            } Else {
                $msg="  No changes needed."; Write-Host $msg; Add-Content -path $Reportfile -value $msg
            }



            #
            # Start comparing the ReadWrite-Permission-Group.
            #
            $msg="Checking group $CurrentReadWriteGroupName..."; Write-Host $msg; Add-Content -path $Reportfile -value $msg
            Try {
                $CurrentReadWriteGroupAD = @(Get-ADGroupMember $CurrentReadWriteGroupName)
            } 
            Catch {
                $msg="AD-Group '$CurrentReadWriteGroupName' not found. Check ExcelFile and AD-groups. If everything is correct, please start this script with flag '-CreateNotFoundGroups'."; Write-Warning $msg; Add-Content -path $Reportfile -value $msg
                break
            }
            $DifferentMembers = @(Compare-Object $CurrentReadWriteGroupAD $CurrentReadWriteGroupCSV -property SamAccountName)

            If ($DifferentMembers.Count -gt 0) {
                $UsersToRemove = @($DifferentMembers |?{$_.SideIndicator -eq "<="})
                $UsersToAdd = @($DifferentMembers |?{$_.SideIndicator -eq "=>"})    
                If ($UsersToAdd.Count -gt 0) {           
                    If ($AddGroupmembers) {   
                        # Add missing group members
                        $msg="  These users will be added:"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
                        $UsersToAdd | %{
                            $msg="    Adding $($_.SamAccountName) to AD Group..."; Write-Host -f green $msg; Add-Content -path $Reportfile -value $msg
                            Try {
                                Add-ADGroupMember -Identity $CurrentReadWriteGroupName -Members $_.SamAccountName
                            } Catch {$msg="    Error adding User to AD Group. $($Error[0])"; Write-Error $msg; $Error.Clear; Add-Content -path $Reportfile -value $msg}
                        }
                    } Else {
                        # Report Only
                        $msg="  These users must be added:"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
                        $UsersToAdd | %{$msg="    $($_.SamAccountName)"; Write-Host -f green $msg; Add-Content -path $Reportfile -value $msg}
                    }
                }
                If ($UsersToRemove.Count -gt 0) {
                    If ($RemoveGroupmembers) {
                        # Remove group members not in permission list
                        $msg="  These users will be removed:"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
                        $UsersToRemove | %{
                            $msg="    Removing $($_.SamAccountName) from AD Group"; Write-Host -f red $msg; Add-Content -path $Reportfile -value $msg
                            Try {
                                Remove-ADGroupMember -Identity $CurrentReadWriteGroupName -Members $_.SamAccountName -Confirm:$false
                            } Catch {$msg="    Error removing AD User from Group. $($Error[0])"; Write-Error $msg; $Error.Clear; Add-Content -path $Reportfile -value $msg}
                        }
                    } Else {
                        # Report Only
                        $msg="  These users must be removed:"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
                        $UsersToRemove | %{$msg="    $($_.SamAccountName)"; Write-Host -f red $msg; Add-Content -path $Reportfile -value $msg}
                    }

                }
            } Else {
                $msg="  No changes needed."; Write-Host $msg; Add-Content -path $Reportfile -value $msg
            }

        }

        #
        # Initialize Name and List for the found permission list.
        #

        $CurrentReadOnlyGroupName = $Entry.ReadOnly.Split("\")[1]
        $CurrentReadWriteGroupName = $Entry.ReadWrite.Split("\")[1]

        $CurrentReadOnlyGroupCSV = @()
        $CurrentReadWriteGroupCSV = @()

    }
    Else {
        If ($Entry.ReadOnly -ne "") { $CurrentReadOnlyGroupCSV += New-Object PSObject -Property @{SamAccountName = $Entry.ReadOnly}   }
        If ($Entry.ReadWrite -ne ""){ $CurrentReadWriteGroupCSV += New-Object PSObject -Property @{SamAccountName = $Entry.ReadWrite} }
    }

    #
    # If the end of the permissions defintions is found, then leave the loop
    #
    If ($Entry.ReadOnly -like "HGROUP\PermissionEnd") {
        break
    }

}


#--------------------------------------------------------------------------------------------------------------------------
# 9. Add all groups found in $ExcelFile to the list-group for that folder structure
#--------------------------------------------------------------------------------------------------------------------------
Try {
    $ListGroupMembers = @()
    $ListGroupMembersTMP = @(Get-ADGroupMember $ListGroup | where {$_.objectclass-eq “group”} | Select-Object SamAccountName)
    ForEach ($Group in $ListGroupMembersTMP) {
        $ListGroupMembers += $($Group.SamAccountName)
        # $msg="ListGroupMember '$($Group.SamAccountName)' ...";Write-Host -f green $msg; Add-Content -path $Reportfile -value $msg
    }

    $msg="Checking group '$ListGroup' ...";Write-Host $msg; Add-Content -path $Reportfile -value $msg

    $DifferentMembers = @(Compare-Object $ListGroupMembers $AllGroupsInCSV)

    If ($DifferentMembers.Count -gt 0) {
        $GroupsToRemove = @($DifferentMembers |?{$_.SideIndicator -eq "<="})
        $GroupsToAdd = @($DifferentMembers |?{$_.SideIndicator -eq "=>"})

        IF ($GroupsToAdd.Count -gt 0) {           
            If ($AddListGroupmembers) {   
                # Add missing ListGroup members
                $msg="  These groups will be added:"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
                $GroupsToAdd | %{
                    $msg="    Adding $($_.InputObject) to List Group..."; Write-Host -f green $msg; Add-Content -path $Reportfile -value $msg
                    Try {
                        Add-ADGroupMember -Identity $ListGroup -Members $_.InputObject
                    } Catch {$msg="    Error adding AD User to List Group. $($Error[0])"; Write-Error $msg; $Error.Clear; Add-Content -path $Reportfile -value $msg}
                }
            } Else {
                # Report Only
                $msg="  These groups must be added:"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
                $GroupsToAdd | %{$msg="    $($_.InputObject)"; Write-Host -f green $msg; Add-Content -path $Reportfile -value $msg}
            }
        }   
        If ($GroupsToRemove.Count -gt 0) {
            If ($RemoveListGroupmembers) {
                # Remove group members not in permission list
                $msg="  These groups will be removed:"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
                $GroupsToRemove | %{
                    $msg="    Removing $($_.InputObject) from List Group"; Write-Host -f red $msg; Add-Content -path $Reportfile -value $msg
                    Try {
                        Remove-ADGroupMember -Identity $ListGroup -Members $_.InputObject -Confirm:$false
                    } Catch {$msg="    Error removing AD User from List Group. $($Error[0])"; Write-Error $msg; $Error.Clear; Add-Content -path $Reportfile -value $msg}
                }
            } Else {
                # Report Only
                $msg="  These users must be removed:"; Write-Host $msg; Add-Content -path $Reportfile -value $msg
                $GroupsToRemove | %{$msg="    $($_.InputObject)"; Write-Host -f red $msg; Add-Content -path $Reportfile -value $msg}
            }
        }    
               
    } Else {
        $msg="  No changes needed."; Write-Host $msg; Add-Content -path $Reportfile -value $msg
    }


} Catch {
    $msg="AD-Gruppe '$ListGroup' nicht gefunden. Kontrollieren sie in der Excel-Datei die Zeile List-Groups und passen Sie diese ggf. entsprechend an."; Write-Error $msg; $Error.Clear; Add-Content -path $Reportfile -value $msg
}


#--------------------------------------------------------------------------------------------------------------------------
# Clean up
#--------------------------------------------------------------------------------------------------------------------------
Remove-Item $CSVfile




