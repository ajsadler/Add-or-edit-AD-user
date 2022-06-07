# Clears current variables
Remove-Variable * -ErrorAction SilentlyContinue
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

$xl = New-Object -comobject Excel.Application
$xlbook = $xl.Workbooks.Open("\\hgroup\data\GB-IGDoors\zPermissions\Department-Data_Permissions_IGDoors.xlsx")
$xlsheet = $xlbook.Sheets.Item("Permissions")
$xlsheet.Activate()

$csv = Import-Csv -Path "\\hgroup\data\GB-IGDoors\zPermissions\zPermissions to update.csv"
$csvRowAmount = ($csv | Measure-Object).Count # How many new rows are in the csv / used for stating how many times to run through the 'for' function
   
    for ($i=0; $i -le ($csvRowAmount-1); $i++) # For each row in the csv...
    {
        $newUsername =       @($csv."Username")[$i]
        $newFirstName =      @($csv."First Name")[$i]
        $newSurname =        @($csv."Surname")[$i]
        $newDepartmentAD =   @($csv."Department")[$i]

        # Find the row where the department header is located
        $deptInsertPosition = $xlsheet.Range('C1:C1000').Find("*${newDepartmentAD}-rw", [Type]::Missing,[Type]::Missing,1).Row

        # Define a range, from where the department header is located, to the next department header
        $rangeToCheckUserExists = "C" + $deptInsertPosition + ":C" + $xlsheet.Range("C${deptInsertPosition}:C1000").Find("HGROUP\*", [Type]::Missing,[Type]::Missing,1).Row
        $userExists = $xlsheet.Range($rangeToCheckUserExists).Find($newUsername,[Type]::Missing,[Type]::Missing,1).Row
        if ($userExists -gt 0) {continue} # Check if the user exists already in that range, so duplicates aren't added

        $deptInsertPosition = $deptInsertPosition + 1 # To insert the row underneath the department header
        $whereToInsert = $xl.Range("C" + $deptInsertPosition).EntireRow # Insert the row
        [void] $whereToInsert.Insert(1)

        $xlsheet.Cells.Item($deptInsertPosition,3) = $newUsername # Add the names
        $xlsheet.Cells.Item($deptInsertPosition,4) = $newFirstName
        $xlsheet.Cells.Item($deptInsertPosition,5) = $newSurname
        
        $highlightRange = "A${deptInsertPosition}:E${deptInsertPosition}";
        $xlsheet.Range($highlightRange).Interior.Color = "&H33FF66" # Highlight the row green
    }

Set-Content "\\hgroup\data\GB-IGDoors\zPermissions\zPermissions to update.csv" -Value "Username,Department,First Name,Surname" # Reset the csv once complete

$xl.DisplayAlerts = $false; # Hides the pop-up dialog for saving, defaults to Yes
$xlbook.Save()
$xlbook.Close() # Important, otherwise the file is still open and becomes read-only
$xl.Quit()