# Clears current variables
Remove-Variable * -ErrorAction SilentlyContinue
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

$xl = New-Object -comobject Excel.Application
$xlbook = $xl.Workbooks.Open("\\hgroup\data\GB-IGDoors\zPermissions\Department-Data_Permissions_IGDoors.xlsx")
$xlsheet = $xlbook.Sheets.Item("Permissions")
$xlsheet.Activate()

$csv = Import-Csv -Path "\\hgroup\data\GB-IGDoors\zPermissions\zPermissions to update.csv"
$csvRowAmount = ($csv | Measure-Object).Count # How many new rows are in the csv / used for stating how many times to run through the 'for' function


    function EachDeptRow # On what row is the header for each department
    {
        $script:rowFor3tec =              $xlsheet.Range('C1:C1000').Find('HGROUP\zzd-gbblw-igdoors-3tec-rw',                              [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForDataExchange =      $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-data-exchange-rw",                     [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForProjects =          $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-projects-rw",                          [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForSagePayroll =       $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-sage-payroll-rw",                      [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForAfterSales =        $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-AfterSales-Team-rw",                   [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForCustomerCare =      $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-CustomerCare-Team-rw",                 [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForSeniorManagers =    $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-Senior-Managers-rw",                   [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForTeamLeaders =       $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-Team-Leaders-rw",                      [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForFinance =           $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-Finance-Team-rw",                      [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForSeniorFinance =     $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-SeniorFinance-Team-rw",                [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForItTeam =            $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-IT-Team-rw",                           [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForOrderProc =         $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-OrderProcessing-Team-rw",              [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForRandDteam =         $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-randd-team-rw",                        [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForEngineering =       $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-engineering-team-rw",                  [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForSales =             $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-sales-team-rw",                        [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForSocialHousing =     $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-SocialHousing-Team-rw",                [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForTechnical =         $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-technical-team-rw",                    [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForQuality =           $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-quality-team-rw",                      [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForPublic =            $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-public-rw",                            [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForMaterials =         $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-MaterialsandWarehouse-Management-rw",  [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForTrade =             $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-trade-team-rw",                        [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForLogon =             $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-logon-rw",                             [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForInformation =       $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-Information-Centre-rw",                [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForProduction =        $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-production-rw",                        [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForPlanning =          $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-planning-rw",                          [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForDatabase =          $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-Database-Storage-rw",                  [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForReception =         $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-reception-rw",                         [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForNewFactory =        $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-NewFactory-rw",                        [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForNewBuild =          $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-newbuild-rw",                          [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForNewBuildMan =       $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-newbuild-management-rw",               [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForHumanResSenior =    $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-human-resources-senior-rw",            [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForHumanResources =    $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-human-resources-rw",                   [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForQualityImages =     $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-quality-images-rw",                    [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForDespatchImages =    $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-Despatch-Images-rw",                   [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForCNCTeam =           $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-CNC-Team-rw",                          [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForBridgetime =        $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-bridgetime-rw",                        [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForBridgetimeScans =   $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-Bridgetime-Scans-rw",                  [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForNPDproject =        $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-NPD-Project-Planning-rw",              [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForHealth =            $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-HealthandSafety-rw",                   [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForRMA =               $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-RMA-rw",                               [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForLogistics =         $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-Logistics-rw",                         [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForProjectEngi =       $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-Project-Engineering-rw",               [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForZpermissions =      $xlsheet.Range('C1:C1000').Find("HGROUP\zzd-gbblw-igdoors-zpermissions-rw",                      [Type]::Missing,[Type]::Missing,1).Row
        $script:rowForEnd =               $xlsheet.Range('C1:C1000').Find("HGROUP\PermissionEnd",                                          [Type]::Missing,[Type]::Missing,1).Row
    }

    function EachDeptRange # What is the range for each department (between its row, and the next department's row)
    {
        $script:rangeFor3tec =            "C" + $script:rowFor3tec              + ":C" + $script:rowForDataExchange
        $script:rangeForDataExchange =    "C" + $script:rowForDataExchange      + ":C" + $script:rowForProjects
        $script:rangeForProjects =        "C" + $script:rowForProjects          + ":C" + $script:rowForSagePayroll
        $script:rangeForSagePayroll =     "C" + $script:rowForSagePayroll       + ":C" + $script:rowForAfterSales
        $script:rangeForAfterSales =      "C" + $script:rowForAfterSales        + ":C" + $script:rowForCustomerCare
        $script:rangeForCustomerCare =    "C" + $script:rowForCustomerCare      + ":C" + $script:rowForSeniorManagers
        $script:rangeForSeniorManagers =  "C" + $script:rowForSeniorManagers    + ":C" + $script:rowForTeamLeaders
        $script:rangeForTeamLeaders =     "C" + $script:rowForTeamLeaders       + ":C" + $script:rowForFinance
        $script:rangeForFinance =         "C" + $script:rowForFinance           + ":C" + $script:rowForSeniorFinance
        $script:rangeForSeniorFinance =   "C" + $script:rowForSeniorFinance     + ":C" + $script:rowForItTeam
        $script:rangeForItTeam =          "C" + $script:rowForItTeam            + ":C" + $script:rowForOrderProc
        $script:rangeForOrderProc =       "C" + $script:rowForOrderProc         + ":C" + $script:rowForRandDteam
        $script:rangeForRandDteam =       "C" + $script:rowForRandDteam         + ":C" + $script:rowForEngineering
        $script:rangeForEngineering =     "C" + $script:rowForEngineering       + ":C" + $script:rowForSales
        $script:rangeForSales =           "C" + $script:rowForSales             + ":C" + $script:rowForSocialHousing
        $script:rangeForSocialHousing =   "C" + $script:rowForSocialHousing     + ":C" + $script:rowForTechnical
        $script:rangeForTechnical =       "C" + $script:rowForTechnical         + ":C" + $script:rowForQuality
        $script:rangeForQuality =         "C" + $script:rowForQuality           + ":C" + $script:rowForPublic
        $script:rangeForPublic =          "C" + $script:rowForPublic            + ":C" + $script:rowForMaterials
        $script:rangeForMaterials =       "C" + $script:rowForMaterials         + ":C" + $script:rowForTrade
        $script:rangeForTrade =           "C" + $script:rowForTrade             + ":C" + $script:rowForLogon
        $script:rangeForLogon =           "C" + $script:rowForLogon             + ":C" + $script:rowForInformation
        $script:rangeForInformation =     "C" + $script:rowForInformation       + ":C" + $script:rowForProduction
        $script:rangeForProduction =      "C" + $script:rowForProduction        + ":C" + $script:rowForPlanning
        $script:rangeForPlanning =        "C" + $script:rowForPlanning          + ":C" + $script:rowForDatabase
        $script:rangeForDatabase =        "C" + $script:rowForDatabase          + ":C" + $script:rowForReception
        $script:rangeForReception =       "C" + $script:rowForReception         + ":C" + $script:rowForNewFactory
        $script:rangeForNewFactory =      "C" + $script:rowForNewFactory        + ":C" + $script:rowForNewBuild
        $script:rangeForNewBuild =        "C" + $script:rowForNewBuild          + ":C" + $script:rowForNewBuildMan
        $script:rangeForNewBuildMan =     "C" + $script:rowForNewBuildMan       + ":C" + $script:rowForHumanResSenior
        $script:rangeForHumanResSenior =  "C" + $script:rowForHumanResSenior    + ":C" + $script:rowForHumanResources
        $script:rangeForHumanResources =  "C" + $script:rowForHumanResources    + ":C" + $script:rowForQualityImages
        $script:rangeForQualityImages =   "C" + $script:rowForQualityImages     + ":C" + $script:rowForDespatchImages
        $script:rangeForDespatchImages =  "C" + $script:rowForDespatchImages    + ":C" + $script:rowForCNCTeam
        $script:rangeForCNCteam =         "C" + $script:rowForCNCTeam           + ":C" + $script:rowForBridgetime
        $script:rangeForBridgetime =      "C" + $script:rowForBridgetime        + ":C" + $script:rowForBridgetimeScans
        $script:rangeForBridgetimeScans = "C" + $script:rowForBridgetimeScans   + ":C" + $script:rowForNPDproject
        $script:rangeForNPDproject =      "C" + $script:rowForNPDproject        + ":C" + $script:rowForHealth
        $script:rangeForHealth =          "C" + $script:rowForHealth            + ":C" + $script:rowForRMA
        $script:rangeForRMA =             "C" + $script:rowForRMA               + ":C" + $script:rowForLogistics
        $script:rangeForLogistics =       "C" + $script:rowForLogistics         + ":C" + $script:rowForProjectEngi
        $script:rangeForProjectEngi =     "C" + $script:rowForProjectEngi       + ":C" + $script:rowForZpermissions
        $script:rangeForZpermissions =    "C" + $script:rowForZpermissions      + ":C" + $script:rowForEnd
    }

   
    for ($i=0; $i -lt ($csvRowAmount-1); $i++) # For each row in the csv...
    {
        $script:newUsername =       $csv."Username"[$i]
        $script:newFirstName =      $csv."First Name"[$i]
        $script:newSurname =        $csv."Surname"[$i]
        $script:newDepartmentAD =   $csv."Department"[$i]

        EachDeptRow # Re-load the row and range functions each time
        EachDeptRange # Because if rows are added, then they will have moved

        switch ($script:newDepartmentAD) # Get the positions where to insert the new row, write the name, and highlight the row
        {
            '3tec'
            {
                $rangeForDeptToCheck = $script:rangeFor3tec;
                $deptInsertPosition = $script:rowFor3tec;
                continue
            }
            'aftersales-team'
            {
                $rangeForDeptToCheck = $script:rangeForAfterSales;
                $deptInsertPosition = $script:rowForAfterSales;
                continue
            }
            'Bridgetime-Scans'
            {
                $rangeForDeptToCheck = $script:rangeForBridgetimeScans;
                $deptInsertPosition = $script:rowForBridgetimeScans;
                continue
            }
            'CNC-Team'
            {
                $rangeForDeptToCheck = $script:rangeForCNCteam;
                $deptInsertPosition = $script:rowForCNCTeam;
                continue
            }
            'customercare-team'
            {
                $rangeForDeptToCheck = $script:rangeForCustomerCare;
                $deptInsertPosition = $script:rowForCustomerCare;
                continue
            }
            'data-exchange'
            {
                $rangeForDeptToCheck = $script:rangeForDataExchange;
                $deptInsertPosition = $script:rowForDataExchange;
                continue
            }
            'Despatch-Images'
            {
                $rangeForDeptToCheck = $script:rangeForDespatchImages;
                $deptInsertPosition = $script:rowForDespatchImages;
                continue
            }
            'engineering-team'
            {
                $rangeForDeptToCheck = $script:rangeForEngineering;
                $deptInsertPosition = $script:rowForEngineering;
                continue
            }
            'finance-team'
            {
                $rangeForDeptToCheck = $script:rangeForFinance;
                $deptInsertPosition = $script:rowForFinance;
                continue
            }
            'HealthandSafety'
            {
                $rangeForDeptToCheck = $script:rangeForHealth;
                $deptInsertPosition = $script:rowForHealth;
                continue
            }
            'human-resources'
            {
                $rangeForDeptToCheck = $script:rangeForHumanResources;
                $deptInsertPosition = $script:rowForHumanResources;
                continue
            }
            'Information-Centre'
            {
                $rangeForDeptToCheck = $script:rangeForInformation;
                $deptInsertPosition = $script:rowForInformation;
                continue
            }
            'it-team'
            {
                $rangeForDeptToCheck = $script:rangeForItTeam;
                $deptInsertPosition = $script:rowForItTeam;
                continue
            }
            'Logistics'
            {
                $rangeForDeptToCheck = $script:rangeForLogistics;
                $deptInsertPosition = $script:rowForLogistics;
                continue
            }
            'bridgetime'
            {
                $rangeForDeptToCheck = $script:rangeForBridgetime;
                $deptInsertPosition = $script:rowForBridgetime;
                continue
            }
            'MaterialsandWarehouse-Management'
            {
                $rangeForDeptToCheck = $script:rangeForMaterials;
                $deptInsertPosition = $script:rowForMaterials;
                continue
            }
            'newbuild'
            {
                $rangeForDeptToCheck = $script:rangeForNewBuild;
                $deptInsertPosition = $script:rowForNewBuild;
                continue
            }
            'newbuild-management'
            {
                $rangeForDeptToCheck = $script:rangeForNewBuildMan;
                $deptInsertPosition = $script:rowForNewBuildMan;
                continue
            }
            'NPD-project-planning'
            {
                $rangeForDeptToCheck = $script:rangeForNPDproject;
                $deptInsertPosition = $script:rowForNPDproject;
                continue
            }
            'orderprocessing-team'
            {
                $rangeForDeptToCheck = $script:rangeForOrderProc;
                $deptInsertPosition = $script:rowForOrderProc;
                continue
            }
            'planning'
            {
                $rangeForDeptToCheck = $script:rangeForPlanning;
                $deptInsertPosition = $script:rowForPlanning;
                continue
            }
            'production'
            {
                $rangeForDeptToCheck = $script:rangeForProduction;
                $deptInsertPosition = $script:rowForProduction;
                continue
            }
            'project-engineering'
            {
                $rangeForDeptToCheck = $script:rangeForProjectEngi;
                $deptInsertPosition = $script:rowForProjectEngi;
                continue
            }
            'quality-team'
            {
                $rangeForDeptToCheck = $script:rangeForQuality;
                $deptInsertPosition = $script:rowForQuality;
                continue
            }
            'quality-images'
            {
                $rangeForDeptToCheck = $script:rangeForQualityImages;
                $deptInsertPosition = $script:rowForQualityImages;
                continue
            }
            'reception'
            {
                $rangeForDeptToCheck = $script:rangeForReception
                $deptInsertPosition = $script:rowForReception;
                continue
            }
            'randd-team'
            {
                $rangeForDeptToCheck = $script:rangeForRandDteam;
                $deptInsertPosition = $script:rowForRandDteam;
                continue
            }
            'RMA'
            {
                $rangeForDeptToCheck = $script:rangeForRMA;
                $deptInsertPosition = $script:rowForRMA;
                continue
            }
            'sage-payroll'
            {
                $rangeForDeptToCheck = $script:rangeForSagePayroll;
                $deptInsertPosition = $script:rowForSagePayroll;
                continue
            }
            'sales-team'
            {
                $rangeForDeptToCheck = $script:rangeForSales;
                $deptInsertPosition = $script:rowForSales;
                continue
            }
            'Senior-Managers'
            {
                $rangeForDeptToCheck = $script:rangeForSeniorManagers;
                $deptInsertPosition = $script:rowForSeniorManagers;
                continue
            }
            'SeniorFinance-Team'
            {
                $rangeForDeptToCheck = $script:rangeForSeniorFinance;
                $deptInsertPosition = $script:rowForSeniorFinance;
                continue
            }
            'socialhousing-team'
            {
                $rangeForDeptToCheck = $script:rangeForSocialHousing;
                $deptInsertPosition = $script:rowForSocialHousing;
                continue
            }
            'team-leaders'
            {
                $rangeForDeptToCheck = $script:rangeForTeamLeaders;
                $deptInsertPosition = $script:rowForTeamLeaders;
                continue
            }
            'technical-team'
            {
                $rangeForDeptToCheck = $script:rangeForTechnical;
                $deptInsertPosition = $script:rowForTechnical;
                continue
            }
            'trade-team'
            {
                $rangeForDeptToCheck = $script:rangeForTrade;
                $deptInsertPosition = $script:rowForTrade;
                continue
            }
        }

        $userExists = $xlsheet.Range($rangeForDeptToCheck).Find($script:newUsername,[Type]::Missing,[Type]::Missing,1).Row
        if ($userExists -gt 0) {continue} # Check if the user exists already in that range, so duplicates aren't added

        $deptInsertPosition = $deptInsertPosition + 1 # To insert the row underneath the department header
        $whereToInsert = $xl.Range("C" + $deptInsertPosition).EntireRow # Insert the row
        [void] $whereToInsert.Insert(1)

        $xlsheet.Cells.Item($deptInsertPosition,3) = $script:newUsername # Add the names
        $xlsheet.Cells.Item($deptInsertPosition,4) = $script:newFirstName
        $xlsheet.Cells.Item($deptInsertPosition,5) = $script:newSurname
        
        $highlightRange = "A${deptInsertPosition}:E${deptInsertPosition}";
        $xlsheet.Range($highlightRange).Interior.Color = "&H33FF66" # Highlight the row green
    }

Set-Content "\\hgroup\data\GB-IGDoors\zPermissions\zPermissions to update.csv" -Value "Username,Department,First Name,Surname" # Reset the csv once complete

$xl.DisplayAlerts = $false; # Hides the pop-up dialog for saving, defaults to Yes
$xlbook.Save()
$xlbook.Close() # Important, otherwise the file is still open and becomes read-only
$xl.Quit()