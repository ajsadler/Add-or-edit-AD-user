# Clears current variables
Remove-Variable * -ErrorAction SilentlyContinue
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

# Defines the log file. Adds to existing file instead of over-writing
$Logfile = ".\Add or edit AD user log.log"
function WriteLog {
Param ([string]$LogString)
$Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
$LogMessage = "$Stamp $LogString"
Add-content $LogFile -value $LogMessage }

# Method for hiding the console in the background
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

# Picks out the current logged in user (ie. username-adm)
$logonName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$logonNameSplit = $logonName.Split("\")
$admUsername = $logonNameSplit[1]

# Defines the text culture necessary for 'ToTitleCase'
$textInfo = (Get-Culture).TextInfo

# GUI initialisation
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# GUI design
$labelfont = [System.Drawing.Font]::new("Tahoma", 11)
$errorfont = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.Fontstyle]::Bold)
$errorlabel = [System.Drawing.Font]::new("Arial", 9)

$form = New-Object System.Windows.Forms.Form
        $form.ClientSize = "500,600" # width 500, height 600
        $form.AutoSize = $true
        $form.StartPosition = 'CenterScreen'
        $form.Text = "Add/edit AD User"
        $form.BackColor = "#0091d3"

$igdoorslabel = New-Object System.Windows.Forms.Label
        $igdoorslabel.Location = New-Object System.Drawing.Point(400,10) # x 400, y 10
        $igdoorslabel.AutoSize = $true
        $igdoorslabel.Text = "IG Doors Ltd."
        $igdoorslabel.Font = "Impact, 13"
        $igdoorslabel.ForeColor = "white"

$edituserlabel = New-Object System.Windows.Forms.Label
        $edituserlabel.Location = New-Object System.Drawing.Point(30,48)
        $edituserlabel.AutoSize = $true
        $edituserlabel.Text = "Add/edit a user in Active Directory"
        $edituserlabel.Font = "Arial, 14"
        $edituserlabel.ForeColor = "white"

$namelabel = New-Object System.Windows.Forms.Label
        $namelabel.Location = New-Object System.Drawing.Point(30,92)
        $namelabel.AutoSize = $true
        $namelabel.Text = "Name"
        $namelabel.Font = $labelfont
        $namelabel.ForeColor = "white"

$nameerrorlabel = New-Object System.Windows.Forms.Label
        $nameerrorlabel.Location = New-Object System.Drawing.Point(150,120)
        $nameerrorlabel.AutoSize = $true
        $nameerrorlabel.Text = "Error: "
        $nameerrorlabel.Font = $errorfont
        $nameerrorlabel.ForeColor = "orange"
        $nameerrorlabel.Visible = $false

$nameinvalidlabel = New-Object System.Windows.Forms.Label
        $nameinvalidlabel.Location = New-Object System.Drawing.Point(190,120)
        $nameinvalidlabel.AutoSize = $true
        $nameinvalidlabel.Text = "Invalid name format"
        $nameinvalidlabel.Font = $errorlabel
        $nameinvalidlabel.ForeColor = "white"
        $nameinvalidlabel.Visible = $false

$namenouserlabel = New-Object System.Windows.Forms.Label
        $namenouserlabel.Location = New-Object System.Drawing.Point(190,120)
        $namenouserlabel.AutoSize = $true
        $namenouserlabel.Text = "No user found"
        $namenouserlabel.Font = $errorlabel
        $namenouserlabel.ForeColor = "white"
        $namenouserlabel.Visible = $false

$nameentryformatlabel = New-Object System.Windows.Forms.Label
        $nameentryformatlabel.Location = New-Object System.Drawing.Point(150,140)
        $nameentryformatlabel.AutoSize = $true
        $nameentryformatlabel.Text = "New user: enter their full name, eg. 'John Smith'
Edit user:  enter their username, eg. 'jsmith'"
        $nameentryformatlabel.Font = $labelfont
        $nameentryformatlabel.ForeColor = "white"

$nameentry = New-Object System.Windows.Forms.TextBox
        $nameentry.Location = New-Object System.Drawing.Point(150,90)
        $nameentry.AutoSize = $true
        $nameentry.Width = 200
        $nameentry.Text = ""
        $nameentry.Font = $labelfont
        $nameentry.TabIndex = 1

$newuser = New-Object System.Windows.Forms.Button
        $newuser.Location = New-Object System.Drawing.Point(360,90)
        $newuser.AutoSize = $true
        $newuser.Width = 50
        $newuser.Text = "New"
        $newuser.Font = $labelfont
        $newuser.ForeColor = "white"
        $newuser.BackColor = "green"
        $newuser.TabIndex = 2

$edituser = New-Object System.Windows.Forms.Button
        $edituser.Location = New-Object System.Drawing.Point(415,90)
        $edituser.AutoSize = $true
        $edituser.Width = 50
        $edituser.Text = "Edit"
        $edituser.Font = $labelfont
        $edituser.ForeColor = "black"
        $edituser.BackColor = "yellow"
        $edituser.TabIndex = 3

$resetall = New-Object System.Windows.Forms.Button
        $resetall.Location = New-Object System.Drawing.Point(360,90)
        $resetall.AutoSize = $true
        $resetall.Width = 105
        $resetall.Text = "Reset All"
        $resetall.Font = $labelfont
        $resetall.ForeColor = "black"
        $resetall.BackColor = "orange"
        $resetall.Visible = $false
        $resetall.TabIndex = 15

$datelabel = New-Object System.Windows.Forms.Label
        $datelabel.Location = New-Object System.Drawing.Point(30,142)
        $datelabel.AutoSize = $true
        $datelabel.Text = "Start date"
        $datelabel.Font = $labelfont
        $datelabel.ForeColor = "white"
        $datelabel.Visible = $false

$dateerrorlabel = New-Object System.Windows.Forms.Label
        $dateerrorlabel.Location = New-Object System.Drawing.Point(150,170)
        $dateerrorlabel.AutoSize = $true
        $dateerrorlabel.Text = "Error: "
        $dateerrorlabel.Font = $errorfont
        $dateerrorlabel.ForeColor = "orange"
        $dateerrorlabel.Visible = $false

$dateinvalidlabel = New-Object System.Windows.Forms.Label
        $dateinvalidlabel.Location = New-Object System.Drawing.Point(190,170)
        $dateinvalidlabel.AutoSize = $true
        $dateinvalidlabel.Text = "Invalid date format (DD/MM/YYYY)"
        $dateinvalidlabel.Font = $errorlabel
        $dateinvalidlabel.ForeColor = "white"
        $dateinvalidlabel.Visible = $false

$today = Get-Date -Format "dd/MM/yyyy"
$dateentry = New-Object System.Windows.Forms.DateTimePicker
        $dateentry.Location = New-Object System.Drawing.Point(150,140)
        $dateentry.AutoSize = $true
        $dateentry.Width = 200
        $dateentry.Text = "$today"
        $dateentry.Font = $labelfont
        $dateentry.TabIndex = 4
        $dateentry.Visible = $false

$titlelabel = New-Object System.Windows.Forms.Label
        $titlelabel.Location = New-Object System.Drawing.Point(30,192)
        $titlelabel.AutoSize = $true
        $titlelabel.Text = "Job Title"
        $titlelabel.Font = $labelfont
        $titlelabel.ForeColor = "white"
        $titlelabel.Visible = $false

$titleentry = New-Object System.Windows.Forms.TextBox
        $titleentry.Location = New-Object System.Drawing.Point(150,190)
        $titleentry.AutoSize = $true
        $titleentry.Width = 200
        $titleentry.Text = ""
        $titleentry.Font = $labelfont
        $titleentry.TabIndex = 5
        $titleentry.Visible = $false

$deptlabel = New-Object System.Windows.Forms.Label
        $deptlabel.Location = New-Object System.Drawing.Point(30,242)
        $deptlabel.AutoSize = $true
        $deptlabel.Text = "Department(s)"
        $deptlabel.Font = $labelfont
        $deptlabel.ForeColor = "white"
        $deptlabel.Visible = $false

$depterrorlabel = New-Object System.Windows.Forms.Label
        $depterrorlabel.Location = New-Object System.Drawing.Point(150,270)
        $depterrorlabel.AutoSize = $true
        $depterrorlabel.Text = "Error: "
        $depterrorlabel.Font = $errorfont
        $depterrorlabel.ForeColor = "orange"
        $depterrorlabel.Visible = $false

$deptnoselectedlabel = New-Object System.Windows.Forms.Label
        $deptnoselectedlabel.Location = New-Object System.Drawing.Point(190,270)
        $deptnoselectedlabel.AutoSize = $true
        $deptnoselectedlabel.Text = "No departments selected"
        $deptnoselectedlabel.Font = $errorlabel
        $deptnoselectedlabel.ForeColor = "white"
        $deptnoselectedlabel.Visible = $false

$deptinvalidlabel = New-Object System.Windows.Forms.Label
        $deptinvalidlabel.Location = New-Object System.Drawing.Point(190,270)
        $deptinvalidlabel.AutoSize = $true
        $deptinvalidlabel.Text = "No primary department set"
        $deptinvalidlabel.Font = $errorlabel
        $deptinvalidlabel.ForeColor = "white"
        $deptinvalidlabel.Visible = $false

$deptentry = New-Object System.Windows.Forms.ComboBox
        $deptentry.Location = New-Object System.Drawing.Point(150,240)
        $deptentry.Autosize = $true
        $deptentry.Width = 200
        $deptentry.Text = ""
        $deptlist = @(' ','3Tec','Accounts Department','After Sales Department','Bridgetime Scans','CNC Team','Customer Care Department','Data Exchange','Despatch Images','Engineering Department','Finance Department','Front Office Department','General Management','Health & Safety Department', 'Human Resources Department', 'Information Centre','I.T. Systems', 'Logistics', 'Logistics (Bridgetime)', 'Materials & Warehouse Management', 'New Build Division', 'New Build Management','NPD Project Planning','Operations Department', 'Order Processing Department', 'Payroll Department', 'Planning Department', 'Production Department', 'Project Engineering Department', 'Purchasing Department', 'Quality & HSE Department', 'Quality Images','Reception','Research & Development', 'RMA', 'Sales Department', 'Senior Managers','Senior Finance Team','Social Housing Department', 'Team Leaders','Technical Department', 'Trade Division')
        @($deptlist) | ForEach-Object {[void] $deptentry.Items.Add($_)}
        $deptentry.SelectedIndex = 0
        $deptentry.Font = $labelfont
        $deptentry.TabIndex = 6
        $deptentry.Visible = $false

$accessgroupbox = New-Object System.Windows.Forms.GroupBox
        $accessgroupbox.Location = New-Object System.Drawing.Point(25,290)
        $accessgroupbox.Size = "325,57"
        $accessgroupbox.Text = "Access group:"
        $accessgroupbox.Font = $labelfont
        $accessgroupbox.ForeColor = "white"
        $accessgroupbox.TabIndex = 10
        $accessgroupbox.Visible = $false

$trustedlocation = New-Object System.Windows.Forms.RadioButton
        $trustedlocation.Location = New-Object System.Drawing.Point(130,10)
        $trustedlocation.Autosize = $true
        $trustedlocation.Checked = $true
        $trustedlocation.Text = "Allow Trusted Locations"
        $trustedlocation.Font = $labelfont
        $trustedlocation.ForeColor = "white"
        $trustedlocation.Visible = $false

$requireusermfa = New-Object System.Windows.Forms.RadioButton
        $requireusermfa.Location = New-Object System.Drawing.Point(130,30)
        $requireusermfa.Autosize = $true
        $requireusermfa.Checked = $false
        $requireusermfa.Text = "Require User MFA"
        $requireusermfa.Font = $labelfont
        $requireusermfa.ForeColor = "white"
        $requireusermfa.Visible = $false

$adddept = New-Object System.Windows.Forms.Button
        $adddept.Location = New-Object System.Drawing.Point(360,240)
        $adddept.AutoSize = $true
        $adddept.Width = 50
        $adddept.Text = "Add"
        $adddept.Font = $labelfont
        $adddept.ForeColor = "#0091d3"
        $adddept.BackColor = "white"
        $adddept.TabIndex = 7
        $adddept.Visible = $false

$deldept = New-Object System.Windows.Forms.Button
        $deldept.Location = New-Object System.Drawing.Point(415,240)
        $deldept.AutoSize = $true
        $deldept.Width = 50
        $deldept.Text = "Del"
        $deldept.Font = $labelfont
        $deldept.ForeColor = "#0091d3"
        $deldept.BackColor = "white"
        $deldept.TabIndex = 8
        $deldept.Visible = $false

$confirmnew = New-Object System.Windows.Forms.Button
        $confirmnew.Location = New-Object System.Drawing.Point(150,360)
        $confirmnew.AutoSize = $true
        $confirmnew.Width = 90
        $confirmnew.Text = "Check details"
        $confirmnew.Font = $labelfont
        $confirmnew.ForeColor = "black"
        $confirmnew.BackColor = "yellow"
        $confirmnew.Visible = $false
        $confirmnew.TabIndex = 11

$confirmedit = New-Object System.Windows.Forms.Button
        $confirmedit.Location = New-Object System.Drawing.Point(145,360)
        $confirmedit.AutoSize = $true
        $confirmedit.Width = 90
        $confirmedit.Text = "Update details"
        $confirmedit.Font = $labelfont
        $confirmedit.ForeColor = "black"
        $confirmedit.BackColor = "yellow"
        $confirmedit.Visible = $false
        $confirmedit.TabIndex = 12

$gobuttonnew = New-Object System.Windows.Forms.Button
        $gobuttonnew.Location = New-Object System.Drawing.Point(260,360)
        $gobuttonnew.AutoSize = $true
        $gobuttonnew.Width = 90
        $gobuttonnew.Text = "Run script"
        $gobuttonnew.Font = $labelfont
        $gobuttonnew.ForeColor = "white"
        $gobuttonnew.BackColor = "green"
        $gobuttonnew.Visible = $false
        $gobuttonnew.TabIndex = 13

$gobuttonedit = New-Object System.Windows.Forms.Button
        $gobuttonedit.Location = New-Object System.Drawing.Point(260,360)
        $gobuttonedit.AutoSize = $true
        $gobuttonedit.Width = 90
        $gobuttonedit.Text = "Run script"
        $gobuttonedit.Font = $labelfont
        $gobuttonedit.ForeColor = "white"
        $gobuttonedit.BackColor = "green"
        $gobuttonedit.Visible = $false
        $gobuttonedit.TabIndex = 14

$resetdeptnew = New-Object System.Windows.Forms.Button
        $resetdeptnew.Location = New-Object System.Drawing.Point(360,275)
        $resetdeptnew.Height = 25
        $resetdeptnew.Width = 105
        $resetdeptnew.Text = "Reset Depts"
        $resetdeptnew.Font = $labelfont
        $resetdeptnew.ForeColor = "black"
        $resetdeptnew.BackColor = "orange"
        $resetdeptnew.Visible = $false
        $resetdeptnew.TabIndex = 9

$resetdeptedit = New-Object System.Windows.Forms.Button
        $resetdeptedit.Location = New-Object System.Drawing.Point(360,275)
        $resetdeptedit.Height = 25
        $resetdeptedit.Width = 105
        $resetdeptedit.Text = "Reset Depts"
        $resetdeptedit.Font = $labelfont
        $resetdeptedit.ForeColor = "black"
        $resetdeptedit.BackColor = "orange"
        $resetdeptedit.Visible = $false
        $resetdeptedit.TabIndex = 9

$exitbutton = New-Object System.Windows.Forms.Button
        $exitbutton.Location = New-Object System.Drawing.Point(440,40)
        $exitbutton.AutoSize = $true
        $exitbutton.Width = 50
        $exitbutton.Text = "Exit"
        $exitbutton.Font = $labelfont
        $exitbutton.ForeColor = "white"
        $exitbutton.BackColor = "red"
        $exitbutton.TabIndex = 16

$categorylabel = New-Object System.Windows.Forms.Label
        $categorylabel.Location = New-Object System.Drawing.Point(30,380)
        $categorylabel.AutoSize = $true
        $categorylabel.Text = ""
        $categorylabel.Font = $labelfont
        $categorylabel.ForeColor = "white"
        $categorylabel.Visible = $false

$detailslabel = New-Object System.Windows.Forms.Label
        $detailslabel.Location = New-Object System.Drawing.Point(200,380)
        $detailslabel.AutoSize = $true
        $detailslabel.Text = ""
        $detailslabel.Font = $labelfont
        $detailslabel.ForeColor = "white"
        $detailslabel.Visible = $false

$progressbar = New-Object System.Windows.Forms.ProgressBar
        $progressbar.Location = New-Object System.Drawing.Point(10,13)
        $progressbar.Size = New-Object System.Drawing.Size(380,15)
        $progressbar.Value = 0
        $progressbar.Style = "Continuous"
        $progressbar.ForeColor = "green"
        $progressbar.BackColor = "white"

# When adding a new component, make sure to add it to the $form.Controls.AddRange(@())
$accessgroupbox.Controls.AddRange(@($trustedlocation, $requireusermfa))
$form.Controls.AddRange(@($igdoorslabel, $edituserlabel, $namelabel, $nameerrorlabel, $nameinvalidlabel, $namenouserlabel, $nameentryformatlabel, $nameentry, $newuser, $edituser, $resetall, $datelabel, $dateerrorlabel, $dateinvalidlabel, $dateentry, $titlelabel, $titleentry, $deptlabel, $depterrorlabel, $deptnoselectedlabel, $deptinvalidlabel, $deptentry, $accessgroupbox, $adddept, $deldept, $confirmnew, $confirmedit, $gobuttonnew, $gobuttonedit, $exitbutton, $resetdeptnew, $resetdeptedit, $categorylabel, $detailslabel, $progressbar))

function StartVisibility # Reveals all the necessary components when starting with 'New' or 'Edit'
{
    $newuser.Visible = $false
    $edituser.Visible = $false
    $nameentryformatlabel.Visible = $false
    $resetall.Visible = $true
    $titlelabel.Visible = $true
    $titleentry.Visible = $true
    $deptlabel.Visible = $true
    $deptentry.Visible = $true
    $accessgroupbox.Visible = $true
    $trustedlocation.Visible = $true
    $requireusermfa.Visible = $true
    $adddept.Visible = $true
    $deldept.Visible = $true
    $categorylabel.Visible = $true
    $detailslabel.Visible = $true
}

function StartDeptArrays # Resets all the department arrays. Needed when possibly clicking multiple add/delete/confirm buttons
{
    $script:newDepartment = @()
    $script:newDepartmentFull = @()
    $script:newDepartmentAD = @()
    $script:delDepartment = @()
    $script:delDepartmentAD = @()
}

function AddDepartment # Adds the selection to a list of departments to be added (and deletes it from the list to be removed)
{
    DeptNoSelectionCheck
    $script:newDepartment += $deptentry.Text
    $script:delDepartment = $script:delDepartment -ne $deptentry.Text
    if ($confirmedit.Visible -eq $false) {$confirmnew.Visible = $true}
}

function DelDepartment # Adds the selection to a list of departments to be removed (and deletes it from the list to be added)
{
    DeptNoSelectionCheck
    $script:newDepartment = $script:newDepartment -ne $deptentry.Text
    $script:newDepartmentAD = @()
    $script:delDepartment += $deptentry.Text
    $script:delDepartmentAD = @()
    $gobuttonnew.Visible = $false
}

function EachDeptInNew # For each dept in the 'new' array, transform it to the AD group name
{
    $script:newDepartmentAD = @()
    foreach ($department in $script:newDepartment)
    {
        switch -wildcard ($department)
        {
            '3t*'           {$script:newDepartmentFull += "3Tec";                                  $script:newDepartmentAD += "3tec"}
            'acc*'          {$script:newDepartmentFull += "Accounts Department";                   } # $script:newDepartmentAD += ""}
            'aft*'          {$script:newDepartmentFull += "After Sales Department";                $script:newDepartmentAD += "aftersales-team"}
            'bri*'          {                                                                      $script:newDepartmentAD += "Bridgetime-Scans"}
            'cnc*'          {$script:newDepartmentFull += "CNC Team";                              $script:newDepartmentAD += "CNC-Team"}
            'cus*'          {$script:newDepartmentFull += "Customer Care Department";              $script:newDepartmentAD += "customercare-team"}
            'dat*'          {                                                                      $script:newDepartmentAD += "data-exchange"}
            'des*'          {                                                                      $script:newDepartmentAD += "Despatch-Images"}
            'eng*'          {$script:newDepartmentFull += "Engineering Department";                $script:newDepartmentAD += "engineering-team"}
            'fin*'          {$script:newDepartmentFull += "Finance Department";                    $script:newDepartmentAD += "finance-team"}
            'hea*'          {$script:newDepartmentFull += "Health && Safety Department";           $script:newDepartmentAD += "HealthandSafety"}
            'hum*'          {$script:newDepartmentFull += "Human Resources Department";            $script:newDepartmentAD += "human-resources"}
            'hr'            {$script:newDepartmentFull += "Human Resources Department";            $script:newDepartmentAD += "human-resources"}
            'inf*'          {                                                                      $script:newDepartmentAD += "Information-Centre"}
            'i.t*'          {$script:newDepartmentFull += "I.T. Systems";                          $script:newDepartmentAD += "it-team"}
            'it*'           {$script:newDepartmentFull += "I.T. Systems";                          $script:newDepartmentAD += "it-team"}
            'log*'          {$script:newDepartmentFull += "Logistics";                             $script:newDepartmentAD += "Logistics"}
            '*brid*'        {$script:newDepartmentFull += "Logistics (Bridgetime)";                $script:newDepartmentAD += "bridgetime"}
            'mat*'          {$script:newDepartmentFull += "Materials && Warehouse Management";     $script:newDepartmentAD += "MaterialsandWarehouse-Management"}
            'new build d*'  {$script:newDepartmentFull += "New Build Division";                    $script:newDepartmentAD += "newbuild"}
            'new build m*'  {                                                                      $script:newDepartmentAD += "newbuild-management"}
            'npd*'          {                                                                      $script:newDepartmentAD += "NPD-Project-Planning"}
            'ope*'          {$script:newDepartmentFull += "Operations Department";                 } # $script:newDepartmentAD += ""}
            'ord*'          {$script:newDepartmentFull += "Order Processing Department";           $script:newDepartmentAD += "orderprocessing-team"}
            'pla*'          {$script:newDepartmentFull += "Planning Department";                   $script:newDepartmentAD += "planning"}
            'prod*'         {$script:newDepartmentFull += "Production Department";                 $script:newDepartmentAD += "production"}
            'proj*'         {$script:newDepartmentFull += "Project Engineering Department";        $script:newDepartmentAD += "project-engineering"}
            'pur*'          {$script:newDepartmentFull += "Purchasing Department";                 } # $script:newDepartmentAD += ""}
            'quality &*'    {$script:newDepartmentFull += "Quality && HSE Department";             $script:newDepartmentAD += "quality-team"}
            'quality i*'    {                                                                      $script:newDepartmentAD += "quality-images"}
            'rec*'          {                                                                      $script:newDepartmentAD += "reception"}
            'res*'          {$script:newDepartmentFull += "Research && Development";               $script:newDepartmentAD += "randd-team"}
            'rma*'          {$script:newDepartmentFull += "RMA";                                   $script:newDepartmentAD += "RMA"}
            'sag*'          {$script:newDepartmentFull += "Payroll Department";                    $script:newDepartmentAD += "sage-payroll"}
            'pay*'          {$script:newDepartmentFull += "Payroll Department";                    $script:newDepartmentAD += "sage-payroll"}
            'sal*'          {$script:newDepartmentFull += "Sales Department";                      $script:newDepartmentAD += "sales-team"}
            'senior m*'     {                                                                      $script:newDepartmentAD += "Senior-Managers"}
            'senior f*'     {                                                                      $script:newDepartmentAD += "SeniorFinance-Team"}
            'soc*'          {$script:newDepartmentFull += "Social Housing Department";             $script:newDepartmentAD += "socialhousing-team"}
            'tea*'          {                                                                      $script:newDepartmentAD += "team-leaders"}
            'tec*'          {$script:newDepartmentFull += "Technical Department";                  $script:newDepartmentAD += "technical-team"}
            'tra*'          {$script:newDepartmentFull += "Trade Division";                        $script:newDepartmentAD += "trade-team"}
            Default         {""}
        }
    }
    $script:newDepartmentPrimary = $deptentry.Text
    try {$script:newDepartmentPrimary = $script:newDepartmentPrimary.Replace("&", "&&")} catch {}
    $script:newDepartmentAD = $script:newDepartmentAD | Select-Object -Unique
    $script:newDepartmentADList = $script:newDepartmentAD -join '
    '
    $script:newDepartmentADLogList = $script:newDepartmentAD -join ', '
}

function EachDeptInDel # For each dept in the 'del' array, transform it to the AD group name
{
    $script:delDepartmentAD = @()
    foreach ($department in $script:delDepartment)
    {
        switch -wildcard ($department)
        {
            '3t*'           {$script:delDepartmentAD += "3tec"}
            'acc*'          {} # $script:delDepartmentAD += ""}
            'aft*'          {$script:delDepartmentAD += "aftersales-team"}
            'bri*'          {$script:delDepartmentAD += "Bridgetime-Scans"}
            'cnc*'          {$script:delDepartmentAD += "CNC-Team"}
            'cus*'          {$script:delDepartmentAD += "customercare-team"}
            'dat*'          {$script:delDepartmentAD += "data-exchange"}
            'des*'          {$script:delDepartmentAD += "Despatch-Images"}
            'eng*'          {$script:delDepartmentAD += "engineering-team"}
            'fin*'          {$script:delDepartmentAD += "finance-team"}
            'hea*'          {$script:delDepartmentAD += "HealthandSafety"}
            'hum*'          {$script:delDepartmentAD += "human-resources"}
            'hr'            {$script:delDepartmentAD += "human-resources"}
            'inf*'          {$script:delDepartmentAD += "Information-Centre"}
            'i.t*'          {$script:delDepartmentAD += "it-team"}
            'it*'           {$script:delDepartmentAD += "it-team"}
            'log*'          {$script:delDepartmentAD += "Logistics"}
            '*brid*'        {$script:delDepartmentAD += "bridgetime"}
            'mat*'          {$script:delDepartmentAD += "MaterialsandWarehouse-Management"}
            'new build d*'  {$script:delDepartmentAD += "newbuild"}
            'new build m*'  {$script:delDepartmentAD += "newbuild-management"}
            'npd*'          {$script:delDepartmentAD += "NPD-Project-Planning"}
            'ope*'          {} # $script:delDepartmentAD += ""}
            'ord*'          {$script:delDepartmentAD += "orderprocessing-team"}
            'pla*'          {$script:delDepartmentAD += "planning"}
            'prod*'         {$script:delDepartmentAD += "production"}
            'proj*'         {$script:delDepartmentAD += "project-engineering"}
            'pur*'          {} # $script:delDepartmentAD += ""}
            'quality &*'    {$script:delDepartmentAD += "quality-team"}
            'quality i*'    {$script:delDepartmentAD += "quality-images"}
            'rec*'          {$script:delDepartmentAD += "reception"}
            'res*'          {$script:delDepartmentAD += "randd-team"}
            'rma*'          {$script:delDepartmentAD += "RMA"}
            'sag*'          {$script:delDepartmentAD += "sage-payroll"}
            'pay*'          {$script:delDepartmentAD += "sage-payroll"}
            'sal*'          {$script:delDepartmentAD += "sales-team"}
            'senior m*'     {$script:delDepartmentAD += "Senior-Managers"}
            'senior f*'     {$script:delDepartmentAD += "SeniorFinance-Team"}
            'soc*'          {$script:delDepartmentAD += "socialhousing-team"}
            'tea*'          {$script:delDepartmentAD += "team-leaders"}
            'tec*'          {$script:delDepartmentAD += "technical-team"}
            'tra*'          {$script:delDepartmentAD += "trade-team"}
            Default         {""}
        }
    }

    $script:delDepartmentAD = $script:delDepartmentAD | Select-Object -Unique
    $script:delDepartmentADList = $script:delDepartmentAD -join '
    '
    $script:delDepartmentADLogList = $script:delDepartmentAD -join ', '
}

function ReadAccessGroup # Check the Access Group selected
{
    $script:accessGroup = ""
    if ($trustedlocation.Checked) {$script:accessGroup = "allowTrustedLocations"}
        elseif ($requireusermfa.Checked) {$script:accessGroup = "requireUserMFA"}
        else {return}
}

function DeptNoSelectionCheck # Check for a valid department to be selected from the drop-down box
{
    if ($deptentry.Text -eq "" -or $deptentry.Text -eq " ") {$depterrorlabel.Visible = $true; $deptnoselectedlabel.Visible = $true; return}
        else {$depterrorlabel.Visible = $false; $deptnoselectedlabel.Visible = $false; $deptinvalidlabel.Visible = $false}
}

function DeptErrorCheck # Check for a primary department and at least one AD group
{
    if ($null -eq $script:newDepartmentPrimary) {$depterrorlabel.Visible = $true; $deptinvalidlabel.Visible = $true; return}
        else {$depterrorlabel.Visible = $false; $deptinvalidlabel.Visible = $false}
    if ($null -eq $script:newDepartmentAD) {$depterrorlabel.Visible = $true; $deptnoselectedlabel.Visible = $true; return}
}

function NewLabelDetails # Change the label details format, including the start date
{
    $categorylabel.Text = "
New user:
Display name:
Login name:
Email address:
Start date:
Job Title:
Department:
Active Directory groups:"

    $detailslabel.Text = "
    $script:newFirstName $script:newSurname
    $script:newDisplayName
    $script:newUsername
    $script:newEmail
    $script:newStartDateDDMMYYYY
    $script:newJobTitle
    $script:newDepartmentPrimary
    $script:accessGroup
    $script:newDepartmentADList"
}

function EditLabelDetails # Change the label details format, without the start date
{
    $categorylabel.Text = "
New user:
Display name:
Login name:
Email address:

Job Title:
Department:
Active Directory groups:"

    $detailslabel.Text = "
    $script:newFirstName $script:newSurname
    $script:newDisplayName
    $script:newUsername
    $script:newEmail

    $script:newJobTitle
    $script:newDepartmentPrimary
    $script:accessGroup
    $script:newDepartmentADList"
}

function ResetAll # Reset all the components back to as if the exe was just opened
{
    ResetDeptNew
    ResetDeptEdit
    StartDeptArrays
    $newuser.Visible = $true
    $edituser.Visible = $true
    $nameentryformatlabel.Visible = $true
    $nameerrorlabel.Visible = $false
    $nameinvalidlabel.Visible = $false
    $namenouserlabel.Visible = $false
    $confirmedit.Visible = $false
    $resetall.Visible = $false
    $datelabel.Visible = $false
    $dateentry.Visible = $false
    $titlelabel.Visible = $false
    $titleentry.Visible = $false
    $deptlabel.Visible = $false
    $deptentry.Visible = $false
    $depterrorlabel.Visible = $false
    $deptnoselectedlabel.Visible = $false
    $deptinvalidlabel.Visible = $false
    $trustedlocation.Visible = $false
    $requireusermfa.Visible = $false
    $accessgroupbox.Visible = $false
    $trustedlocation.Checked = $true
    $adddept.Visible = $false
    $deldept.Visible = $false
    $resetdeptnew.Visible = $false
    $resetdeptnew.Visible = $false
    $resetdeptedit.Visible = $false
    $categorylabel.Visible = $false
    $detailslabel.Visible = $false
}

function ResetDeptNew # Reset all the departments selected when adding a new user
{
    StartDeptArrays
    NewLabelDetails
    $script:newDepartmentPrimary = ""
    $script:newDepartmentADList = ""
    $deptentry.SelectedIndex = 0
    $confirmnew.Visible = $false
    $confirmedit.Visible = $false
    $gobuttonnew.Visible = $false
}

function ResetDeptEdit # Reset all the department changes when editing an existing user (ie. read the current details from scratch again)
{
    EditUser
}

function ExitButton # Close the exe
{
    $form.Close()
}

function NewUser # Start a new user ; validate the name entry ; check for duplicates ; transform the name ; reveal the next components
{
    $script:selectNewOrEdit = "New"
    StartDeptArrays
    $titleentry.Text = ""
    $script:newJobTitle = ""
    $deptentry.Text = ""
    $script:newName = $nameentry.Text
    $script:newFullName = $script:newName.split(" ")
    
    if ($newFullName.Length -ne 2) {$nameerrorlabel.Visible = $true; $nameinvalidlabel.Visible = $true; return}
        elseif ($null -eq $newFullName) {$nameerrorlabel.Visible = $true; $nameinvalidlabel.Visible = $true; return}
        elseif ($newName -match '\d') {$nameerrorlabel.Visible = $true; $nameinvalidlabel.Visible = $true; return}
        elseif ($newName -match '[,.;:@!?]') {$nameerrorlabel.Visible = $true; $nameinvalidlabel.Visible = $true; return}
        else {$nameerrorlabel.Visible = $false; $nameinvalidlabel.Visible = $false; $namenouserlabel.Visible = $false}

    $script:newFirstName = $textInfo.ToTitleCase($script:newFullName[0].ToLower())
    $script:newSurname = $textInfo.ToTitleCase($script:newFullName[1].ToLower())
    $script:newFirstInitial = $script:newFirstName.Substring(0,1).ToLower()
    $script:newSurnameInitial = $script:newSurname.Substring(0,1).ToLower()
    $script:newFirstNameLower = $script:newFirstName.ToLower()
    $script:newSurnameLower = $script:newSurname.ToLower()
    $script:newSurnameTruncate = $script:newSurnameLower.Replace("-", "")
    $script:newSurnameTruncateLength = $script:newSurnameTruncate.Length
    if ($script:newSurnameTruncateLength -gt 9) {$script:newSurnameTruncate = $script:newSurnameTruncate.Substring(0,9)}
    $script:newUsername = $script:newFirstInitial + $script:newSurnameTruncate
    $script:newEmail = "${script:newFirstNameLower}.${script:newSurnameLower}@igdoors.co.uk"
    $script:newDisplayName = "${script:newSurname}, ${script:newFirstName}"
    $script:newEmailLong = "${script:newFirstNameLower}.${script:newSurnameLower}.igdoors.co.uk@hgrouponline.mail.onmicrosoft.com"

    $error.Clear()
    $i = 1
    do
    {   try {Get-ADUser $script:newUsername}
        catch {$i--
            "Duplicate user(s) found: $i"}
        if ($error) {continue}
        else {
            $i++
            if ($script:newSurnameTruncateLength -gt 8) {$script:newSurnameTruncate = $script:newSurnameTruncate.Substring(0,8)}
            $script:newUsername = $script:newFirstInitial + $script:newSurnameTruncate + $i
            $script:newEmail = "${script:newFirstNameLower}.${script:newSurnameLower}${i}@igdoors.co.uk"
            $script:newDisplayName = "${script:newSurname}, ${script:newFirstName}${i}"
            $script:newEmailLong = "${script:newFirstNameLower}.${script:newSurnameLower}${i}.igdoors.co.uk@hgrouponline.mail.onmicrosoft.com"
            }
    } while (!$error)

    StartVisibility
    $resetdeptnew.Visible = $true
    $datelabel.Visible = $true
    $dateentry.Visible = $true
    NewLabelDetails
}

function ConfirmDetailsNew # Validate the date entry ; transform the date ; run EachDeptInNew ; run ReadAccessGroup ; update the label ; reveal the 'run' button
{
    $delimiters = "/",".","-",",",":",";","\","=","_","*"," "
    $newStartDate = $dateentry.Text
    $newStartDateSplit = $newStartDate -Split {$delimiters -contains $_}

    switch ($newStartDateSplit[1])
    {
        'January'   {$newStartDateSplit[1] = '01'}
        'February'  {$newStartDateSplit[1] = '02'}
        'March'     {$newStartDateSplit[1] = '03'}
        'April'     {$newStartDateSplit[1] = '04'}
        "May"       {$newStartDateSplit[1] = "05"}
        'June'      {$newStartDateSplit[1] = '06'}
        'July'      {$newStartDateSplit[1] = '07'}
        'August'    {$newStartDateSplit[1] = '08'}
        'September' {$newStartDateSplit[1] = '09'}
        'October'   {$newStartDateSplit[1] = '10'}
        'November'  {$newStartDateSplit[1] = '11'}
        'December'  {$newStartDateSplit[1] = '12'}
    }

    if ($newStartDateSplit[0].Length -ne 2) {$dateerrorlabel.Visible = $true; $dateinvalidlabel.Visible = $true; return}
        elseif ($newStartDateSplit[1].Length -ne 2) {$dateerrorlabel.Visible = $true; $dateinvalidlabel.Visible = $true; return}
        elseif ($newStartDateSplit[2].Length -ne 4) {$dateerrorlabel.Visible = $true; $dateinvalidlabel.Visible = $true; return}
        else {$dateerrorlabel.Visible = $false; $dateinvalidlabel.Visible = $false}

    $script:newStartDateDD = $newStartDateSplit[0]
    $script:newStartDateMM = $newStartDateSplit[1]
    $script:newStartDateYYYY = $newStartDateSplit[2]
    $script:newStartDateDDMMYYYY = $newStartDateSplit[0] + "/" + $newStartDateSplit[1] + "/" + $newStartDateSplit[2]

    $script:newPasswordString = "${script:newFirstInitial}${script:newSurnameInitial}${script:newStartDateDD}${script:newStartDateMM}"
    $script:newPassword = "IGD${script:newPasswordString}#"

    $script:newJobTitle = $titleentry.Text
    EachDeptInNew
    ReadAccessGroup
    NewLabelDetails
    DeptErrorCheck
    $GoButtonNew.Visible = $true
}

function GoButtonNew # Write to the log ; create new user ; set the user attributes ; add proxy addresses ; add to Access Group ; add to zP groups ; add to zP excel list
{

    WriteLog "`n

    NEW USER
    "
    WriteLog "The script started by user $admUsername."
    WriteLog "New user: $script:newFirstName $script:newSurname"
    WriteLog "Displayname: $script:newDisplayName"
    WriteLog "Login name: $script:newUsername"
    WriteLog "Email address: $script:newEmail"
    WriteLog "Start date: $script:newStartDateDDMMYYYY"
    WriteLog "Job Title: $script:newJobTitle"
    WriteLog "Department: $script:newDepartmentPrimary"
    WriteLog "Active Directory groups: $script:newDepartmentADLogList"
    WriteLog "Access group: $script:accessGroup"
    
    $i=0
    do {$i++; $progressbar.Value = $i; Start-Sleep -Seconds 0.5} while ($i -lt 10)

    New-ADUser -Name "$script:newUsername" -path "OU=internal Users,OU=user,OU=blw,OU=gb,OU=hgroup-production,DC=hgroup,DC=intra" -AccountPassword $(ConvertTo-SecureString -AsPlainText $script:newPassword -Force) -Enabled $true

    do {$i++; $progressbar.Value = $i; Start-Sleep -Seconds 1.0} while ($i -lt 30) # Longer delay to ensure user is created before setting attributes and groups

    Set-ADUser -Identity "$script:newUsername" -Replace @{
        'Company' = "IG Doors Ltd.";
        'StreetAddress' = "Lon Gellideg, Oakdale Business Park";
        'l' = "Blackwood";
        'st' = "South Wales";
        'PostalCode' = "NP12 4AE";
        'c' = "GB";
        'co' = "United Kingdom";
        'countryCode' = "826";
        'GivenName' = "$script:newFirstName";
        'sn' = "$script:newSurname";
        'DisplayName' = "$script:newDisplayName";
        'mailNickname' = "$script:newUsername";
        'mail' = "$script:newEmail";
        'UserPrincipalName' = "$script:newEmail";
        'targetAddress' = "SMTP:$script:newEmailLong";
        'Department' = "$script:newDepartmentPrimary";
        'Title' = "$script:newJobTitle";
        'extensionAttribute1' = "Included";
        'extensionAttribute15' = "AADSync";
    }
    
    do {$i++; $progressbar.Value = $i; Start-Sleep -Seconds 0.5} while ($i -lt 40)

    Get-ADUser "$script:newUsername" | Rename-ADObject -NewName "$script:newDisplayName"
    
    do {$i++; $progressbar.Value = $i; Start-Sleep -Seconds 0.5} while ($i -lt 50)

    $User = Get-ADUser $script:newUsername -Properties proxyAddresses
    $User.proxyAddresses.Add("SMTP:$script:newEmail")
    $User.proxyAddresses.Add("smtp:$script:newEmailLong")
    Set-ADUser -instance $User

    do {$i++; $progressbar.Value = $i; Start-Sleep -Seconds 0.5} while ($i -lt 60)

    Add-ADGroupMember -Identity zzx-gbblw-m365-ConditionalAccess-$script:accessGroup -Members $script:newUsername # Access group - office only or remote enabled
    Add-ADGroupMember -Identity zzo-gbblw-user-all -Members $script:newUsername # User group - grants default zpermissions file access
    Add-ADGroupMember -Identity zzl-gbblw-M365E3-default -Members $script:newUsername # License group - default license for Microsoft365
    Add-ADGroupMember -Identity zzl-gbblw-M365E3-ExchangeOnlinePlan2 -Members $script:newUsername # License group - Exchange Online for Microsoft365
    
    do {$i++; $progressbar.Value = $i; Start-Sleep -Seconds 0.5} while ($i -lt 75)
    
    foreach ($department in $script:newDepartmentAD)
    {
    Add-ADGroupMember -Identity zzd-gbblw-igdoors-$department-rw -Members $script:newUsername # Access group - grants zpermissions file access based on their departments
        switch ($department) # Mail group - any relevant mail groups based on their departments
        {
            "finance-team"                      {Add-ADGroupMember -Identity DL-H-IGDoors-Accounts_Department -Members $script:newUsername}
            "customercare-team"                 {Add-ADGroupMember -Identity DL-H-IGDoors-Customer_Care -Members $script:newUsername}
            "newbuild"                          {Add-ADGroupMember -Identity DL-H-IGDoors-New_Build_Division -Members $script:newUsername}
            "socialhousing-team"                {Add-ADGroupMember -Identity DL-H-IGDoors-Social -Members $script:newUsername}
            "technical-team"                    {Add-ADGroupMember -Identity DL-H-IGDoors-Estimating -Members $script:newUsername}
            "quality-team"                      {Add-ADGroupMember -Identity DL-H-IGDoors-HSQE -Members $script:newUsername;
                                                 Add-ADGroupMember -Identity DL-H-IGDoors-Non-Conform -Members $script:newUsername}
            "MaterialsandWarehouse-Management"  {Add-ADGroupMember -Identity DL-H-IGDoors-Non-Conform -Members $script:newUsername}
            "Logistics"                         {Add-ADGroupMember -Identity DL-H-IGDoors-Operations_Department -Members $script:newUsername}
            "production"                        {Add-ADGroupMember -Identity DL-H-IGDoors-Operations_Department -Members $script:newUsername}
            Default                             {""}
        }
    }
    
    do {$i++; $progressbar.Value = $i; Start-Sleep -Seconds 0.5} while ($i -lt 90)

WriteLog "
Enable-Remotemailbox -Identity '$script:newUsername' -PrimarySmtpAddress '$script:newEmail' -Remoteroutingaddress '$script:newEmailLong'

Enable-RemoteMailbox $script:newEmail -Archive

Connect-ExchangeOnline -UserPrincipalName '$admUsername@igdoors.co.uk'

Set-Mailbox -Identity $script:newEmail -RetentionPolicy 'HGROUP Default RetentionPolicy'"

    AddToExcel

    do {$i++; $progressbar.Value = $i; Start-Sleep -Seconds 0.1} while ($i -lt 100)

    WriteLog "The task has run successfully.`n

    "
    # Close the script and open the log
    [System.Windows.Forms.MessageBox]::Show("The task has run successfully.")
    $form.Close()
    Invoke-Item ".\Add or edit AD user log.log"
}

function EditUser # Edit an existing user ; validate the username ; display the current details ; create zP groups array ; reveal the next components
{
    $script:selectNewOrEdit = "Edit"
    StartDeptArrays
    $gobuttonedit.Visible = $false
    $script:newUsername = $nameentry.Text
   
    $error.Clear()
    try {Get-ADUser $newUsername} catch {}
        if($error) {$nameerrorlabel.Visible = $true; $namenouserlabel.Visible = $true; return}
            else {$nameerrorlabel.Visible = $false; $nameinvalidlabel.Visible = $false; $namenouserlabel.Visible = $false}

    $script:editUserDetails = Get-ADUser -Identity $newUsername -Properties Department, MemberOf, Title
    $script:newFirstName = $script:editUserDetails.GivenName
    $script:newSurname = $script:editUserDetails.Surname
    $script:newDisplayName = $script:editUserDetails.Name
    $script:newEmail = $script:editUserDetails.UserPrincipalName
    $script:newJobTitle = $script:editUserDetails.Title
    $script:newDepartmentPrimary = $script:editUserDetails.Department
    $script:newMemberOf = $script:editUserDetails.MemberOf
    $titleentry.Text = $script:newJobTitle
    try {$script:newJobTitle = $script:newJobTitle.Replace("&", "&&")} catch {}
    $deptentry.Text = $script:newDepartmentPrimary
    try {$script:newDepartmentPrimary = $script:newDepartmentPrimary.Replace("&", "&&")} catch {}
    $script:newDepartmentFull += $deptentry.Text

    $script:newMemberOf = $script:newMemberOf -split "CN="
    foreach ($department in $script:newMemberOf)
    {
        switch -wildcard ($department)
        {
            '*3tec*'                    {$script:newDepartment += "3Tec";                               $script:newDepartmentAD += "3tec"}
            '*accounts*'                {$script:newDepartment += "Accounts Department";                } # $script:newDepartmentAD += ""
            '*aftersales-team*'         {$script:newDepartment += "After Sales Department";             $script:newDepartmentAD += "aftersales-team"}
            '*bridgetime-scans*'        {$script:newDepartment += "Bridgetime Scans";                   $script:newDepartmentAD += "Bridgetime-Scans"}
            '*CNC-team*'                {$script:newDepartment += "CNC Team";                           $script:newDepartmentAD += "CNC-Team"}
            '*customercare-team*'       {$script:newDepartment += "Customer Care Department";           $script:newDepartmentAD += "customercare-team"}
            '*data-exchange*'           {$script:newDepartment += "Data Exchange";                      $script:newDepartmentAD += "data-exchange"}
            '*despatch-images*'         {$script:newDepartment += "Despatch Images";                    $script:newDepartmentAD += "Despatch-Images"}
            '*engineering-team*'        {$script:newDepartment += "Engineering Department";             $script:newDepartmentAD += "engineering-team"}
            '*finance-team*'            {$script:newDepartment += "Finance Department";                 $script:newDepartmentAD += "finance-team"}
            '*HealthandSafety*'         {$script:newDepartment += "Health & Safety Department";         $script:newDepartmentAD += "HealthandSafety"}
            '*human-resources*'         {$script:newDepartment += "Human Resources Department";         $script:newDepartmentAD += "human-resources"}
            '*information-centre*'      {$script:newDepartment += "Information Centre";                 $script:newDepartmentAD += "Information-Centre"}
            '*it-team*'                 {$script:newDepartment += "I.T. Systems";                       $script:newDepartmentAD += "it-team"}
            '*Logistics*'               {$script:newDepartment += "Logistics";                          $script:newDepartmentAD += "Logistics"}
            '*bridgetime-rw*'           {$script:newDepartment += "Logistics (Bridgetime)";             $script:newDepartmentAD += "bridgetime"}
            '*MaterialsandWarehouse*'   {$script:newDepartment += "Materials & Warehouse Management";   $script:newDepartmentAD += "MaterialsandWarehouse-Management"}
            '*newbuild-rw*'             {$script:newDepartment += "New Build Division";                 $script:newDepartmentAD += "newbuild"}
            '*newbuild-management*'     {$script:newDepartment += "New Build Management";               $script:newDepartmentAD += "newbuild-management"}
            '*operations*'              {$script:newDepartment += "Operations Department";              } # $script:newDepartmentAD += ""
            '*orderprocessing-team*'    {$script:newDepartment += "Order Processing Department";        $script:newDepartmentAD += "orderprocessing-team"}
            '*doors-planning*'          {$script:newDepartment += "Planning Department";                $script:newDepartmentAD += "planning"}
            '*production-rw*'           {$script:newDepartment += "Production Department";              $script:newDepartmentAD += "production"}
            '*project-engineering*'     {$script:newDepartment += "Project Engineering Department";     $script:newDepartmentAD += "project-engineering"}
            '*purchasing*'              {$script:newDepartment += "Purchasing Department";              } # $script:newDepartmentAD += ""
            '*quality-team*'            {$script:newDepartment += "Quality & HSE Department";           $script:newDepartmentAD += "quality-team"}
            '*quality-images*'          {$script:newDepartment += "Quality Images";                     $script:newDepartmentAD += "quality-images"}
            '*randd-team*'              {$script:newDepartment += "Research & Development";             $script:newDepartmentAD += "randd-team"}
            '*doors-RMA*'               {$script:newDepartment += "RMA";                                $script:newDepartmentAD += "RMA"}
            '*sage-payroll*'            {$script:newDepartment += "Payroll Department";                 $script:newDepartmentAD += "sage-payroll"}
            '*doors-sales-team*'        {$script:newDepartment += "Sales Department";                   $script:newDepartmentAD += "sales-team"}
            '*senior-managers*'         {$script:newDepartment += "Senior Managers";                    $script:newDepartmentAD += "Senior-Managers"}
            '*seniorfinance-team*'      {$script:newDepartment += "Senior Finance Team";                $script:newDepartmentAD += "SeniorFinance-Team"}
            '*socialhousing-team*'      {$script:newDepartment += "Social Housing Department";          $script:newDepartmentAD += "socialhousing-team"}
            '*team-leaders*'            {$script:newDepartment += "Team Leaders";                       $script:newDepartmentAD += "team-leaders"}
            '*technical-team*'          {$script:newDepartment += "Technical Department";               $script:newDepartmentAD += "technical-team"}
            '*trade-team*'              {$script:newDepartment += "Trade Division";                     $script:newDepartmentAD += "trade-team"}
            '*allowTrustedLocations*'   {$trustedlocation.Checked = $true; $requireusermfa.Checked = $false; $script:accessGroupOriginal = "allowTrustedLocations"}
            '*requireUserMFA*'          {$trustedlocation.Checked = $false; $requireusermfa.Checked = $true; $script:accessGroupOriginal = "requireUserMFA"}
            Default     {""}
        }
    }
    $script:newDepartmentAD = $script:newDepartmentAD | Select-Object -Unique
    $script:newDepartmentADList = $script:newDepartmentAD -join '
    '
    $script:newDepartmentADLogList = $script:newDepartmentAD -join ', '
    $script:accessGroup = $script:accessGroupOriginal
    StartVisibility
    $resetdeptedit.Visible = $true
    $confirmedit.Visible = $true
    $gobuttonnew.Visible = $false
    EditLabelDetails
}

function ConfirmDetailsEdit # Read the job title entry ; run EachDeptInNew, EachDeptInDel, ReadAccessGroup ; update the label ; reveal the 'run' button
{
    $script:newJobTitle = $titleentry.Text
    EachDeptInNew
    EachDeptInDel
    ReadAccessGroup
    EditLabelDetails
    DeptErrorCheck
    $GoButtonEdit.Visible = $true
}

function GoButtonEdit # Write to the log ; set the new job title and department ; add/remove zP groups ; change Access Group if selection was changed ; add to zP excel list
{
    if ($script:accessGroup -match $script:accessGroupOriginal)
        {$script:accessGroupLog = "${script:accessGroup} (unchanged)"}
        else {$script:accessGroupLog = "${script:accessGroup} (changed from ${script:accessGroupOriginal})"}

    WriteLog "`n

    EDIT USER
    "
    WriteLog "The script started by user $admUsername."
    WriteLog "Edit user: $script:newFirstName $script:newSurname"
    WriteLog "Displayname: $script:newDisplayName"
    WriteLog "Login name: $script:newUsername"
    WriteLog "Email address: $script:newEmail"
    WriteLog "Job Title: $script:newJobTitle"
    WriteLog "Department: $script:newDepartmentPrimary"
    WriteLog "Active Directory groups: $script:newDepartmentADLogList"
    WriteLog "Removed from AD groups: $script:delDepartmentADLogList"
    WriteLog "Access group: $script:accessGroupLog"

    $i=0
    do {$i++; $progressbar.Value = $i; Start-Sleep -Seconds 0.5} while ($i -lt 20)

    Set-ADUser $newUsername -Title "$newJobTitle"
    Set-ADUser $newUsername -Department "$newDepartmentPrimary"
    
    do {$i++; $progressbar.Value = $i; Start-Sleep -Seconds 0.5} while ($i -lt 40)

    foreach ($department in $script:newDepartmentAD)
        {Add-ADGroupMember -Identity zzd-gbblw-igdoors-$department-rw -Members $script:newUsername}

    foreach ($department in $script:delDepartmentAD)
        {Remove-ADGroupMember -Identity zzd-gbblw-igdoors-$department-rw -Members $script:newUsername -Confirm:$false}

    do {$i++; $progressbar.Value = $i; Start-Sleep -Seconds 0.5} while ($i -lt 60)

    if ($script:accessGroup -notmatch $script:accessGroupOriginal)
    {
        Remove-ADGroupMember -Identity zzx-gbblw-m365-ConditionalAccess-$script:accessGroupOriginal -Members $script:newUsername -Confirm:$false;
        Add-ADGroupMember -Identity zzx-gbblw-m365-ConditionalAccess-$script:accessGroup -Members $script:newUsername
    }

    do {$i++; $progressbar.Value = $i; Start-Sleep -Seconds 0.5} while ($i -lt 80)

    AddToExcel

    do {$i++; $progressbar.Value = $i; Start-Sleep -Seconds 0.1} while ($i -lt 100)
        
    WriteLog "The task has run successfully.`n

    "

    [System.Windows.Forms.MessageBox]::Show("The task has run successfully.")
    $form.Close()
    Invoke-Item ".\Add or edit AD user log.log"
}

function AddToExcel # Creates csv to run the update zPermissions excel script
{
    $csvOutput = foreach ($dept in $script:newDepartmentAD)
        {
            New-Object -TypeName PSCustomObject -Property @{
            "Username" = $script:newUsername
            "First Name" = $script:newFirstName
            "Surname" = $script:newSurname
            "Department" = $dept}
        }
    $csvOutput | Export-Csv -Path "\\hgroup\data\GB-IGDoors\zPermissions\zPermissions to update.csv" -NoTypeInformation -Append
}

$newuser.Add_Click({NewUser})
$edituser.Add_Click({EditUser})
$resetall.Add_Click({ResetAll})
$adddept.Add_Click({AddDepartment})
$deldept.Add_Click({DelDepartment})
$confirmnew.Add_Click({ConfirmDetailsNew})
$confirmedit.Add_Click({ConfirmDetailsEdit})
$gobuttonnew.Add_Click({GoButtonNew})
$gobuttonedit.Add_Click({GoButtonEdit})
$resetdeptnew.Add_Click({ResetDeptNew})
$resetdeptedit.Add_Click({ResetDeptEdit})
$exitbutton.Add_Click({ExitButton})

$form.Add_Shown({$form.Activate()})
$form.ShowDialog()