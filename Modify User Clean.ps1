<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    ModifyUser
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '800,650'
$Form.text                       = "Form"
$Form.TopMost                    = $false

$GroupboxNewSetting              = New-Object system.Windows.Forms.Groupbox
$GroupboxNewSetting.height       = 145
$GroupboxNewSetting.width        = 271
$GroupboxNewSetting.text         = "New Settings:"
$GroupboxNewSetting.location     = New-Object System.Drawing.Point(415,310)

$GroupboxSearchMatchUser         = New-Object system.Windows.Forms.Groupbox
$GroupboxSearchMatchUser.height  = 95
$GroupboxSearchMatchUser.width   = 360
$GroupboxSearchMatchUser.text    = "Match User:"
$GroupboxSearchMatchUser.location  = New-Object System.Drawing.Point(220,187)

$GroupboxCurrentSettings         = New-Object system.Windows.Forms.Groupbox
$GroupboxCurrentSettings.height  = 145
$GroupboxCurrentSettings.width   = 271
$GroupboxCurrentSettings.text    = "Current Settings:"
$GroupboxCurrentSettings.location  = New-Object System.Drawing.Point(132,310)

$TextBoxSearchModifyUser         = New-Object system.Windows.Forms.TextBox
$TextBoxSearchModifyUser.multiline  = $false
$TextBoxSearchModifyUser.text    = "Enter a portion of the user name"
$TextBoxSearchModifyUser.width   = 200
$TextBoxSearchModifyUser.height  = 20
$TextBoxSearchModifyUser.location  = New-Object System.Drawing.Point(20,25)
$TextBoxSearchModifyUser.Font    = 'Microsoft Sans Serif,10'

$TextBoxSelectedModifyUser       = New-Object system.Windows.Forms.TextBox
$TextBoxSelectedModifyUser.multiline  = $false
$TextBoxSelectedModifyUser.width  = 200
$TextBoxSelectedModifyUser.height  = 20
$TextBoxSelectedModifyUser.location  = New-Object System.Drawing.Point(20,60)
$TextBoxSelectedModifyUser.Font  = 'Microsoft Sans Serif,10'
$TextBoxSelectedModifyUser.ReadOnly = $true

$TextBoxSearchMatchUser          = New-Object system.Windows.Forms.TextBox
$TextBoxSearchMatchUser.multiline  = $false
$TextBoxSearchMatchUser.text     = "Enter a portion of the user name"
$TextBoxSearchMatchUser.width    = 200
$TextBoxSearchMatchUser.height   = 20
$TextBoxSearchMatchUser.location  = New-Object System.Drawing.Point(26,24)
$TextBoxSearchMatchUser.Font     = 'Microsoft Sans Serif,10'

$TextBoxSelectedMatchUser        = New-Object system.Windows.Forms.TextBox
$TextBoxSelectedMatchUser.multiline  = $false
$TextBoxSelectedMatchUser.width  = 200
$TextBoxSelectedMatchUser.height  = 20
$TextBoxSelectedMatchUser.location  = New-Object System.Drawing.Point(25,61)
$TextBoxSelectedMatchUser.Font   = 'Microsoft Sans Serif,10'
$TextBoxSelectedMatchUser.ReadOnly = $true

$RadioButtonSelectNewInfo        = New-Object system.Windows.Forms.RadioButton
$RadioButtonSelectNewInfo.text   = "Select New Info"
$RadioButtonSelectNewInfo.AutoSize  = $true
$RadioButtonSelectNewInfo.width  = 104
$RadioButtonSelectNewInfo.height  = 20
$RadioButtonSelectNewInfo.location  = New-Object System.Drawing.Point(20,10)
$RadioButtonSelectNewInfo.Font   = 'Microsoft Sans Serif,10'
$RadioButtonSelectNewInfo.Checked = $true

$RadioButtonMatchExistingUser    = New-Object system.Windows.Forms.RadioButton
$RadioButtonMatchExistingUser.text  = "Match Existing User"
$RadioButtonMatchExistingUser.AutoSize  = $true
$RadioButtonMatchExistingUser.width  = 104
$RadioButtonMatchExistingUser.height  = 20
$RadioButtonMatchExistingUser.location  = New-Object System.Drawing.Point(180,10)
$RadioButtonMatchExistingUser.Font  = 'Microsoft Sans Serif,10'

$ButtonSearchModifyUser          = New-Object system.Windows.Forms.Button
$ButtonSearchModifyUser.text     = "Search"
$ButtonSearchModifyUser.width    = 60
$ButtonSearchModifyUser.height   = 30
$ButtonSearchModifyUser.location  = New-Object System.Drawing.Point(275,33)
$ButtonSearchModifyUser.Font     = 'Microsoft Sans Serif,10'

$ButtonSearchMatchUser           = New-Object system.Windows.Forms.Button
$ButtonSearchMatchUser.text      = "Search"
$ButtonSearchMatchUser.width     = 60
$ButtonSearchMatchUser.height    = 30
$ButtonSearchMatchUser.location  = New-Object System.Drawing.Point(275,33)
$ButtonSearchMatchUser.Font      = 'Microsoft Sans Serif,10'

$ComboBoxNewMainDept             = New-Object system.Windows.Forms.ComboBox
$ComboBoxNewMainDept.text        = "New Main Dept"
$ComboBoxNewMainDept.width       = 175
$ComboBoxNewMainDept.height      = 20
$ComboBoxNewMainDept.location    = New-Object System.Drawing.Point(10,30)
$ComboBoxNewMainDept.Font        = 'Microsoft Sans Serif,10'

$ComboBoxNewSubDept              = New-Object system.Windows.Forms.ComboBox
$ComboBoxNewSubDept.text         = "New Sub Dept"
$ComboBoxNewSubDept.width        = 175
$ComboBoxNewSubDept.height       = 20
$ComboBoxNewSubDept.location     = New-Object System.Drawing.Point(10,70)
$ComboBoxNewSubDept.Font         = 'Microsoft Sans Serif,10'

$ComboBoxNewJobTitle             = New-Object system.Windows.Forms.ComboBox
$ComboBoxNewJobTitle.text        = "New Job Title"
$ComboBoxNewJobTitle.width       = 250
$ComboBoxNewJobTitle.height      = 20
$ComboBoxNewJobTitle.location    = New-Object System.Drawing.Point(10,110)
$ComboBoxNewJobTitle.Font        = 'Microsoft Sans Serif,10'

$ButtonCommitChanges             = New-Object system.Windows.Forms.Button
$ButtonCommitChanges.text        = "Commit Changes"
$ButtonCommitChanges.width       = 150
$ButtonCommitChanges.height      = 30
$ButtonCommitChanges.location    = New-Object System.Drawing.Point(295,475)
$ButtonCommitChanges.Font        = 'Microsoft Sans Serif,12,style=Bold'

$LabelMainDept                   = New-Object system.Windows.Forms.Label
$LabelMainDept.text              = "Main Deparatment:"
$LabelMainDept.AutoSize          = $true
$LabelMainDept.width             = 25
$LabelMainDept.height            = 10
$LabelMainDept.location          = New-Object System.Drawing.Point(12,332)
$LabelMainDept.Font              = 'Microsoft Sans Serif,10'

$LabelSubDept                    = New-Object system.Windows.Forms.Label
$LabelSubDept.text               = "Sub Department:"
$LabelSubDept.AutoSize           = $true
$LabelSubDept.width              = 25
$LabelSubDept.height             = 10
$LabelSubDept.location           = New-Object System.Drawing.Point(12,373)
$LabelSubDept.Font               = 'Microsoft Sans Serif,10'

$LabelJobTitle                   = New-Object system.Windows.Forms.Label
$LabelJobTitle.text              = "Job Title:"
$LabelJobTitle.AutoSize          = $true
$LabelJobTitle.width             = 25
$LabelJobTitle.height            = 10
$LabelJobTitle.location          = New-Object System.Drawing.Point(12,414)
$LabelJobTitle.Font              = 'Microsoft Sans Serif,10'

$GroupboxSelectOptions           = New-Object system.Windows.Forms.Groupbox
$GroupboxSelectOptions.height    = 35
$GroupboxSelectOptions.width     = 335
$GroupboxSelectOptions.location  = New-Object System.Drawing.Point(235,130)

$GroupboxSearchModifyUser        = New-Object system.Windows.Forms.Groupbox
$GroupboxSearchModifyUser.height  = 95
$GroupboxSearchModifyUser.width  = 360
$GroupboxSearchModifyUser.text   = "Group Box"
$GroupboxSearchModifyUser.location  = New-Object System.Drawing.Point(220,20)

$TextBoxCurrentMainDepartment    = New-Object system.Windows.Forms.TextBox
$TextBoxCurrentMainDepartment.multiline  = $false
$TextBoxCurrentMainDepartment.width  = 175
$TextBoxCurrentMainDepartment.height  = 20
$TextBoxCurrentMainDepartment.location  = New-Object System.Drawing.Point(10,30)
$TextBoxCurrentMainDepartment.Font  = 'Microsoft Sans Serif,10'
$TextBoxCurrentMainDepartment.ReadOnly = $true

$TextBoxCurrentSubDepartment     = New-Object system.Windows.Forms.TextBox
$TextBoxCurrentSubDepartment.multiline  = $false
$TextBoxCurrentSubDepartment.width  = 175
$TextBoxCurrentSubDepartment.height  = 20
$TextBoxCurrentSubDepartment.location  = New-Object System.Drawing.Point(10,70)
$TextBoxCurrentSubDepartment.Font  = 'Microsoft Sans Serif,10'
$TextBoxCurrentSubDepartment.ReadOnly = $true

$TextBoxCurrentJobTitle          = New-Object system.Windows.Forms.TextBox
$TextBoxCurrentJobTitle.multiline  = $false
$TextBoxCurrentJobTitle.width    = 250
$TextBoxCurrentJobTitle.height   = 20
$TextBoxCurrentJobTitle.location  = New-Object System.Drawing.Point(10,110)
$TextBoxCurrentJobTitle.Font     = 'Microsoft Sans Serif,10'
$TextBoxCurrentJobTitle.ReadOnly = $true

$Form.controls.AddRange(@($GroupboxNewSetting,$GroupboxSearchMatchUser,$GroupboxCurrentSettings,$ButtonCommitChanges,$LabelMainDept,$LabelSubDept,$LabelJobTitle,$GroupboxSelectOptions,$GroupboxSearchModifyUser))
$GroupboxSearchModifyUser.controls.AddRange(@($TextBoxSearchModifyUser,$TextBoxSelectedModifyUser,$ButtonSearchModifyUser))
$GroupboxSearchMatchUser.controls.AddRange(@($TextBoxSearchMatchUser,$TextBoxSelectedMatchUser,$ButtonSearchMatchUser))
$GroupboxSelectOptions.controls.AddRange(@($RadioButtonSelectNewInfo,$RadioButtonMatchExistingUser))
$GroupboxNewSetting.controls.AddRange(@($ComboBoxNewMainDept,$ComboBoxNewSubDept,$ComboBoxNewJobTitle))
$GroupboxCurrentSettings.controls.AddRange(@($TextBoxCurrentMainDepartment,$TextBoxCurrentSubDepartment,$TextBoxCurrentJobTitle))

#######################################################
##### Second Form - Select Modify User dialog box #####
#######################################################

$FormModifyUserSelection         = New-Object system.Windows.Forms.Form
$FormModifyUserSelection.ClientSize  = '400,400'
$FormModifyUserSelection.text    = "Select a User"
$FormModifyUserSelection.TopMost  = $false
$FormModifyUserSelection.StartPosition = "manual"

$ListBoxSelectModifyUser         = New-Object system.Windows.Forms.ListBox
$ListBoxSelectModifyUser.text    = "listBox"
$ListBoxSelectModifyUser.width   = 340
$ListBoxSelectModifyUser.height  = 300
$ListBoxSelectModifyUser.location  = New-Object System.Drawing.Point(30,30)

$ButtonSelectModifyUser          = New-Object system.Windows.Forms.Button
$ButtonSelectModifyUser.text     = "Select"
$ButtonSelectModifyUser.width    = 80
$ButtonSelectModifyUser.height   = 30
$ButtonSelectModifyUser.location  = New-Object System.Drawing.Point(147,345)
$ButtonSelectModifyUser.Font     = 'Microsoft Sans Serif,14'

$FormModifyUserSelection.controls.AddRange(@($ListBoxSelectModifyUser,$ButtonSelectModifyUser))

#######################################
##### End Select Modify User form #####
#######################################

#####################################################
##### Third Form - Select Match User dialog box #####
#####################################################

$FormMatchUserSelection         = New-Object system.Windows.Forms.Form
$FormMatchUserSelection.ClientSize  = '400,400'
$FormMatchUserSelection.text    = "Select a User"
$FormMatchUserSelection.TopMost  = $false
$FormMatchUserSelection.StartPosition = "manual"

$ListBoxSelectMatchUser         = New-Object system.Windows.Forms.ListBox
$ListBoxSelectMatchUser.text    = "listBox"
$ListBoxSelectMatchUser.width   = 340
$ListBoxSelectMatchUser.height  = 300
$ListBoxSelectMatchUser.location  = New-Object System.Drawing.Point(30,30)

$ButtonSelectMatchUser          = New-Object system.Windows.Forms.Button
$ButtonSelectMatchUser.text     = "Select"
$ButtonSelectMatchUser.width    = 80
$ButtonSelectMatchUser.height   = 30
$ButtonSelectMatchUser.location  = New-Object System.Drawing.Point(147,345)
$ButtonSelectMatchUser.Font     = 'Microsoft Sans Serif,14'

$FormMatchUserSelection.controls.AddRange(@($ListBoxSelectMatchUser,$ButtonSelectMatchUser))

###############################
##### End Match User Form #####
###############################

################################
##### Set initial settings #####
################################
$GroupboxSearchMatchUser.Enabled = $false
$ButtonSearchModifyUser.Enabled = $false
$ListBoxSelectModifyUser.DisplayMember = "name"
$ListBoxSelectMatchUser.DisplayMember = "name"
$ComboBoxNewMainDept.DisplayMember = "name"
$ComboBoxNewSubDept.DisplayMember = "name"
$ComboBoxNewJobTitle.DisplayMember = "name"

##############################
##### Add Event Handlers #####
##############################

$RadioButtonSelectNewInfo.Add_CheckedChanged($RadioButtonSelectNewInfo_CheckedChanged)
$RadioButtonMatchExistingUser.Add_CheckedChanged($RadioButtonMatchExistingUser_CheckedChanged)
$TextBoxSearchModifyUser.Add_TextChanged($TextBoxSearchModifyUser_TextChanged)
$TextBoxSearchModifyUser.Add_Enter($TextBoxSearchModifyUser_Enter)
$TextBoxSearchModifyUser.Add_Click($TextBoxSearchModifyUser_Click)
$TextBoxSearchModifyUser.Add_Leave($TextBoxSearchModifyUser_Leave)
$TextBoxSearchMatchUser.Add_TextChanged($TextBoxSearchMatchUser_TextChanged)
$TextBoxSearchMatchUser.Add_Enter($TextBoxSearchMatchUser_Enter)
$TextBoxSearchMatchUser.Add_Click($TextBoxSearchMatchUser_Click)
$TextBoxSearchMatchUser.Add_Leave($TextBoxSearchMatchUser_Leave)
$ButtonSearchModifyUser.Add_Click($ButtonSearchModifyUser_Click)
$ButtonSearchMatchUser.Add_Click($ButtonSearchMatchUser_Click)
$ComboBoxNewMainDept.Add_SelectedIndexChanged($ComboBoxNewMainDept_SelectedIndexChanged)
$ComboBoxNewSubDept.Add_SelectedIndexChanged($ComboBoxNewSubDept_SelectedIndexChanged)
$ComboBoxNewJobTitle.Add_SelectedIndexChanged($ComboBoxNewJobTitle_SelectedIndexChanged)

######################################################
##### Add Event Handlers Select Modify User Forms #####
######################################################

$ButtonSelectModifyUser.Add_Click($ButtonSelectModifyUser_Click)
$ButtonSelectMatchUser.Add_Click($ButtonSelectMatchUser_Click)

##########################
##### Event Handlers #####
##########################

$RadioButtonSelectNewInfo_CheckedChanged = {
    if ($RadioButtonSelectNewInfo.Checked){
        #$RadioButtonMatchExistingUser.Checked = $false
        $GroupboxSearchMatchUser.Enabled = $false
        $ButtonSearchMatchUser.Enabled = $false
        $GroupboxNewSetting.Enabled = $true
        Clear-NewSettings
        $Mains = Get-ADOrganizationalUnit -SearchBase 'OU=Garden City,DC=GardenCityks,DC=US' -Filter {description -like "D - *"} | select name,distinguishedname | sort name

        foreach($dept in $Mains){
            $ComboBoxNewMainDept.items.Add($dept)
        }
    }
}

$RadioButtonMatchExistingUser_CheckedChanged = {
    if ($RadioButtonMatchExistingUser.Checked){
        #$RadioButtonSelectNewInfo.Checked = $false
        $GroupboxSearchMatchUser.Enabled = $true
        $ButtonSearchMatchUser.Enabled = $false
        $GroupboxNewSetting.Enabled = $false
        Clear-NewSettings
    }
}

$ComboBoxNewMainDept_SelectedIndexChanged = {
    $MainDept = $ComboBoxNewMainDept.SelectedItem.distinguishedname
    $SubFind = Get-ADOrganizationalUnit -Searchbase $MainDept -SearchScope OneLevel -Filter {Name -like "Us*"} | select name,distinguishedname | sort name
    $ComboBoxNewSubDept.items.Clear()
    #$CBSubDepartment.Text = "Select Sub Department"
    clearTitle
    If ($SubFind){
        $SubDepts = 'null'
    
        $Titles = Get-ADGroup -Searchbase $MainDept -Filter {description -like "RBAC*"} | select name,distinguishedname | sort name
    
        foreach($job in $Titles){
            $ComboBoxNewJobTitle.items.Add($job)
            }
        $ComboBoxNewSubDept.Text = ""
    }

    Else{
        $SubDepts = Get-ADOrganizationalUnit -SearchBase $MainDept -SearchScope OneLevel -Filter {Description -like "*"} | select name,distinguishedname | sort name
        foreach($dept in $SubDepts){
            $ComboBoxNewSubDept.items.Add($dept)
            }
        $ComboBoxNewSubDept.Text = "New Sub Dept"
        }
}

$ComboBoxNewSubDept_SelectedIndexChanged = {
    clearTitle
    $MainDept = $ComboBoxNewMainDept.SelectedItem.distinguishedname
    $SubDept = $ComboBoxNewSubDept.SelectedItem.distinguishedname

    $Titles = Get-ADGroup -Searchbase $SubDept -Filter {description -like "RBAC*"} | select name,distinguishedname | sort name
    #Write-Host "JobTitle selected index: " $ComboBoxNewJobTitle.SelectedIndex
    foreach($job in $Titles){
        $ComboBoxNewJobTitle.items.Add($job)
        }
}

$TextBoxSearchModifyUser_TextChanged = {
    if ($TextBoxSearchModifyUser.TextLength -gt 0 -and $TextBoxSearchModifyUser.Text -ne "Enter a portion of the user name"){
        $ButtonSearchModifyUser.Enabled = $true
    }
    else {$ButtonSearchModifyUser.Enabled = $false}
}
$TextBoxSearchModifyUser_Enter = {
    if ($TextBoxSearchModifyUser.Text -eq "Enter a portion of the user name"){
        $TextBoxSearchModifyUser.SelectAll()
    }
}
$TextBoxSearchModifyUser_Click = {
    if ($TextBoxSearchModifyUser.Text -eq "Enter a portion of the user name"){
        $TextBoxSearchModifyUser.SelectAll()
    }
}
$TextBoxSearchModifyUser_Leave = {
    if ($TextBoxSearchModifyUser.TextLength -eq 0){
        $TextBoxSearchModifyUser.Text = "Enter a portion of the user name"
    }
}

$TextBoxSearchMatchUser_TextChanged = {
    if ($TextBoxSearchMatchUser.TextLength -gt 0 -and $TextBoxSearchMatchUser.Text -ne "Enter a portion of the user name"){
        $ButtonSearchMatchUser.Enabled = $true
    }
    else {$ButtonSearchMatchUser.Enabled = $false}
}
$TextBoxSearchMatchUser_Enter = {
    if ($TextBoxSearchMatchUser.Text -eq "Enter a portion of the user name"){
        $TextBoxSearchMatchUser.SelectAll()
    }
}
$TextBoxSearchMatchUser_Click = {
    if ($TextBoxSearchMatchUser.Text -eq "Enter a portion of the user name"){
        $TextBoxSearchMatchUser.SelectAll()
    }
}
$TextBoxSearchMatchUser_Leave = {
    if ($TextBoxSearchMatchUser.TextLength -eq 0){
        $TextBoxSearchMatchUser.Text = "Enter a portion of the user name"
    }
}

$ButtonSearchModifyUser_Click = {
    $ListBoxSelectModifyUser.Items.Clear()
    $NameSearch = $TextBoxSearchModifyUser.Text
    $NameSearch = "*" + $NameSearch + "*"
    $FoundUsers = Get-ADUser -Filter {Name -like $NameSearch} | select name,distinguishedname,samaccountname | sort name
    foreach ($user in $FoundUsers){
        $ListBoxSelectModifyUser.Items.Add($user)
    }
    $X = $Form.Location.X + 200
    $Y = $form.Location.Y + 20
    $FormModifyUserSelection.Location = "$X,$Y"
    [void]$FormModifyUserSelection.ShowDialog()
}

$ButtonSearchMatchUser_Click = {
    $ListBoxSelectMatchUser.Items.Clear()
    $NameSearch = $TextBoxSearchMatchUser.Text
    $NameSearch = "*" + $NameSearch + "*"
    $FoundUsers = Get-ADUser -Filter {Name -like $NameSearch} | select name,distinguishedname,samaccountname | sort name
    foreach ($user in $FoundUsers){
        $ListBoxSelectMatchUser.Items.Add($user)
    }
    $X = $Form.Location.X + 200
    $Y = $form.Location.Y + 20
    $FormMatchUserSelection.Location = "$X,$Y"
    [void]$FormMatchUserSelection.ShowDialog()
}

$ButtonSelectModifyUser_Click = {
    $TextBoxSelectedModifyUser.Text = $ListBoxSelectModifyUser.SelectedItem.name
    
    #$OU = (Get-ADUser -Identity $ListBoxSelectModifyUser.SelectedItem.distinguishedname).distinguishedname
    $OU = ($ListBoxSelectModifyUser.SelectedItem.distinguishedname -split ",",3)[2]
    $ParentOU = ($OU -split ",",2)[1]
    $isSubDept = ((Get-ADOrganizationalUnit -Identity $ParentOU -Properties description).description) -like "D - *"
    if ($isSubDept) {
        #$TextBoxCurrentMainDepartment.Text = (Get-ADOrganizationalUnit -Identity $ParentOU | select name).name
        #$TextBoxCurrentSubDepartment.Text = (Get-ADOrganizationalUnit -Identity $OU | select name).name
        $TextBoxCurrentMainDepartment.Text = (($ParentOU -split ",",2)[0] -split "=",2)[1]
        $TextBoxCurrentSubDepartment.Text = (($OU -split ",",2)[0] -split "=",2)[1]
    }
    else {
        #$TextBoxCurrentMainDepartment.Text = (Get-ADOrganizationalUnit -Identity $OU | select name).name
        $TextBoxCurrentMainDepartment.Text = (($OU -split ",",2)[0] -split "=",2)[1]
    }
    $TextBoxCurrentJobTitle.Text = (Get-ADPrincipalGroupMembership $ListBoxSelectModifyUser.SelectedItem.samaccountname | Where-Object { $_ -like "*Role*"}).name
    #$TextBoxCurrentJobTitle.Text = $OU
    $FormModifyUserSelection.close()
}

$ButtonSelectMatchUser_Click = {
    $TextBoxSelectedMatchUser.Text = $ListBoxSelectMatchUser.SelectedItem.name

    #$OU = (Get-ADUser -Identity $ListBoxSelectMatchUser.SelectedItem.distinguishedname).distinguishedname
    $OU = ($ListBoxSelectMatchUser.SelectedItem.distinguishedname -split ",",3)[2]
    $ParentOU = ($OU -split ",",2)[1]
    $isSubDept = ((Get-ADOrganizationalUnit -Identity $ParentOU -Properties description).description) -like "D - *"
    if ($isSubDept) {
        #$ComboboxNewMainDept.Text = (Get-ADOrganizationalUnit -Identity $ParentOU | select name).name
        #$ComboboxNewSubDept.Text = (Get-ADOrganizationalUnit -Identity $OU | select name).name
        $ComboboxNewMainDept.Text = (($ParentOU -split ",",2)[0] -split "=",2)[1]
        $ComboboxNewMainDept.SelectedItem = $ParentOU
        $ComboboxNewSubDept.Text = (($OU -split ",",2)[0] -split "=",2)[1]
        $ComboboxNewSubDept.SelectedItem = $OU
    }
    else {
        #$ComboboxNewMainDept.Text = (Get-ADOrganizationalUnit -Identity $OU | select name).name
        $ComboboxNewMainDept.Text = (($OU -split ",",2)[0] -split "=",2)[1]
        $ComboboxNewMainDept.SelectedItem = "$OU"
    }
    $ComboboxNewJobTitle.Text = (Get-ADPrincipalGroupMembership $ListBoxSelectMatchUser.SelectedItem.samaccountname | Where-Object { $_ -like "*Role*"}).name
    #$ComboboxNewJobTitle.Text = $OU
    $FormMatchUserSelection.close()
}

#####################
##### Functions #####
#####################

function Clear-NewSettings {
    $ComboBoxNewMainDept.Items.Clear()
    $ComboBoxNewMainDept.SelectedIndex = -1
    $ComboBoxNewMainDept.Text = "New Main Dept"
    $ComboBoxNewSubDept.Items.Clear()
    $ComboBoxNewSubDept.SelectedIndex = -1
    $ComboBoxNewSubDept.Text = "New Sub Dept"
    $ComboBoxNewJobTitle.Items.Clear()
    $ComboBoxNewJobTitle.SelectedIndex = -1
    $ComboBoxNewJobTitle.Text = "New Job Title"
}

function clearTitle {
    $ComboBoxNewJobTitle.items.Clear()
    $ComboBoxNewJobTitle.Text = "New Job Title"
    $ComboBoxNewJobTitle.SelectedIndex = -1
    #$global:Title = $null
    #$global:TitleDN = $null
}


###############
##### Run #####
###############
$Mains = Get-ADOrganizationalUnit -SearchBase 'OU=Garden City,DC=GardenCityks,DC=US' -Filter {description -like "D - *"} | select name,distinguishedname | sort name

foreach($dept in $Mains){
    $ComboBoxNewMainDept.items.Add($dept)
    }


[void]$Form.ShowDialog()