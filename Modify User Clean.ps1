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

$TextBoxSearchModifyUser         = New-Object system.Windows.Forms.TextBox
$TextBoxSearchModifyUser.multiline  = $false
$TextBoxSearchModifyUser.text    = "Enter a portion of the user name"
$TextBoxSearchModifyUser.width   = 200
$TextBoxSearchModifyUser.height  = 20
$TextBoxSearchModifyUser.location  = New-Object System.Drawing.Point(286,49)
$TextBoxSearchModifyUser.Font    = 'Microsoft Sans Serif,10'

$TextBoxSelectedModifyUser       = New-Object system.Windows.Forms.TextBox
$TextBoxSelectedModifyUser.multiline  = $false
$TextBoxSelectedModifyUser.width  = 200
$TextBoxSelectedModifyUser.height  = 20
$TextBoxSelectedModifyUser.location  = New-Object System.Drawing.Point(286,81)
$TextBoxSelectedModifyUser.Font  = 'Microsoft Sans Serif,10'

$TextBoxSearchMatchUser          = New-Object system.Windows.Forms.TextBox
$TextBoxSearchMatchUser.multiline  = $false
$TextBoxSearchMatchUser.text     = "Enter a portion of the user name"
$TextBoxSearchMatchUser.width    = 200
$TextBoxSearchMatchUser.height   = 20
$TextBoxSearchMatchUser.location  = New-Object System.Drawing.Point(284,164)
$TextBoxSearchMatchUser.Font     = 'Microsoft Sans Serif,10'
$TextBoxSearchMatchUser.Enabled  = $false

$TextBoxSelectedMatchUser        = New-Object system.Windows.Forms.TextBox
$TextBoxSelectedMatchUser.multiline  = $false
$TextBoxSelectedMatchUser.width  = 200
$TextBoxSelectedMatchUser.height  = 20
$TextBoxSelectedMatchUser.location  = New-Object System.Drawing.Point(286,203)
$TextBoxSelectedMatchUser.Font   = 'Microsoft Sans Serif,10'
$TextBoxSelectedMatchUser.Enabled = $false

$RadioButtonSelectNewInfo        = New-Object system.Windows.Forms.RadioButton
$RadioButtonSelectNewInfo.text   = "Select New Info"
$RadioButtonSelectNewInfo.AutoSize  = $true
$RadioButtonSelectNewInfo.width  = 104
$RadioButtonSelectNewInfo.height  = 20
$RadioButtonSelectNewInfo.location  = New-Object System.Drawing.Point(223,127)
$RadioButtonSelectNewInfo.Font   = 'Microsoft Sans Serif,10'
$RadioButtonSelectNewInfo.Checked = $true

$RadioButtonMatchExistingUser    = New-Object system.Windows.Forms.RadioButton
$RadioButtonMatchExistingUser.text  = "Match Existing User"
$RadioButtonMatchExistingUser.AutoSize  = $true
$RadioButtonMatchExistingUser.width  = 104
$RadioButtonMatchExistingUser.height  = 20
$RadioButtonMatchExistingUser.location  = New-Object System.Drawing.Point(414,127)
$RadioButtonMatchExistingUser.Font  = 'Microsoft Sans Serif,10'
$RadioButtonMatchExistingUser.Checked

$LabelModifyUser                 = New-Object system.Windows.Forms.Label
$LabelModifyUser.text            = "Modify:"
$LabelModifyUser.AutoSize        = $true
$LabelModifyUser.width           = 25
$LabelModifyUser.height          = 10
$LabelModifyUser.location        = New-Object System.Drawing.Point(179,64)
$LabelModifyUser.Font            = 'Microsoft Sans Serif,14,style=Bold'

$LabelMatchUser                  = New-Object system.Windows.Forms.Label
$LabelMatchUser.text             = "Match:"
$LabelMatchUser.AutoSize         = $true
$LabelMatchUser.width            = 25
$LabelMatchUser.height           = 10
$LabelMatchUser.location         = New-Object System.Drawing.Point(179,185)
$LabelMatchUser.Font             = 'Microsoft Sans Serif,14,style=Bold'

$ButtonSearchModifyUser          = New-Object system.Windows.Forms.Button
$ButtonSearchModifyUser.text     = "Search"
$ButtonSearchModifyUser.width    = 60
$ButtonSearchModifyUser.height   = 30
$ButtonSearchModifyUser.location  = New-Object System.Drawing.Point(544,64)
$ButtonSearchModifyUser.Font     = 'Microsoft Sans Serif,10'
$ButtonSearchModifyUser.Enabled  = $false

$ButtonSearchMatchUser                 = New-Object system.Windows.Forms.Button
$ButtonSearchMatchUser.text            = "Search"
$ButtonSearchMatchUser.width           = 60
$ButtonSearchMatchUser.height          = 30
$ButtonSearchMatchUser.location        = New-Object System.Drawing.Point(544,178)
$ButtonSearchMatchUser.Font            = 'Microsoft Sans Serif,10'

$ComboBoxNewMainDept             = New-Object system.Windows.Forms.ComboBox
$ComboBoxNewMainDept.text        = "comboBox"
$ComboBoxNewMainDept.width       = 175
$ComboBoxNewMainDept.height      = 20
$ComboBoxNewMainDept.location    = New-Object System.Drawing.Point(430,329)
$ComboBoxNewMainDept.Font        = 'Microsoft Sans Serif,10'

$LabelNewInfo                    = New-Object system.Windows.Forms.Label
$LabelNewInfo.text               = "New Settings:"
$LabelNewInfo.AutoSize           = $true
$LabelNewInfo.width              = 25
$LabelNewInfo.height             = 10
$LabelNewInfo.location           = New-Object System.Drawing.Point(460,292)
$LabelNewInfo.Font               = 'Microsoft Sans Serif,14,style=Bold'

$LabelCurrentInfo                = New-Object system.Windows.Forms.Label
$LabelCurrentInfo.text           = "Current Settings:"
$LabelCurrentInfo.AutoSize       = $true
$LabelCurrentInfo.width          = 25
$LabelCurrentInfo.height         = 10
$LabelCurrentInfo.location       = New-Object System.Drawing.Point(160,260)#292
$LabelCurrentInfo.Font           = 'Microsoft Sans Serif,14,style=Bold'

$ComboBoxNewSubDept              = New-Object system.Windows.Forms.ComboBox
$ComboBoxNewSubDept.text         = "comboBox"
$ComboBoxNewSubDept.width        = 175
$ComboBoxNewSubDept.height       = 20
$ComboBoxNewSubDept.location     = New-Object System.Drawing.Point(430,369)
$ComboBoxNewSubDept.Font         = 'Microsoft Sans Serif,10'

$ComboBoxCurrentMainDept         = New-Object system.Windows.Forms.ComboBox
$ComboBoxCurrentMainDept.text    = "comboBox"
$ComboBoxCurrentMainDept.width   = 175
$ComboBoxCurrentMainDept.height  = 20
$ComboBoxCurrentMainDept.location  = New-Object System.Drawing.Point(10,30)#140,329
$ComboBoxCurrentMainDept.Font    = 'Microsoft Sans Serif,10'

$ComboBoxCurrentSubDept          = New-Object system.Windows.Forms.ComboBox
$ComboBoxCurrentSubDept.text     = "comboBox"
$ComboBoxCurrentSubDept.width    = 175
$ComboBoxCurrentSubDept.height   = 20
$ComboBoxCurrentSubDept.location  = New-Object System.Drawing.Point(10,70)#140,369
$ComboBoxCurrentSubDept.Font     = 'Microsoft Sans Serif,10'

$ComboBoxCurrentJobTitle         = New-Object system.Windows.Forms.ComboBox
$ComboBoxCurrentJobTitle.text    = "comboBox"
$ComboBoxCurrentJobTitle.width   = 250
$ComboBoxCurrentJobTitle.height  = 20
$ComboBoxCurrentJobTitle.location  = New-Object System.Drawing.Point(10,110)#140,408
$ComboBoxCurrentJobTitle.Font    = 'Microsoft Sans Serif,10'

$ComboBoxNewJobTitle             = New-Object system.Windows.Forms.ComboBox
$ComboBoxNewJobTitle.text        = "comboBox"
$ComboBoxNewJobTitle.width       = 250
$ComboBoxNewJobTitle.height      = 20
$ComboBoxNewJobTitle.location    = New-Object System.Drawing.Point(430,408)
$ComboBoxNewJobTitle.Font        = 'Microsoft Sans Serif,10'

$ButtonCommitChanges             = New-Object system.Windows.Forms.Button
$ButtonCommitChanges.text        = "Commit Changes"
$ButtonCommitChanges.width       = 150
$ButtonCommitChanges.height      = 30
$ButtonCommitChanges.location    = New-Object System.Drawing.Point(295,466)
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

$GroupboxCurrentSettings          = New-Object system.Windows.Forms.Groupbox
$GroupboxCurrentSettings.height   = 145
$GroupboxCurrentSettings.width    = 271
$GroupboxCurrentSettings.text     = "Current Settings:"
$GroupboxCurrentSettings.location  = New-Object System.Drawing.Point(133,298)

##############################
##### Add Event Handlers #####
##############################

$RadioButtonSelectNewInfo.Add_CheckedChanged($RadioButtonSelectNewInfo_CheckedChanged)
$RadioButtonMatchExistingUser.Add_CheckedChanged($RadioButtonMatchExistingUser_CheckedChanged)

$Form.controls.AddRange(@($GroupboxCurrentSettings,$TextBoxSearchModifyUser,$TextBoxSelectedModifyUser,$TextBoxSearchMatchUser,$TextBoxSelectedMatchUser,$RadioButtonSelectNewInfo,$RadioButtonMatchExistingUser,$LabelModifyUser,$LabelMatchUser,$ButtonSearchModifyUser,$SearchMatchUser,$ComboBoxNewMainDept,$LabelNewInfo,$LabelCurrentInfo,$ComboBoxNewSubDept,$ComboBoxNewJobTitle,$ButtonCommitChanges,$LabelMainDept,$LabelSubDept,$LabelJobTitle))
$GroupboxCurrentSettings.Controls.AddRange(@($ComboBoxCurrentMainDept,$ComboBoxCurrentSubDept,$ComboBoxCurrentJobTitle))

[void]$Form.ShowDialog()

##########################
##### Event Handlers #####
##########################

$RadioButtonSelectNewInfo_CheckedChanged = {
    if ($RadioButtonSelectNewInfo.Checked = $true){
        $RadioButtonMatchExistingUser.Checked = $false
    }
}

$RadioButtonMatchExistingUser_CheckedChanged = {
    <#if ($RadioButtonMatchExistingUser.Checked = $true){
        $RadioButtonSelectNewInfo.Checked = $false
    }#>
    if ($RadioButtonSelectNewInfo.Checked = $true){
        $RadioButtonMatchExistingUser.Checked = $false
    }
}