<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '800,800'
$Form.text                       = "Form"
$Form.TopMost                    = $false

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "button"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(50,750)
$Button1.Font                    = 'Microsoft Sans Serif,10'

$TabControl                     = New-object System.Windows.Forms.TabControl
$TabControl.Width               = 800
$TabControl.Height              = 710
$TabControl.Location            = New-Object System.Drawing.Point(0,20)

$NewUserTab                     = New-Object System.Windows.Forms.TabPage
$NewUserTab.text                = "New User"

$DeleteUserTab                  = New-Object System.Windows.Forms.TabPage
$DeleteUserTab.Text             = "Delete User"

$ModifyUserTab                  = New-Object System.Windows.Forms.TabPage
$ModifyUserTab.Text             = "Modify User"

$TBfirstName                     = New-Object system.Windows.Forms.TextBox
$TBfirstName.multiline           = $false
$TBfirstName.width               = 143
$TBfirstName.height              = 20
$TBfirstName.location            = New-Object System.Drawing.Point(120,68)
$TBfirstName.Font                = 'Microsoft Sans Serif,10'

$TBlastName                      = New-Object system.Windows.Forms.TextBox
$TBlastName.multiline            = $false
$TBlastName.width                = 143
$TBlastName.height               = 20
$TBlastName.location             = New-Object System.Drawing.Point(121,105)
$TBlastName.Font                 = 'Microsoft Sans Serif,10'

$CBDepartment                       = New-Object system.Windows.Forms.ComboBox
$CBDepartment.text                  = "Select Department"
$CBDepartment.width                 = 179
$CBDepartment.height                = 20
$CBDepartment.location              = New-Object System.Drawing.Point(37,145)
$CBDepartment.Font                  = 'Microsoft Sans Serif,10'

$LBFname                           = New-Object system.Windows.Forms.Label
$LBFname.text                      = "First Name"
$LBFname.AutoSize                  = $true
$LBFname.width                     = 25
$LBFname.height                    = 10
$LBFname.location                  = New-Object System.Drawing.Point(37,70)
$LBFname.Font                      = 'Microsoft Sans Serif,10'

$LBLname                           = New-Object system.Windows.Forms.Label
$LBLname.text                      = "Last Name"
$LBLname.AutoSize                  = $true
$LBLname.width                     = 13
$LBLname.height                    = 10
$LBLname.location                  = New-Object System.Drawing.Point(36,110)
$LBLname.Font                      = 'Microsoft Sans Serif,10'

$CBSubDepartment                       = New-Object system.Windows.Forms.ComboBox
$CBSubDepartment.text                  = "Select Sub Department"
$CBSubDepartment.width                 = 181
$CBSubDepartment.height                = 20
$CBSubDepartment.location              = New-Object System.Drawing.Point(36,181)
$CBSubDepartment.Font                  = 'Microsoft Sans Serif,10'

$CBJobTitle                        = New-Object system.Windows.Forms.ComboBox
$CBJobTitle.text                   = "Select Job Title"
$CBJobTitle.width                  = 260
$CBJobTitle.height                 = 20
$CBJobTitle.location               = New-Object System.Drawing.Point(36,219)
$CBJobTitle.Font                   = 'Microsoft Sans Serif,10'

$CkBcreateEmail                  = New-Object system.Windows.Forms.CheckBox
$CkBcreateEmail.text             = "Create Email Account"
$CkBcreateEmail.AutoSize         = $false
$CkBcreateEmail.width            = 179
$CkBcreateEmail.height           = 20
$CkBcreateEmail.location         = New-Object System.Drawing.Point(36,260)
$CkBcreateEmail.Font             = 'Microsoft Sans Serif,10'

$BTcreateAccount                 = New-Object system.Windows.Forms.Button
$BTcreateAccount.text            = "Create Account"
$BTcreateAccount.width           = 147
$BTcreateAccount.height          = 30
$BTcreateAccount.location        = New-Object System.Drawing.Point(36,320)
$BTcreateAccount.Font            = 'Microsoft Sans Serif,10'
$BTcreateAccount.ForeColor       = ""

$InfoBox                         = New-Object system.Windows.Forms.TextBox
$InfoBox.multiline               = $true
$InfoBox.width                   = 586
$InfoBox.height                  = 220
$InfoBox.location                = New-Object System.Drawing.Point(36,380)
$InfoBox.Font                    = 'Microsoft Sans Serif,10'
$InfoBox.ReadOnly                = $true
$InfoBox.ScrollBars              = "Vertical"

$Form.controls.AddRange(@($Button1,$TabControl))
$tabControl.Controls.AddRange(@($NewUserTab,$DeleteUserTab,$ModifyUserTab))
$NewUserTab.Controls.AddRange(@($TBfirstName,$TBlastName,$CBDepartment,$LBFname,$LBLname,$CBSubDepartment,$CBJobTitle,$CkBcreateEmail,$BTcreateAccount,$InfoBox))


[void]$Form.ShowDialog()