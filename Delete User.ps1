<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '671,681'
$Form.text                       = "Form"
$Form.TopMost                    = $false

$TBName                          = New-Object system.Windows.Forms.TextBox
$TBName.multiline                = $false
$TBName.width                    = 269
$TBName.height                   = 20
$TBName.location                 = New-Object System.Drawing.Point(55,59)
$TBName.Font                     = 'Microsoft Sans Serif,10'

$LBNameSearch                    = New-Object system.Windows.Forms.Label
$LBNameSearch.text               = "Enter part of the User`'s name:"
$LBNameSearch.AutoSize           = $true
$LBNameSearch.width              = 25
$LBNameSearch.height             = 10
$LBNameSearch.location           = New-Object System.Drawing.Point(56,30)
$LBNameSearch.Font               = 'Microsoft Sans Serif,12,style=Bold'

$BTSearchUsers                   = New-Object system.Windows.Forms.Button
$BTSearchUsers.text              = "Search"
$BTSearchUsers.width             = 60
$BTSearchUsers.height            = 30
$BTSearchUsers.location          = New-Object System.Drawing.Point(358,49)
$BTSearchUsers.Font              = 'Microsoft Sans Serif,10,style=Bold'

$ListBSearchResults              = New-Object system.Windows.Forms.ListBox
$ListBSearchResults.text         = "listBox"
$ListBSearchResults.width        = 364
$ListBSearchResults.height       = 211
$ListBSearchResults.location     = New-Object System.Drawing.Point(56,108)

$BTSelectUser                    = New-Object system.Windows.Forms.Button
$BTSelectUser.text               = "Select"
$BTSelectUser.width              = 60
$BTSelectUser.height             = 30
$BTSelectUser.location           = New-Object System.Drawing.Point(436,248)
$BTSelectUser.Font               = 'Microsoft Sans Serif,10,style=Bold'

$Form.controls.AddRange(@($TBName,$LBNameSearch,$BTSearchUsers,$ListBSearchResults,$BTSelectUser))

$BTSearchUsers.Add_Click($BTSearchUsers_Click)
$BTSelectUser.Add_Click($BTSelectUser_Click)

Import-Module ActiveDirectory

[void]$Form.ShowDialog()


$BTSearchUsers_Click = {
    $ListBSearchResults.Items.Clear()
    $ListBSearchResults
    $NameSearch = $TBName.Text
    $NameSearch = "*" + $NameSearch + "*"
    $FoundUsers = Get-ADUser -Filter {Name -like $NameSearch}
    foreach ($user in $FoundUsers){
        $ListBSearchResults.Items.Add($user.name)
    }
}