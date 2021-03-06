﻿Add-Type -AssemblyName System.Windows.Forms
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '1080,720'
$Form.text                       = "Form"
$Form.TopMost                    = $false

$TBfirstName                     = New-Object system.Windows.Forms.TextBox
$TBfirstName.multiline           = $false
$TBfirstName.width               = 150
$TBfirstName.height              = 20
$TBfirstName.location            = New-Object System.Drawing.Point(314,58)
$TBfirstName.Font                = 'Microsoft Sans Serif,10'

$TBlastName                      = New-Object system.Windows.Forms.TextBox
$TBlastName.multiline            = $false
$TBlastName.width                = 143
$TBlastName.height               = 20
$TBlastName.location             = New-Object System.Drawing.Point(621,58)
$TBlastName.Font                 = 'Microsoft Sans Serif,10'

$CBDepartment                       = New-Object system.Windows.Forms.ComboBox
$CBDepartment.text                  = "Select Department"
$CBDepartment.width                 = 175
$CBDepartment.height                = 20
$CBDepartment.location              = New-Object System.Drawing.Point(313,100)
$CBDepartment.Font                  = 'Microsoft Sans Serif,10'

$LBFname                           = New-Object system.Windows.Forms.Label
$LBFname.text                      = "First Name"
$LBFname.AutoSize                  = $true
$LBFname.width                     = 25
$LBFname.height                    = 10
$LBFname.location                  = New-Object System.Drawing.Point(231,62)
$LBFname.Font                      = 'Microsoft Sans Serif,10'
$LBFname.ForeColor               = "Firebrick"

$LBLname                           = New-Object system.Windows.Forms.Label
$LBLname.text                      = "Last Name"
$LBLname.AutoSize                  = $true
$LBLname.width                     = 13
$LBLname.height                    = 10
$LBLname.location                  = New-Object System.Drawing.Point(536,62)
$LBLname.Font                      = 'Microsoft Sans Serif,10'
$LBLname.ForeColor               = "Firebrick"

$CBSubDepartment                       = New-Object system.Windows.Forms.ComboBox
$CBSubDepartment.text                  = "" #"Select Sub Department"
$CBSubDepartment.width                 = 175
$CBSubDepartment.height                = 20
$CBSubDepartment.location              = New-Object System.Drawing.Point(622,100)
$CBSubDepartment.Font                  = 'Microsoft Sans Serif,10'

$CBJobTitle                        = New-Object system.Windows.Forms.ComboBox
$CBJobTitle.text                   = "Select Job Title"
$CBJobTitle.width                  = 275
$CBJobTitle.height                 = 20
$CBJobTitle.location               = New-Object System.Drawing.Point(313,142)
$CBJobTitle.Font                   = 'Microsoft Sans Serif,10'

$CkBcreateEmail                  = New-Object system.Windows.Forms.CheckBox
$CkBcreateEmail.text             = "Email"
$CkBcreateEmail.AutoSize         = $false
$CkBcreateEmail.width            = 95
$CkBcreateEmail.height           = 20
$CkBcreateEmail.location         = New-Object System.Drawing.Point(315,228)
$CkBcreateEmail.Font             = 'Microsoft Sans Serif,10'

$CkBaccountEnabled               = New-Object system.Windows.Forms.CheckBox
$CkBaccountEnabled.text          = "Enable Account"
$CkBaccountEnabled.AutoSize      = $true
$CkBaccountEnabled.width         = 95
$CkBaccountEnabled.height        = 20
$CkBaccountEnabled.location      = New-Object System.Drawing.Point(410,228)
$CkBaccountEnabled.Font          = 'Microsoft Sans Serif,10'

$CkBchangePassword               = New-Object system.Windows.Forms.CheckBox
$CkBchangePassword.text          = "Password Must be Changed"
$CkBchangePassword.AutoSize      = $true
$CkBchangePassword.width         = 95
$CkBchangePassword.height        = 20
$CkBchangePassword.location      = New-Object System.Drawing.Point(563,228)
$CkBchangePassword.Font          = 'Microsoft Sans Serif,10'

$BTcreateAccount                 = New-Object system.Windows.Forms.Button
$BTcreateAccount.text            = "Create Account"
$BTcreateAccount.width           = 150
$BTcreateAccount.height          = 30
$BTcreateAccount.location        = New-Object System.Drawing.Point(413,264)
$BTcreateAccount.Font            = 'Microsoft Sans Serif,10'
$BTcreateAccount.ForeColor       = ""
$BTcreateAccount.Enabled         = $false

$InfoBox                         = New-Object system.Windows.Forms.TextBox
$InfoBox.multiline               = $true
$InfoBox.width                   = 1000
$InfoBox.height                  = 275
$InfoBox.location                = New-Object System.Drawing.Point(40,380)
$InfoBox.Font                    = 'Microsoft Sans Serif,10'
$InfoBox.ReadOnly                = $true
$InfoBox.ScrollBars              = "Vertical"

$Form.controls.AddRange(@($TBfirstName,$TBlastName,$CBDepartment,$LBFname,$LBLname,$CBSubDepartment,$CBJobTitle,$CkBcreateEmail,$CkBaccountEnabled,$CkBchangePassword,$BTcreateAccount,$InfoBox))


$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Password"
$objForm.Size = New-Object System.Drawing.Size(300,200) 
$objForm.StartPosition = "CenterScreen"
$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$x=$objTextBox.Text;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.Add_Click({$x=$objTextBox.Text;$objForm.Close()})
$objForm.Controls.Add($OKButton)
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(280,20) 
$objLabel.Text = "Please enter the new account password:"
$objForm.Controls.Add($objLabel) 
$MaskedTextBox = New-Object System.Windows.Forms.MaskedTextBox
$MaskedTextBox.PasswordChar = '*'
$MaskedTextBox.Location = New-Object System.Drawing.Size(10,40) 
$MaskedTextBox.Size = New-Object System.Drawing.Size(260,20) 
$objForm.Controls.Add($MaskedTextBox) 
$objForm.Topmost = $True
$objForm.Add_Shown({$objForm.Activate()})

$CBDepartment.add_selectedindexchanged($CBDepartment_SelectedIndexChanged)
$CBSubDepartment.add_selectedindexchanged($CBSubDepartment_SelectedIndexChanged)
$CBJobTitle.add_selectedindexchanged($CBJobTitle_SelectedIndexChanged)
$BTcreateAccount.Add_Click($BTcreateAccount_Click)
$InfoBox.Add_TextChanged($InfoBox_TextChanged)

$TBfirstName.Add_TextChanged({ 
    if($TBfirstName.TextLength -gt 0)
        { $LBFname.ForeColor = "Black"}
    else
        { $LBFname.ForeColor = "Firebrick" }
    evaluateInfoComplete
})
$TBlastName.Add_TextChanged({ 
    if($TBlastName.TextLength -gt 0)
        { $LBLname.ForeColor = "Black"}
    else
        { $LBLname.ForeColor = "Firebrick" }
    evaluateInfoComplete
})

$CBDepartment.DisplayMember = 'Name'
$CBSubDepartment.DisplayMember = 'Name'
$CBJobTitle.DisplayMember = 'Name'

##########################################
##### Global Variables and Functions #####
##########################################
$Global:Attempts = 12
$Global:Domain = "@gardencityks.us"
$Global:ProxyDomain = "@gardencityksus.mail.onmicrosoft.com"

#region Get-DateSortable
function Get-datesortable
{
	$global:datesortable = Get-Date -Format "HH':'mm':'ss"
	return $global:datesortable
}#endregion Get-DateSortable

#region Add-Logs
function Add-Logs
{
	[CmdletBinding()]
	param ($text)
	Get-datesortable
    $InfoBox.Text += "[$global:datesortable] - $text`r`n"
	Set-Alias alogs Add-Logs -Description "Add content to the TextBoxLogs"
	Set-Alias Add-Log Add-Logs -Description "Add content to the TextBoxLogs"
}

##### End Global Variables and functions #####
##############################################

# Finds Main Department
$Mains = Get-ADOrganizationalUnit -SearchBase 'OU=Garden City,DC=GardenCityks,DC=US' -Filter {description -like "D - *"} | select name,distinguishedname | sort name


foreach($dept in $Mains){
    $CBDepartment.items.Add($dept)
    }


$CBDepartment_SelectedIndexChanged = {
    #$MainDept = $Mains[$CBDepartment.SelectedIndex].distinguishedname
    $MainDept = $CBDepartment.SelectedItem.distinguishedname
    $SubFind = Get-ADOrganizationalUnit -Searchbase $MainDept -SearchScope OneLevel -Filter {Name -like "Us*"} | select name,distinguishedname | sort name
    $CBSubDepartment.items.Clear()
    #$CBSubDepartment.Text = "Select Sub Department"
    clearTitle
    evaluateInfoComplete
    If ($SubFind){
        $SubDepts = 'null'
    
        $Titles = Get-ADGroup -Searchbase $MainDept -Filter {description -like "RBAC*"} | select name,distinguishedname | sort name
    
        foreach($job in $Titles){
            $CBJobTitle.items.Add($job)
            }
        $CBSubDepartment.Text = ""
    }

    Else{
        $SubDepts = Get-ADOrganizationalUnit -SearchBase $MainDept -SearchScope OneLevel -Filter {Description -like "*"} | select name,distinguishedname | sort name
        foreach($dept in $SubDepts){
            $CBSubDepartment.items.Add($dept)
            }
        $CBSubDepartment.Text = "Select Sub Department"
        }
}

$CBSubDepartment_SelectedIndexChanged = {
    clearTitle
    $MainDept = $CBDepartment.SelectedItem.distinguishedname
    $SubDept = $CBSubDepartment.SelectedItem.distinguishedname

    $Titles = Get-ADGroup -Searchbase $SubDept -Filter {description -like "RBAC*"} | select name,distinguishedname | sort name
    #Write-Host "JobTitle selected index: " $CBJobTitle.SelectedIndex
    foreach($job in $Titles){
        $CBJobTitle.items.Add($job)
        }
}

$CBJobTitle_SelectedIndexChanged = {
    evaluateInfoComplete
}

$BTcreateAccount_Click = {
    $MainDept = $CBDepartment.SelectedItem.distinguishedname

    if ($CBSubDepartment.items.Count -gt 0) {
        $SubDept = $CBSubDepartment.SelectedItem.distinguishedname
        [String]$Path = Get-ADOrganizationalUnit -SearchBase $SubDept -Filter {Name -like "US*"}
    }
    else {
    [String]$Path = Get-ADOrganizationalUnit -SearchBase $MainDept -Filter {Name -like "US*"}
    }
 
    $Title = $CBJobTitle.SelectedItem.name
    $TitleDN = $CBJobTitle.SelectedItem.distinguishedname
 
    $FirstName = $TBfirstName.Text -replace '\s',''
    $LastName = $TBlastName.Text -replace '\s',''
    Add-Logs -text "LOG: Main Department is: "
}

$InfoBox_TextChanged = {
	$InfoBox.SelectionStart = $InfoBox.TextLength;
	$InfoBox.ScrollToCaret()
	$InfoBox.Focus()
		
	If ($Global:ExternalLog -ne $null)
	{
		$InfoBox.Text | Out-File $Global:ExternalLog 
    }
}



# Function - Create Active Directory account
function createAD {
    # Get User Names and use it to set account variables
    $FirstName = $TBfirstName.Text -replace '\s',''
    $FirstLower = $FirstName.ToLower()
    $LastName = $TBlastName.Text -replace '\s',''
    $LastLower = $LastName.ToLower()

    $SAMAccountName = $FirstLower + '.' + $LastLower

    $DisplayName = $FirstName + ' ' + $LastName

    $MailNickName = $FirstName+$LastName

    $UserPrincipalName = $FirstLower + '.' + $LastLower + $Global:Domain

    $RemoteRoutingAddress = $FirstLower + '.' + $LastLower + $Global:ProxyDomain

    $ProxyEmailAddress = $FirstLower + '.' + $LastLower + $Global:ProxyDomain

    $EmailAddress = $FirstLower + '.' + $LastLower + $Global:Domain

    # Prompt for the new user's password
    [void] $objForm.ShowDialog()
    $password = $MaskedTextBox.Text
    # If it was blank or cancelled, stop the process
    if (!$password){
        $InfoBox.AppendText("User Creation cancelled `r`n`r`n")
        Return
    }
    $password = ConvertTo-SecureString -String $password -AsPlainText -Force

    #Create user section - this builds the AD account using the fields above
    # If email checkbox was not selected, only create an AD account
    if (!$CkBcreateEmail.Checked) {
        $InfoBox.AppendText("Account will be created with the following options: `r`n
            SAM Account: $SAMAccountName `r`n
            Display Name: $DisplayName `r`n
            Title: $global:Title `r`n
            UPN: $UserPrincipalName `r`n
            Email: None `r`n
            OU: $Global:Path `r`n`r`n")
        New-ADUser -SAMAccountName $SAMAccountName -name $DisplayName -GivenName $FirstName -Surname $LastName -UserPrincipalName $UserPrincipalName -AccountPassword $password -DisplayName $DisplayName -Path $Path -ChangePasswordAtLogon $True -Enabled $True -OtherAttributes @{title=$Global:Title}
        
        Start-Sleep -s 2
        $InfoBox.AppendText("AD account $SAMAccountName created `r`n`r`n")
        }
    # If email was requested we do some extra stuff to set up O365
    else {
        $InfoBox.AppendText("Account will be created with the following options: `r`n
            SAM Account: $SAMAccountName `r`n
            Display Name: $DisplayName `r`n
            Title: $global:Title `r`n
            UPN: $UserPrincipalName `r`n
            Email: $EmailAddress `r`n
            OU: $Global:Path `r`n`r`n")
        #New-ADUser -SAMAccountName $SAMAccountName -name $DisplayName -GivenName $FirstName -Surname $LastName -UserPrincipalName $UserPrincipalName -AccountPassword $password -DisplayName $DisplayName -Path $Path -ChangePasswordAtLogon $True -Enabled $True -OtherAttributes @{title=$Global:Title;mail=$EmailAddress}
        
        #Set-ADUser -identity $SAMAccountName -Add @{proxyAddresses="smtp:$ProxyEmailAddress"}
        
        #pauses the script to allow AD to replicate
        start-sleep -milliseconds 5000
        
        $InfoBox.AppendText("AD account $SAMAccountName created `r`n`r`n")

        #This section forces and AD to 365 Delta sync from the domain controller, then pauses the script to make sure the sync has completed.
        #Invoke-Command -Computer ITC-sv12ad01 -Scriptblock {Start-ADSyncSyncCycle -PolicyType Delta}
        start-sleep -milliseconds 10000
        
        # Call the function to set up the O365 account
        #$MailboxCode = enableMailbox $SAMAccountName $RemoteRoutingAddress
        #if ($MailboxCode -eq 1){
        #    Return
        #}
        
    }
}
    
function enableMailbox {
    #This part of the script connects to a Powershell session via the on-prem exchange 2016 server (hybrid environment).
    $HybridSession = New-PSSession –ConfigurationName Microsoft.Exchange –ConnectionUri http://ITC-sv16EHVM/powershell -Authentication Kerberos 

    Import-PSSession $HybridSession -DisableNameChecking -AllowClobber

    #This part creates the Office365 mailbox though the on-premise exchange server (hybrid mode)
    $ExitTimer = 0
    while ( -not (Get-ADUser -identity $SAMAccountName)) {
        Start-Sleep -Seconds 5
        $ExitTimer++
        if ($ExitTimer -ge $Global:Attempts){
            $InfoBox.AppendText("User not found. `r`n`r`n")
            Return 1
        }
    }
    Enable-RemoteMailbox -identity $SAMAccountName –remoteroutingaddress $RemoteRoutingAddress

    #This bit turns on mailbox archiving - check your licencing arrangement!
    #Enable-RemoteMailbox $SAMAccountName -Archive

    #Forces the script to pause whilst 365 account is setup
    #start-sleep -milliseconds 20000
    #Write-host "Pausing while the 365 account is setup"

    #Ends hybrid session
    Remove-PSSession $HybridSession
    Return 0
}

function assignO365License($SAMAccountName, $UserPrincipalName) {
    # Ask for credentials to use for connecting to O365
    $Credential = Get-Credential -UserName "$env:USERNAME@$env:USERDNSDOMAIN" -Message "Enter Credentials"

    if (!$Credential){
        $InfoBox.AppendText("Invalid Credentials.  Unable to assign O365 license.")
        Return
    }

    #Connects to Office 365 portal. Will prompt for valid admin credentials. Manually running $AccountSKU Will report back number of licences used / available.
    import-module MsOnline
    Connect-MsolService -Credential $Credential
    $AccountSKU = Get-MsolAccountSKU
    $AccountSKU
    $ExitTimer = 0
    while ( -not (Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction SilentlyContinue)){
        $InfoBox.AppendText("Working... `r`n")
        Start-Sleep -s 5
        $ExitTimer++
        if ($ExitTimer -ge $Global:Attempts) {
            $InfoBox.AppendText("$UserPrincipalName not found in O365. `r`n`r`n")
            Return
            }
        }
    $UserLicence = Get-MsolUser -UserPrincipalName $UserPrincipalName


    #This sets the users location; needed before licences can be assigned
    Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation US

    #Write-host "Assigning licence: Office 365 G3"
    $InfoBox.AppendText("`r`n" + "Assigning licence: Office 365 G3 `r`n`r`n")

    Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses "gardencityksus:ENTERPRISEPACK_GOV"

    #start-sleep -milliseconds 5000

    #$Credential = Get-Credential

    $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid" -Credential $credential -Authentication "Basic" -AllowRedirection

    Import-PSSession $ExchangeSession -CommandName Get-Mailbox, Set-Mailbox

    #start-sleep -milliseconds 10000

    $ExitTimer = 0
    while ( -Not (Get-Mailbox -identity $SAMAccountName -erroraction silentlycontinue)){
        $InfoBox.AppendText("Working... `r`n")
        Start-Sleep -s 5
        $ExitTimer++
        if ($ExitTimer -ge $Global:Attempts) {
            $InfoBox.AppendText("Mailbox for $SAMAccountName not found. `r`n`r`n")
            Return
            }
        }
    Get-Mailbox -identity $SAMAccountName | Set-Mailbox -LitigationHoldEnabled $True

    #start-sleep -milliseconds 5000

    #Cleans up connection to 365 servers
    Remove-PSSession $ExchangeSession

    #write-host "Allow 30 minutes for Microsoft / Office 365 to create the mailbox"
    $InfoBox.AppendText("`r`nFinished!  Allow 30 minutes for Microsoft / Office 365 to create the mailbox `r`n")
}

function clearTitle {
    $CBJobTitle.items.Clear()
    $CBJobTitle.Text = "Select Job Title"
    $CBJobTitle.SelectedIndex = -1
    #$global:Title = $null
    #$global:TitleDN = $null
}

function evaluateInfoComplete {
    if ($CBDepartment.SelectedIndex -ge 0 -and $CBJobTitle.SelectedIndex -ge 0 -and $TBFirstName.TextLength -gt 0 -and $TBLastName.TextLength -gt 0) {
        $BTcreateAccount.Enabled = $true
        }
    else {
        $BTcreateAccount.Enabled = $false
        }
}

function Main{
[System.Windows.Forms.Application]::EnableVisualStyles()
[void]$Form.ShowDialog()
} Main