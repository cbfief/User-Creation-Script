Add-Type -AssemblyName System.Windows.Forms
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '661,757'
$Form.text                       = "Form"
$Form.TopMost                    = $false

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
$BTcreateAccount.location        = New-Object System.Drawing.Point(36,381)
$BTcreateAccount.Font            = 'Microsoft Sans Serif,10'
$BTcreateAccount.ForeColor       = ""

$InfoBox                         = New-Object system.Windows.Forms.TextBox
$InfoBox.multiline               = $true
$InfoBox.width                   = 586
$InfoBox.height                  = 220
$InfoBox.location                = New-Object System.Drawing.Point(37,466)
$InfoBox.Font                    = 'Microsoft Sans Serif,10'
$InfoBox.ReadOnly                = $true
$InfoBox.ScrollBars              = "Vertical"

$Form.controls.AddRange(@($TBfirstName,$TBlastName,$CBDepartment,$LBFname,$LBLname,$CBSubDepartment,$CBJobTitle,$CkBcreateEmail,$BTcreateAccount,$InfoBox))


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


# Global Variables
$global:SubDepts
$global:Titles
$global:Title
$global:TitleDN
$global:Path = $null
$Global:Attempts = 12


# Finds Main Department
$Mains = Get-ADOrganizationalUnit -SearchBase 'OU=Garden City,DC=GardenCityks,DC=US' -Filter {description -like "D - *"} | select name,distinguishedname | sort name

foreach($dept in $Mains){
    $CBDepartment.items.Add($dept.name)
    }


$CBDepartment_SelectedIndexChanged = {
    $MainDept = $Mains[$CBDepartment.SelectedIndex].distinguishedname
    $SubFind = Get-ADOrganizationalUnit -Searchbase $MainDept -SearchScope OneLevel -Filter {Name -like "Us*"} | select name,distinguishedname | sort name
    $CBSubDepartment.items.Clear()
    $CBSubDepartment.Text = "Select Sub Department"
    clearTitle
    $Global:Path = $null

    If ($SubFind){
        $SubDepts = 'null'
    
        $Global:Titles = Get-ADGroup -Searchbase $MainDept -Filter {description -like "RBAC*"} | select name,distinguishedname | sort name
    
        foreach($job in $Titles){
            $CBJobTitle.items.Add($job.name)
            }
    
        [String]$Global:Path = Get-ADOrganizationalUnit -SearchBase $MainDept -Filter {Name -like "US*"}
    }

    Else{
        $Global:SubDepts = Get-ADOrganizationalUnit -SearchBase $MainDept -SearchScope OneLevel -Filter {Description -like "*"} | select name,distinguishedname | sort name

        foreach($dept in $SubDepts){
            $CBSubDepartment.items.Add($dept.name)
            }
        }
}

$CBSubDepartment_SelectedIndexChanged = {
    clearTitle

    $SubDept = $SubDepts[$CBSubDepartment.SelectedIndex].distinguishedname

    $Global:Titles = Get-ADGroup -Searchbase $SubDept -Filter {description -like "RBAC*"} | select name,distinguishedname | sort name

    foreach($job in $Titles){
        $CBJobTitle.items.Add($job.name)
        }

    [String]$Global:Path = Get-ADOrganizationalUnit -SearchBase $SubDept -Filter {Name -like "US*"}
}

$CBJobTitle_SelectedIndexChanged = {
        $global:TitleDN = $Global:Titles[$CBJobTitle.SelectedIndex].distinguishedname
        $global:Title = $Global:Titles[$CBJobTitle.SelectedIndex].name
}

$CBDepartment.add_selectedindexchanged($CBDepartment_SelectedIndexChanged)
$CBSubDepartment.add_selectedindexchanged($CBSubDepartment_SelectedIndexChanged)
$CBJobTitle.add_selectedindexchanged($CBJobTitle_SelectedIndexChanged)

$BTcreateAccount.Add_Click({createAD})


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

    $UserPrincipalName = $FirstLower + '.' + $LastLower + '@gardencityks.us'

    $RemoteRoutingAddress = $FirstLower + '.' + $LastLower + '@gardencityksus.mail.onmicrosoft.com'

    $ProxyEmailAddress = $FirstLower + '.' + $LastLower + '@gardencityksus.mail.onmicrosoft.com'

    $EmailAddress = $FirstLower + '.' + $LastLower + '@gardencityks.us'

    # Make sure both first and last names were given
    if (!$FirstName -or !$LastName){
        $InfoBox.AppendText("Please enter the user's name `r`n`r`n")
        Return
    }
    # Make sure the AD path is specified
    if (!$global:Path){
        $InfoBox.Text = "Please select a department"
    }
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
        assignO365License $SAMAccountName $UserPrincipalName
    }
}
    
function enableMailbox($SAMAccountName, $RemoteRoutingAddress) {
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
    $global:Title = $null
    $global:TitleDN = $null
}

function Main{
[System.Windows.Forms.Application]::EnableVisualStyles()
[void]$Form.ShowDialog()
} Main