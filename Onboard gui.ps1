Add-Type -AssemblyName System.Windows.Forms

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

$ComboBox1                       = New-Object system.Windows.Forms.ComboBox
$ComboBox1.text                  = "Select Department"
$ComboBox1.width                 = 179
$ComboBox1.height                = 20
$ComboBox1.location              = New-Object System.Drawing.Point(37,145)
$ComboBox1.Font                  = 'Microsoft Sans Serif,10'

$Fname                           = New-Object system.Windows.Forms.Label
$Fname.text                      = "First Name"
$Fname.AutoSize                  = $true
$Fname.width                     = 25
$Fname.height                    = 10
$Fname.location                  = New-Object System.Drawing.Point(37,70)
$Fname.Font                      = 'Microsoft Sans Serif,10'

$Lname                           = New-Object system.Windows.Forms.Label
$Lname.text                      = "Last Name"
$Lname.AutoSize                  = $true
$Lname.width                     = 13
$Lname.height                    = 10
$Lname.location                  = New-Object System.Drawing.Point(36,110)
$Lname.Font                      = 'Microsoft Sans Serif,10'

$ComboBox2                       = New-Object system.Windows.Forms.ComboBox
$ComboBox2.text                  = "Select Sub Department"
$ComboBox2.width                 = 181
$ComboBox2.height                = 20
$ComboBox2.location              = New-Object System.Drawing.Point(36,181)
$ComboBox2.Font                  = 'Microsoft Sans Serif,10'

$JobTitle                        = New-Object system.Windows.Forms.ComboBox
$JobTitle.text                   = "Select Job Title"
$JobTitle.width                  = 260
$JobTitle.height                 = 20
$JobTitle.location               = New-Object System.Drawing.Point(36,219)
$JobTitle.Font                   = 'Microsoft Sans Serif,10'

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

$Form.controls.AddRange(@($TBfirstName,$TBlastName,$ComboBox1,$Fname,$Lname,$ComboBox2,$JobTitle,$CkBcreateEmail,$BTcreateAccount,$InfoBox))

# Global Variables
$global:SubDepts
$global:Titles
$global:Title
$global:TitleDN
$global:Path = $null


# Finds Main Department
$Mains = Get-ADOrganizationalUnit -SearchBase 'OU=Garden City,DC=GardenCityks,DC=US' -Filter {description -like "D - *"} | select name,distinguishedname | sort name

foreach($dept in $Mains){
    $ComboBox1.items.Add($dept.name)
    }


$ComboBox1_SelectedIndexChanged = {
    $MainDept = $Mains[$ComboBox1.SelectedIndex].distinguishedname
    $SubFind = Get-ADOrganizationalUnit -Searchbase $MainDept -SearchScope OneLevel -Filter {Name -like "Us*"} | select name,distinguishedname | sort name
    $ComboBox2.items.Clear()
    $ComboBox2.Text = "Select Sub Department"
    clearTitle
    $Global:Path = $null

    If ($SubFind){
        $SubDepts = 'null'
    
        $Global:Titles = Get-ADGroup -Searchbase $MainDept -Filter {description -like "RBAC*"} | select name,distinguishedname | sort name
    
        foreach($job in $Titles){
            $JobTitle.items.Add($job.name)
            }
    
        [String]$Global:Path = Get-ADOrganizationalUnit -SearchBase $MainDept -Filter {Name -like "US*"}
    }

    Else{
        $Global:SubDepts = Get-ADOrganizationalUnit -SearchBase $MainDept -SearchScope OneLevel -Filter {Description -like "*"} | select name,distinguishedname | sort name

        foreach($dept in $SubDepts){
            $ComboBox2.items.Add($dept.name)
            }
        }
}

$ComboBox2_SelectedIndexChanged = {
    clearTitle

    $SubDept = $SubDepts[$ComboBox2.SelectedIndex].distinguishedname

    $Global:Titles = Get-ADGroup -Searchbase $SubDept -Filter {description -like "RBAC*"} | select name,distinguishedname | sort name

    foreach($job in $Titles){
        $JobTitle.items.Add($job.name)
        }

    [String]$Global:Path = Get-ADOrganizationalUnit -SearchBase $SubDept -Filter {Name -like "US*"}
}

$JobTitle_SelectedIndexChanged = {
        $global:TitleDN = $Global:Titles[$JobTitle.SelectedIndex].distinguishedname
        $global:Title = $Global:Titles[$JobTitle.SelectedIndex].name
}

$ComboBox1.add_selectedindexchanged($ComboBox1_SelectedIndexChanged)
$ComboBox2.add_selectedindexchanged($ComboBox2_SelectedIndexChanged)
$JobTitle.add_selectedindexchanged($JobTitle_SelectedIndexChanged)

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

    $UserPrincipalName = $FirstLower + '.' + $LastLower + 'gardencityks.us'

    $RemoteRoutingAddress = $FirstLower + '.' + $LastLower + '@gardencityksus.mail.onmicrosoft.com'

    $ProxyEmailAddress = $FirstLower + '.' + $LastLower + '@gardencityksus.mail.onmicrosoft.com'

    $EmailAddress = $FirstLower + '.' + $LastLower + '@gardencityks.us'

    #$password = Read-Host "Enter Users Password" -AsSecureString

    #Create user section - this builds the AD account using the fields above
    # make sure the AD path is specified
    if ($global:Path){
        # If email checkbox was not selected, only create an AD account
        if (!$CkBcreateEmail.Checked) {
            $InfoBox.Text = "Account will be created with the following options: `r`n
                SAM Account: $SAMAccountName
                Display Name: $DisplayName
                Title: $global:Title
                UPN: $UserPrincipalName
                Email: None
                OU: $Global:Path `r`n`r`n"
            $InfoBox.text += "New-ADUser -SAMAccountName $SAMAccountName -name $DisplayName -GivenName $FirstName -Surname $LastName -UserPrincipalName $UserPrincipalName -AccountPassword $password -DisplayName $DisplayName -Path $Path -ChangePasswordAtLogon $True -Enabled $True -OtherAttributes @{title=$Global:Title}"
            }
        # If email was requested we do some extra stuff to set up O365
        else {
            $InfoBox.Text = "Account will be created with the following options: `r`n
                SAM Account: $SAMAccountName
                Display Name: $DisplayName
                Title: $global:Title
                UPN: $UserPrincipalName
                Email: $EmailAddress
                OU: $Global:Path `r`n`r`n"
            $infoBox.Text += "New-ADUser -SAMAccountName $SAMAccountName -name $DisplayName -GivenName $FirstName -Surname $LastName -UserPrincipalName $UserPrincipalName -AccountPassword $password -DisplayName $DisplayName -Path $Path -ChangePasswordAtLogon $True -Enabled $True -OtherAttributes @{title=$Global:Title;mail=$EmailAddress}"
            
            #Set-ADUser -identity $SAMAccountName -Add @{proxyAddresses="smtp:$ProxyEmailAddress"}
            
            #pauses the script to allow AD to replicate
            #start-sleep -milliseconds 5000

            #This section forces and AD to 365 Delta sync from the domain controller, then pauses the script to make sure the sync has completed.
            #Invoke-Command -Computer ITC-sv12ad01 -Scriptblock {Start-ADSyncSyncCycle -PolicyType Delta}
            #start-sleep -milliseconds 10000
            
            # Call the function to set up the O365 account
            enableMailbox $SAMAccountName $RemoteRoutingAddress
        }
    }
    else {
        $InfoBox.Text = "Please select a department"
    }
}
    
function enableMailbox($SAMAccountName, $RemoteRoutingAddress) {
    #This part of the script connects to a Powershell session via the on-prem exchange 2016 server (hybrid environment).
    #$HybridSession = New-PSSession –ConfigurationName Microsoft.Exchange –ConnectionUri http://ITC-sv16EHVM/powershell -Authentication Kerberos 

    #Import-PSSession $HybridSession -DisableNameChecking -AllowClobber

    #This part creates the Office365 mailbox though the on-premise exchange server (hybrid mode)
    $InfoBox.Text += "`r`n`r`n" + "Enable-RemoteMailbox -identity $SAMAccountName –remoteroutingaddress $RemoteRoutingAddress"

    #This bit turns on mailbox archiving - check your licencing arrangement!
    #Enable-RemoteMailbox $SAMAccountName -Archive

    #Forces the script to pause whilst 365 account is setup
    #start-sleep -milliseconds 20000
    #Write-host "Pausing while the 365 account is setup"

    #Ends hybrid session
    #Remove-PSSession $HybridSession
}

function clearTitle {
    $JobTitle.items.Clear()
    $JobTitle.Text = "Select Job Title"
    $global:Title = $null
    $global:TitleDN = $null
}
    

function Main{
[System.Windows.Forms.Application]::EnableVisualStyles()
[void]$Form.ShowDialog()
} Main