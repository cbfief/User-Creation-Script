# Installs AD modules
import-module activedirectory

# Enables popup text boxes
Add-Type -AssemblyName Microsoft.VisualBasic;

$Attempts = 12

# Write-host "Please complete the following questions, Ensure spelling and case are accurate"

$Credential = Get-Credential

$First = [Microsoft.VisualBasic.Interaction]::InputBox('Enter First Name', 'XA Group', '')

$Last = [Microsoft.VisualBasic.Interaction]::InputBox('Enter Last Name', 'XA Group', '')

# Finds Main Department
$Mains = Get-ADOrganizationalUnit -SearchBase 'OU=Garden City,DC=GardenCityks,DC=US' -Filter {description -like "D - *"} | select name,distinguishedname | sort name

$MainGrid = @{
    OutputMode = 'Single'
    Title = 'Select Main Department and click OK'
}

$Main = $Mains | Out-GridView @MainGrid | Foreach {
    $_.distinguishedname
}

# Runs to see if there is another layer of departments
$SubFind = Get-ADOrganizationalUnit -Searchbase $main -SearchScope OneLevel -Filter {Name -like "Us*"} | select name,distinguishedname | sort name

If ($SubFind){
    $subs = 'null'

    #Finds Job Title and stores it for use
    $Titles = Get-ADGroup -Searchbase $main -Filter {description -like "RBAC*"} | select name,distinguishedname | sort name

    $TitleGrid = @{
        OutputMode = 'Single'
        Title = 'Select Job Title and click OK'
    }

    $TitleD = $Titles | Out-GridView @MainGrid | Foreach {
        $_.distinguishedname
    }
    # Gets name of RBAC group for Account Title
    $title = Get-ADGroup -Identity $TitleD | Select name

    [String]$Path = Get-ADOrganizationalUnit -SearchBase $Main -Filter {Name -like "US*"}
}
Else{
    $subs = Get-ADOrganizationalUnit -SearchBase $main -SearchScope OneLevel -Filter {Description -like "*"} | select name,distinguishedname | sort name

    $SubGrid = @{
        Outputmode = 'Single'
        Title = 'Select Sub Department and click OK'
    }

    $Sub = $Subs | Out-GridView @SubGrid | Foreach {
        $_.Distinguishedname
    }

    #Finds Job Title and stores it for use
    $Titles = Get-ADGroup -Searchbase $Sub -Filter {description -like "RBAC*"} | select name,distinguishedname | sort name

    $TitleGrid = @{
        OutputMode = 'Single'
        Title = 'Select Job Title and click OK'
    }

    $TitleD = $Titles | Out-GridView @MainGrid | Foreach {
        $_.distinguishedname
    }
    # Gets name of RBAC group for Account Title
    $title = Get-ADGroup -Identity $TitleD | Select name

    [String]$Path = Get-ADOrganizationalUnit -SearchBase $sub -Filter {Name -like "US*"}
}

#Pre-set fields generic to all users regardless of location

$FirstLower=$First.ToLower()
$LastLower=$Last.ToLower()

$SAMAccountName=$FirstLower+'.'+$LastLower

$DisplayName=$First+' '+$Last

$Mailnickname=$First+$Last

$UserPrincipalName=$FirstLower+'.'+$LastLower+'@gardencityks.us'

$RemoteRoutingAddress=$FirstLower+'.'+$LastLower+'@gardencityksus.mail.onmicrosoft.com'

$ProxyEmailAddress=$FirstLower+'.'+$LastLower+'@gardencityksus.mail.onmicrosoft.com'

$EmailAddress=$FirstLower+'.'+$LastLower+'@gardencityks.us'

$password=Read-Host "Enter Users Password" -AsSecureString

#Create user section - this builds the AD account using the fields above
New-ADUser -SAMAccountName $SAMAccountName -name $DisplayName -GivenName $First -Surname $Last -UserPrincipalName $UserPrincipalName -AccountPassword $password -DisplayName $DisplayName -Path $Path -ChangePasswordAtLogon $True -Enabled $True -OtherAttributes @{mail=$EmailAddress}

#This section adds the users email addresses. The primary email address should be SMTP in caps, secondary addresses in lowercase.
#Set-ADUser -identity $SAMAccountName -Add @{proxyAddresses="SMTP:$EmailAddress"}
Set-ADUser -identity $SAMAccountName -Add @{proxyAddresses="smtp:$ProxyEmailAddress"}

#pauses the script to allow AD to replicate
start-sleep -milliseconds 5000

#Set-ADAccountPassword -identity $SAMAccountName -NewPassword $password -Reset
Start-sleep -milliseconds 5000

#This section forces and AD to 365 Delta sync from the domain controller, then pauses the script to make sure the sync has completed.
Invoke-Command -Computer ITC-sv12ad01 -Scriptblock {Start-ADSyncSyncCycle -PolicyType Delta}
start-sleep -milliseconds 10000

#This part of the script connects to a Powershell session via the on-prem exchange 2016 server (hybrid environment).
$HybridSession = New-PSSession –ConfigurationName Microsoft.Exchange –ConnectionUri http://ITC-sv16EHVM/powershell -Authentication Kerberos 

Import-PSSession $HybridSession -DisableNameChecking -AllowClobber

#This part creates the Office365 mailbox though the on-premise exchange server (hybrid mode)
Enable-RemoteMailbox -identity $SAMAccountName –remoteroutingaddress $RemoteRoutingAddress

#This bit turns on mailbox archiving - check your licencing arrangement!
#Enable-RemoteMailbox $SAMAccountName -Archive

#Forces the script to pause whilst 365 account is setup
start-sleep -milliseconds 20000
Write-host "Pausing while the 365 accoutn is setup"

#Ends hybrid session
Remove-PSSession $HybridSession

#Connects to Office 365 portal. Will prompt for valid admin credentials. Manually running $AccountSKU Will report back number of licences used / available.
import-module MsOnline
Connect-MsolService -Credential $Credential
$AccountSKU = Get-MsolAccountSKU
$AccountSKU
$ExitTimer = 0
while ( -not (Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction SilentlyContinue)){
    Start-Sleep -s 5
    $ExitTimer++
    if ($ExitTimer -ge $Attempts) {
        [System.Windows.MessageBox]::Show("$UserPrincipalName not found in O365.  Exiting.")
        exit
        }
    }
$UserLicence = Get-MsolUser -UserPrincipalName $UserPrincipalName


#This sets the users location; needed before licences can be assigned
 Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation US

Write-host "Assigning licence: Office 365 G3"
 
 Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses "gardencityksus:ENTERPRISEPACK_GOV"

#start-sleep -milliseconds 5000

#$Credential = Get-Credential

$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid" -Credential $credential -Authentication "Basic" -AllowRedirection

Import-PSSession $ExchangeSession -CommandName Get-Mailbox, Set-Mailbox

#start-sleep -milliseconds 10000

$ExitTimer = 0
while ( -Not (Get-Mailbox -identity $SAMAccountName -erroraction silentlycontinue)){
    Start-Sleep -s 5
    $ExitTimer++
    if ($ExitTimer -ge $Attempts) {
        [System.Windows.MessageBox]::Show("Mailbox for $SAMAccountName not found.  Exiting.")
        Exit
        }
    }
Get-Mailbox -identity $SAMAccountName | Set-Mailbox -LitigationHoldEnabled $True

start-sleep -milliseconds 5000

#Cleans up connection to 365 servers
Remove-PSSession $ExchangeSession

write-host "Allow 30 minutes for Microsoft / Office 365 to create the mailbox"

start-sleep -milliseconds 10000
exit