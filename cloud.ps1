# Checks if the AzureAD module is installed an imported
# Check if a current connection to AzureAD exists


function AssignLicense {
    Param ([string] $skuID)
    $licenseToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $licenseToAssign.SkuId = $skuID
    $licenseGroup = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $licenseGroup.AddLicenses = $licenseToAssign
    Set-AzureADUserLicense -ObjectId (Get-AzureADUser -ObjectId $user).ObjectId -AssignedLicenses $licenseGroup
}
function AddToSafeSenders {
    param ([string] $userUPN)

    if(!(Get-Module -Name "ExchangeOnlineManagement")){
        Install-Module ExchangeOnlineManagement
        Import-Module ExchangeOnlineManagement
    }

    Write-Host "Connecting to ExchangeOnline"
    Connect-ExchangeOnline

    # list the user as a trusted sender
    Write-Host "Adding " $userUPN "to safe senders"
    $All = Get-Mailbox $userUPN;
    $All | ForEach-Object {
        Set-MailboxJunkEmailConfiguration $_.Name -TrustedSendersAndDomains @{
            Add="matt.halliday@ampol.com.au","sdm@ampol.com.au","communications@ampol.com.au","brent.merrick@ampol.com.au"}
    }   

    Read-Host "Completed Successfully! Please press any key to close this window." -ForegroundColor Green
}

if(!(Get-Module -Name "AzureAD")){
    Install-Module AzureAD
    Import-Module AzureAD
    Connect-AzureAD
}

$userUPN
function ValidateUser {
    param ([string] $userUPN)
        try{
            Get-AzureADUser -ObjectID $userUPN
        }catch{
            return $false
        }
    return $true 
}

while(1){
    while(1){
        try{
            $userUPN = Read-Host "Enter the users email address"
            if($userUPN -match '^[a-zA-Z0-9._%±]+@[a-zA-Z0-9.-]+.[a-zA-Z]{2,}$'){
                if(!(ValidateUser -userUPN $userUPN)){
                    Write-Host $userUPN ' does not exist!' -ForegroundColor Red
                }else{
                    Write-Host $userUPN ' exists!' -ForegroundColor Green         
                }              
            }else{
                Write-Host '*********** ' $userUPN is not a valid email format!' **********' -ForegroundColor Red
            }
            break
        }catch{
            Write-Host 'Something went wrong!'
        }
    }
}


# Create an array of compulsary groups
$groups = @('sec-azure-zpa-all-users', 'sec-azure-miro-users', 'AutoPilot Users (Apps)', 'sec-azure-SSPR-Enable')

# Prompt user to dertermine the correct DoneSafe group
# Loop until the user selects a valid number
while(1){
   try {
        # Validates the users input is an integer
        $doneSafe = [int](Read-Host "Please choose from one of the following:`n1: The user reports to the CEO.`n2: The user has direct reports.`n3: None of the above.")
        
        # Checks if the input matches exactly '1', '2' or '3'
        if($doneSafe -match '\b[1-3]\b'){break}
        else{Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
    }
    catch {Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
}

# Check the value of $doneSafe and add it to the $groups Array
if ($doneSafe -eq 1) {$groups += "DoneSafe Z Executives"}
elseif($doneSafe -eq 2){$groups += "DoneSafe People Leaders"}
elseif($doneSafe -eq 3){$groups += 'DoneSafe Leaders of Self'}


<#
    TODO: 
        1. verify the groups have been successfully added by checking the users group memberships.
#>
foreach($group in $groups){
    $groupToAdd = Get-AzureADGroup -SearchString $group

    try{
            Write-Host "Adding user to " $groupToAdd.DisplayName -ForegroundColor Green
            Add-AzureADGroupMember -ObjectId $groupToAdd.ObjectId -RefObjectId (Get-AzureADUser -ObjectId $userUPN).ObjectId
        }
    catch{
        Write-Host $userUPN " is already a member of " $group -ForegroundColor Red
    }
}

# Assign MS Vivia Insights license
Write-Host "Assigning Microsoft Viva Insights License" -ForegroundColor Green
AssignLicense -skuID '3d957427-ecdc-4df2-aacd-01cc9d519da8'

Write-Host "Assigning M365 E5 License" -ForegroundColor Green
AssignLicense -skuID '06ebc4ee-1bb5-47dd-8120-11324bc54e06'

# Call AddToSafeSenders function
AddToSafeSenders -userUPN $userUPN

Read-Host -Prompt "Completed successfully! Press Enter to exit"

