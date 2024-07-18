# Checks if the AzureAD module is installed an imported
# Check if a current connection to AzureAD exists
<#if(!(Get-Module -Name "AzureAD")){
    Install-Module AzureAD
    Import-Module AzureAD
    Connect-AzureAD
}#>

$user = Read-Host "Enter the users email address"
$userObj
# Loop to validate the format of the email string
while ($user -notmatch '^[a-zA-Z0-9._%±]+@[a-zA-Z0-9.-]+.[a-zA-Z]{2,}$') {
    Write-Host '*********** ' $user is not a valid email format!' **********' -ForegroundColor Red
    $user = Read-Host "Enter the users email address"
    
    if(null -eq (Get-AzureADUser -ObjectID $user) ){
        Write-Host $user ' does not exist!' -ForegroundColor Red
    }
}

# Create an array of compulsary groups
$groups = @('sec-azure-zpa-all-users', 'sec-azure-miro-users', 'AutoPilot Users (Apps)')

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
        if($groupToAdd.DisplayName -eq 'AutoPilot Users (Apps)'){
            Write-Host "Assigning E5 license via " $groupToAdd.DisplayName -ForegroundColor Green
        }else{
            Write-Host "Adding user to " $groupToAdd.DisplayName -ForegroundColor Green
            # Add-AzureADGroupMember -ObjectId $groupToAdd.ObjectId -RefObjectId (Get-AzureADUser -ObjectId $user).ObjectId
        }
    }catch{
        Write-Host $user " is already a member of " $group -ForegroundColor Red
    }
}

# Assign MS Vivia Insights license
Write-Host "Assigning Microsoft Viva Insights License" -ForegroundColor Green
$licenseToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$licenseToAssign.SkuId = '3d957427-ecdc-4df2-aacd-01cc9d519da8'
$licenseGroup = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$licenseGroup.AddLicenses = $licenseToAssign
Set-AzureADUserLicense -ObjectId (Get-AzureADUser -ObjectId $user).ObjectId -AssignedLicenses $licenseGroup

# Call AddToSafeSenders function
AddToSafeSenders -userUPN $user

function AddToSafeSenders {
    param ([string] $userUPN)

    # list the user as a trusted sender
    Write-Host "Adding " $userUPN "to safe senders"
    $All = Get-Mailbox $userUPN;
    $All | ForEach-Object {
        Set-MailboxJunkEmailConfiguration $_.Name -TrustedSendersAndDomains @{
            Add="matt.halliday@ampol.com.au","sdm@ampol.com.au","communications@ampol.com.au","brent.merrick@ampol.com.au"}
    }   
}

Connect-ExchangeOnline
$email = read-host "Type user's email address and press enter"
$All = Get-Mailbox $email;
$All | ForEach-Object {
    Set-MailboxJunkEmailConfiguration $_.Name -TrustedSendersAndDomains @{Add="matt.halliday@ampol.com.au","sdm@ampol.com.au","communications@ampol.com.au","brent.merrick@ampol.com.au"}
}