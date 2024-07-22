# Checks if the AzureAD module is installed an imported
# Check if a current connection to AzureAD exists
if(!(Get-Module -Name "AzureAD")){
    Install-Module AzureAD
    Import-Module AzureAD
    Connect-AzureAD
}
if(!(Get-Module -Name "Microsoft.Graph.Identity.DirectoryManagement")){
    Install-Module Microsoft.Graph.Identity.DirectoryManagement
    Import-Module Microsoft.Graph.Identity.DirectoryManagement
}
function ValidateLicense {
    Param ([string] $skuID){
        try{Get-MgSubscribedSku -SubscribedSkuId $skuID}
        catch{return $false}
        return $true
    }
}

function ValidateUser {
    param ([string] $userUPN)
    try{Get-AzureADUser -ObjectID $userUPN}
    catch{return $false}
    return $true 
}

function AssignLicense {
    Param ([string] $skuID)
    $licenseToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $licenseToAssign.SkuId = $skuID
    $licenseGroup = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $licenseGroup.AddLicenses = $licenseToAssign
    Set-AzureADUserLicense -ObjectId (Get-AzureADUser -ObjectId $userUPN).ObjectId -AssignedLicenses $licenseGroup
}
function AddToSafeSenders {
    param ([string] $userUPN)
    if(!(Get-Module -Name "ExchangeOnlineManagement")){
        Install-Module ExchangeOnlineManagement
        Import-Module ExchangeOnlineManagement
    }

    Write-Host "Connecting to ExchangeOnline"
    Connect-ExchangeOnline

    # Mark the defined users as "not spam" on the target users mailbox
    Write-Host "Adding "$userUPN "to safe senders" -ForegroundColor Yellow
    $userMailbox = Get-Mailbox $userUPN;
    $userMailbox | ForEach-Object {
        Set-MailboxJunkEmailConfiguration $_.Name -TrustedSendersAndDomains @{
            Add="matt.halliday@ampol.com.au","sdm@ampol.com.au","communications@ampol.com.au","brent.merrick@ampol.com.au"}
    }   
    Write-Host "Successfully added "$userUPN" to safe senders" -ForegroundColor Green
}

$userObject
$userUPN


function ValidateGroups {
    param([string] $groupToValidate)
    try {Get-AzureADGroup -SearchString $groupToValidate}
    catch {return $false}
    return $true
}

while($true){
    try{
        $userUPN = Read-Host "Enter the users email address"
        if($userUPN -match '^[a-zA-Z0-9._%±]+@[a-zA-Z0-9.-]+.[a-zA-Z]{2,}$'){
            Write-Host "Fetching the user Object....." -ForegroundColor DarkYellow
            if(!(ValidateUser -userUPN $userUPN)){
                Write-Host $userUPN ' does not exist!' -ForegroundColor Red
            }else{
                Write-Host "Successfully fetched the User Object and ready to proceed!" -ForegroundColor Green
                $userObject = Get-AzureADUser -ObjectID $userUPN                            
            }
            break                    
        }else{Write-Host 'Invalid email format: ' $userUPN -ForegroundColor Red}         
    }catch{Write-Host $_}
}


# Create an array of compulsary groups
$groups = @('sec-azure-zpa-all-users', 'sec-azure-miro-users', 'AutoPilot Users (Apps)', 'sec-azure-SSPR-Enable')

# Prompt user to dertermine the correct DoneSafe group
# Loop until the user selects a valid number
while($true){
    try {
         # Validates the users input is an integer
         $doneSafe = [int](Read-Host "Please choose from one of the following:`n1: The user reports to the CEO.`n2: The user has direct reports.`n3: None of the above.")
         
         # Checks if the input matches exactly '1', '2' or '3'
         if($doneSafe -match '\b[1-3]\b'){
             # Check the value of $doneSafe and add it to the $groups Array
             if ($doneSafe -eq 1) {$groups += "DoneSafe Z Executives"}
             elseif($doneSafe -eq 2){$groups += "DoneSafe People Leaders"}
             elseif($doneSafe -eq 3){$groups += 'DoneSafe Leaders of Self'}
             break
         }
         else{Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
     }
     catch {Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
 }

foreach($group in $groups){
    $groupToAdd = Get-AzureADGroup -SearchString $group
    try{
        Write-Host "Adding user to" $groupToAdd.DisplayName -ForegroundColor Green
        Add-AzureADGroupMember -ObjectId $groupToAdd.ObjectId -RefObjectId (Get-AzureADUser -ObjectId $userUPN).ObjectId
        Write-Host "Successfully added" $userUPN "to" $groupToAdd.DisplayName -ForegroundColor Green
        }catch{
            Write-Host $userUPN "is already a member of" $group -ForegroundColor Red
        }
    }

# Assign MS Vivia Insights license
Write-Host "Validating Microsoft Viva Insights License" -ForegroundColor Yellow
if(!(ValidateLicense -skuID '3d957427-ecdc-4df2-aacd-01cc9d519da8')){
    Write-Host "The License does not exist!" -ForeGroundColor Red
}else{
    Write-Host "Assigning Microsoft Viva Insights License" -ForegroundColor Green
    AssignLicense -skuID '3d957427-ecdc-4df2-aacd-01cc9d519da8'
    Write-Host "Successfully assigned the Microsoft Viva Insights License" -ForegroundColor Green
}

Write-Host "Validating M365 E5 License" -ForegroundColor Yellow
if(!(ValidateLicense -skuID '06ebc4ee-1bb5-47dd-8120-11324bc54e06')){
    Write-Host "The License does not exist!" -ForeGroundColor Red
}else{
    Write-Host "Assigning M365 E5 License" -ForegroundColor Green
    AssignLicense -skuID '06ebc4ee-1bb5-47dd-8120-11324bc54e06'
    Write-Host "Successfully assigned the M365 E5 License" -ForegroundColor Green
}


# Call AddToSafeSenders function
AddToSafeSenders -userUPN $userUPN


#Prompt user to dertermine the correct DoneSafe group
# Loop until the user selects a valid number

$distributionLists = @('DLAllUsers@z.co.nz')

while($true){
    try {
         # Validates the users input is an integer
         $DL = [int](Read-Host "Please choose which Distribution List to add the user too:`n1: DL WEL Users (Wellington)`n2: DL CHC Users (Christchurch).`n3: DL Te Whare Rama (Auckland).")
         
         # Checks if the input matches exactly '1', '2' or '3'
         if($DL -match '\b[1-3]\b'){
             # Check the value of $doneSafe and add it to the $groups Array
             if ($DL -eq 1) {$distributionLists += "DLwelusers@z.co.nz"}
             elseif($DL -eq 2){$distributionLists += 'DLchcusers@z.co.nz'}
             elseif($DL -eq 3){$distributionLists += "DLTeWhareRama@z.co.nz"}
             break
         }
         else{Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
     }
     catch {Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
 }

 foreach($dl in $distributionLists){
    $dlToAdd = Get-DistributionGroup -Identity $dl
    Write-Host "Adding user to" $dlToAdd.DisplayName -ForegroundColor Green
    try {
        Add-DistributionGroupMember -Identity $dl -Member $userObject.UserPrincipalName -ErrorAction Stop
        Write-Host "Successfully added" + $userObject.DisplayName "to" $dlToAdd.DisplayName -ForegroundColor Green
    }
    catch {
        Write-Host  $userObject.DisplayName "is already a member of" $dlToAdd.DisplayName -ForegroundColor Red
    }
}

Read-Host -Prompt "Completed successfully! Press Enter to exit"

# $userObject.DisplayName"<"$userObject.UserPrincipalName">" "is already a member of "$dlToAdd.DisplayName"<"$dlToAdd.PrimarySmtpAddress">"