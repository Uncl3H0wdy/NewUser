# Connect to MsolService to assign data location
Import-Module MSOnline
Install-Module MSOnline

# Checks if the AzureAD module is installed an imported
# Check if a current connection to AzureAD exists
if(!(Get-Module -Name "AzureAD")){
    Write-Host "Installing and importing the AzureAD module" -ForegroundColor Yellow
    try{Install-Module AzureAD}
    catch{Write-Host "Could not install AzureAD module. Please try again." -ForegroundColor Red}
    try{Import-Module AzureAD}
    catch{Write-Host "Could not import AzureAD module. Please try again." -ForegroundColor Red}
    Write-Host "AzureAD module has installed imported successfully" -ForegroundColor Green
}

try{
    Write-Host "Connecting to AzureAD - please see the login prompt" -ForegroundColor Yellow
    Connect-AzureAD -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
    Write-Host "Connected to AzureAD" -ForegroundColor Green
}catch{
    Write-Host "Could not connect to AzureAD. Please try again." -ForegroundColor Red
    exit
}

function ValidateRole{

    # Checks the signed in user has the correct roles for running this script.
    Write-Host "Validating your RBAC roles before proceeding with this script" -ForegroundColor Yellow
    $userRole
    $flag = $false
    $roles = @('760908a9-a770-47dd-aa87-139ea74b1897', 'fd18ac2d-fb67-48d2-b9f8-a6417acfcb25', '40713a15-ad64-4b62-9522-ad49ab2b3f9e')
    $currentUser = (Get-AzureADUser -ObjectID (Get-AzureADCurrentSessionInfo).Account.Id)
    while ($flag -eq $false) {
        foreach($role in $roles){
            $userRole = Get-AzureADDirectoryRoleMember -ObjectId $role | Where-Object {$_.UserPrincipalName -eq $currentUser.UserPrincipalName}
    
            if($currentUser -eq $userRole){
                $flag = $true
                Write-Host "Roles have been validated!" -ForegroundColor Green
                break
            }
            Read-Host -Prompt "Do do not have the required RBAC roles. Please assign either, Global Admin, Exchange Admin or Exchange Recipient Admin then press Enter to re-validate." -ForegroundColor Red
        }
    }
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

function ValidateMailBox {
    param ([string] $mbxToValidate)
    $timer = 0
    while($true){
        try{
            Get-EXOMailbox -Identity $mbxToValidate -ErrorAction Stop
            Write-Host "The mailbox has been validated" -ForegroundColor Green
            break
        }catch{
            if($timer -eq 10){
                Write-Error "Could not find the mailbox"
                break
            }
            Start-Sleep -Seconds 5
            $timer += 1
        }
    }
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
        Write-Host "Installing and importing the Exchange Online Management module" -ForegroundColor Yellow
        try{Install-Module ExchangeOnlineManagement}
        catch{Write-Host "Could not install Exchange Online Management module. Please try again." -ForegroundColor Red}
        try{Import-Module ExchangeOnlineManagement}
        catch{Write-Host "Could not import Exchange Online Management module. Please try again." -ForegroundColor Red}
        Write-Host "Exchange Online Management module has installed imported successfully" -ForegroundColor Green
    }
    try{
        Write-Host "Connecting to ExchangeOnline - please see the login prompt" -ForeGroundColor Yellow
        Connect-ExchangeOnline -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
        Write-Host "Connected to ExchangeOnline" -ForeGroundColor Green
    }catch{
        Write-Host "Could not connect to ExchangeOnline. Please try again." -ForegroundColor Red
        exit
    }

    Write-Host "Validating configuration changes before proceeding. Please wait." -ForegroundColor Yellow
    ValidateMailBox -mbxToValidate $userUPN 
    
    # Mark the defined users as "not spam" on the target users mailbox
    Write-Host "Adding "$userUPN "to safe senders" -ForegroundColor Yellow

    $userMailbox = Get-Mailbox $userUPN;
    $userMailbox | ForEach-Object {
        Set-MailboxJunkEmailConfiguration $_.Name -TrustedSendersAndDomains @{
            Add="matt.halliday@ampol.com.au","sdm@ampol.com.au","communications@ampol.com.au","brent.merrick@ampol.com.au"}
    }   
    Write-Host "Successfully added "$userUPN" to safe senders" -ForegroundColor Green
}

function ValidateGroups {
    param([string] $groupToValidate)
    try {Get-AzureADGroup -SearchString $groupToValidate}
    catch {return $false}
    return $true
}

ValidateRole

$userObject
$userUPN
$distributionLists = @('DLAllUsers@z.co.nz')
$groups = @('sec-azure-zpa-all-users', 'sec-azure-miro-users', 'AutoPilot Users (Apps)', 'sec-azure-SSPR-Enable')

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
                break  
            }                    
        }else{Write-Host '*********** ' $userUPN is not a valid email format!' **********' -ForegroundColor Red}         
    }catch{Write-Host $_}
}

try{
    Write-Host "Connecting to MsolService - please see the login prompt" -ForegroundColor Yellow
    Connect-MsolService -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
    Write-Host "Connected to MsolService" -ForegroundColor Green
}catch{
    Write-Host "Could not connect to MsolService. Please try again." -ForegroundColor Red
    exit
}

Write-Host "Setting usage location to New Zealand"

$timer = 0
while($true){
    try{
        Get-MsolUser -userprincipalname $userUPN | Set-MsolUser -UsageLocation NZ
        Write-Host "The usage location has been set to NZ" -ForegroundColor Green
        break
    }catch{
        if($timer -eq 10){
            Write-Error "Could not set the usage location"
            break
        }
        Start-Sleep -Seconds 5
        $timer += 1
    }
}

try{
    Write-Host "Connecting to AzureAD - please see the login prompt" -ForegroundColor Yellow
    Connect-AzureAD -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
    Write-Host "Connected to AzureAD" -ForegroundColor Green
}catch{
    Write-Host "Could not connect to AzureAD. Please try again." -ForegroundColor Red
    exit
}

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
        Add-AzureADGroupMember -ObjectId $groupToAdd.ObjectId -RefObjectId (Get-AzureADUser -ObjectId $userUPN).ObjectId
        Write-Host "Successfully added" $userObject.UserPrincipalName "to" $groupToAdd.DisplayName -ForegroundColor Green
        }catch{
            if($_.Exception.Message.Contains("One or more added object references already exist for the following modified properties: 'members'")){
                Write-Host $userObject.UserPrincipalName "is already a member of" $groupToAdd.DisplayName -ForeGroundColor Red
            }
        }
    }

# Assign MS Vivia Insights license
if(!(ValidateLicense -skuID '3d957427-ecdc-4df2-aacd-01cc9d519da8')){
    Write-Host "The License does not exist!" -ForeGroundColor Red
}else{
    AssignLicense -skuID '3d957427-ecdc-4df2-aacd-01cc9d519da8'
    Write-Host "Successfully assigned the Microsoft Viva Insights License" -ForegroundColor Green
}

if(!(ValidateLicense -skuID '06ebc4ee-1bb5-47dd-8120-11324bc54e06')){
    Write-Host "The License does not exist!" -ForeGroundColor Red
}else{
    AssignLicense -skuID '06ebc4ee-1bb5-47dd-8120-11324bc54e06'
    Write-Host "Successfully assigned the M365 E5 License" -ForegroundColor Green
}

# Call AddToSafeSenders function
AddToSafeSenders -userUPN $userObject.UserPrincipalName


#Prompt user to dertermine the correct Distribution List
# Loop until the user selects a valid number
while($true){
    try {
         # Validates the users input is an integer
         $selectedDL = [int](Read-Host "Please choose which Distribution List to add the user too:`n1: DL WEL Users (Wellington)`n2: DL CHC Users (Christchurch).`n3: DL Te Whare Rama (Auckland).")
         
         # Checks if the input matches exactly '1', '2' or '3'
         if($selectedDL -match '\b[1-3]\b'){
             # Check the value of $doneSafe and add it to the $groups Array
             if ($selectedDL -eq 1) {$distributionLists += "DLwelusers@z.co.nz"}
             elseif($selectedDL -eq 2){$distributionLists += 'DLchcusers@z.co.nz'}
             elseif($selectedDL -eq 3){$distributionLists += "DLTeWhareRama@z.co.nz"}
             break
         }
         else{Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
        }
     catch {Write-Host '*********** Please choose an option from 1 - 3 **********' -ForegroundColor Red}
 }

 foreach($item in $distributionLists){
    $dlToAdd = Get-DistributionGroup -Identity $item
    try {
        Add-DistributionGroupMember -Identity $item -Member $userObject.UserPrincipalName -ErrorAction Stop
        Write-Host "Successfully added" $userObject.DisplayName "to" $dlToAdd.DisplayName -ForegroundColor Green
    }
    catch {
        Write-Host  $userObject.DisplayName "is already a member of" $dlToAdd.DisplayName -ForegroundColor Red
    }
}



Read-Host -Prompt "Completed successfully! Press Enter to exit"
Exit

# $userObject.DisplayName"<"$userObject.UserPrincipalName">" "is already a member of "$dlToAdd.DisplayName"<"$dlToAdd.PrimarySmtpAddress">"