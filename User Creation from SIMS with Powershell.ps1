# Very basic UI to let the user provide a data file manually, or let SIMS produce one with Command Reporter
do{
$answer = Read-Host "Would you like for SIMS to 1. generate a data file, or 2. provide one?"
}while (($answer -ne "1") -and ($answer -ne "2"))

if ($answer -eq 1){
    try{
    # Use Command Reporter to talk to SIMS and retrieve data
    & 'C:\Program Files (x86)\SIMS\SIMS .net\commandReporter.exe' "/user:SIMSAdmin" "/password:password" "/report:$reportname" "/output:C:\SIMSReports\StudentData.csv" "/SERVERNAME:SIMSSERVER\SIMS" "/DATABASENAME:SIMS"
    $import = import-csv "C:\SIMSReports\StudentData.csv"
    }
    catch{
        Write-Host "Error running Command Reporter or reading data file. Exiting in 5 seconds"
        sleep 5
        exit}
    }

if ($answer -eq 2){
    $readfile = Read-Host "Please provide the path to the data file: "
    $readfile = $readfile.Replace("`"","")
    try{
        Test-Path $readfile
        }
    catch{
        Write-Host "Error reading data file. Exiting in 5 seconds"
        sleep 5
        exit}
    $import = Import-Csv $readfile
    }

# Read from AD to get the list of current users to make sure that 
$currentusers = Get-ADUser -SearchBase "OU=Students,OU=Users,OU=School,DC=school,DC=internal" -Properties EmployeeID
# Set variable to be used for user creation
$domain = "@school.internal"

$import | ForEach-Object {
    $ID = $_.SIMS_ID
    if($currentusers -notcontains $ID){ # If the ID isn't found in the current users, start the user creation process
        # Assign variables for user creation
        $forename = $_.forename
        $surname = $_.surname
        $displayname = $forename + " " + $surname
        $phone = $_.phone
        $address = $_.address
        $sex = $_.sex
        $gender = $_.gender

        $OU = "" # Blank this variable just in case
        if($_.year -like "*12*"){$OU = "OU=Intake 21,OU=Students,OU=Users,OU=School,DC=school,DC=internal"} # If in year 12, OU is Intake 21
        if($_.year -like "*13*"){$OU = "OU=Intake 20,OU=Students,OU=Users,OU=School,DC=school,DC=internal"} # If in year 13, OU is Intake 20

        $number = 1 # Start iteration of numbers to make sure unique usernames are used
        $username = ($forename.substring(0,1) + $surname).replace(" ","") # Form username with first forename initial and full surname
        if ($usernametocheck.length -ge 15){
            $username = $username.substring(0,14)} # But if the username is too long, then shorten the surname so that the Active Directory username isn't too long.
        $usernametocheck = $username + $number
        $upn = $usernametocheck + $domain
        $usernameexists = $true # Set a boolean for username checking
        do{ # Start a loop to get a unique username
            try
                {$check = Get-ADUser $usernametocheck} # If a user is found, check next iterated username 
            catch{
                $usernameexists = $false # The loop will stop because this has been changed
                New-ADUser -UserPrincipalName $upn `
                    -SamAccountName $usernametocheck `
                    -Name $displayname `
                    -GivenName $forename `
                    -Surname $surname `
                    -EmployeeID $ID `
                    -Path $OU 
                Set-ADUser -identity $usernametocheck -Replace @{telephonenumber=$phone;streetaddress=$address;extensionAttribute1=$sex;extensionAttribute2=$gender} # Set the extra attributes using extension Attributes for extra fields required 
                }
                $number = $number + 1 # Iterate the number so the loop can try again
                $usernametocheck = $username + $number        
         } while ($usernameexists -eq $true)
    }
    else{# If the user doesn't exist, update their record, based on their ID
        $currentusers | ?{$_.employeeID -like $ID} | Set-ADUser -Replace @{name=$name;givenname=$forename;surname=$surname;telephonenumber=$phone;streetaddress=$address;extensionAttribute1=$sex;extensionAttribute2=$gender}
        }
}
