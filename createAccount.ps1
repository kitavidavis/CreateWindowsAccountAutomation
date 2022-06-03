# Defining a function to convert the xlsx file to csv
$dir = "C:\Users\hp\Desktop\script\" 

$excelFileName = "data"

Function ExportWSToCSV ($excelFileName, $csvLoc)
{
    $excelFile = $dir + $excelFileName + ".xlsx"
    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false
    $wb = $E.Workbooks.Open($excelFile)
    foreach ($ws in $wb.Worksheets)
    {
        $n = $excelFileName
        $ws.SaveAs($csvLoc + $n + ".csv", 6)
    }
    $E.Quit()
}

# For each file in the directory with the xlsx format, convert to CSV using the function above
$ens = Get-ChildItem $dir -filter *.xlsx
foreach($e in $ens)
{
    ExportWSToCSV -excelFileName $e.BaseName -csvLoc $dir
}
#import the csv file
$Data = Import-Csv C:\Users\hp\Desktop\script\data.csv

#add a password column to the csv file
$Data | Select-Object *, @{n="Password"; e={''}} | Export-Csv C:\Users\hp\Desktop\script\finaldata.csv

$Data2 = Import-Csv C:\Users\hp\Desktop\script\finaldata.csv

#creating an identifier
$identifiersArr = @()
foreach ($item in $Data2){
    $firstname = $item.Firstname
    $lastname = $item.Lastname

    $firstnameArr = $firstname.ToCharArray()
    $length = $firstnameArr.count
    $myletters = @()
    for ($i=0; $i -lt $length; $i++){
        $id = $firstnameArr[$i]
        $myletters += $id

        $first = $myletters -join ''
        $identifier = $first + $lastname

        if($identifiersArr.Contains($identifier)){
            $i++ 
        } else {
            $identifiersArr += $identifier
            break
        } 
    }
}
Write-Output $identifiersArr

#creating groups from arrays
$groupsArr = @()
foreach($item in $Data2){
    $groups += $item.Group
}
Write-Output $groupsArr

# expiry date
$expiryArr = @()
foreach($item in $Data2){
    $expiryArr += $item.Expiry
}
Write-Output $expiryArr

# description of account
$descArr = @()
foreach($item in $Data2){
    $descArr += $item.Department
}
Write-Output $descArr

#create valid password to the csv file
$words = @('Sun', 'Strawberry', 'Pencil')
$symbols = @('$', '%', '!')

$passwordsArr = @()
foreach ($item in $Data2) {
    $randomNumber1 = 0, 1,2 | Get-Random
    $randomNumber2 = 0, 1, 2 | Get-Random

    $ctxWord = $words[$randomNumber1]
    $ctxSymbol = $symbols[$randomNumber2]

    $EmployeeNum = $item.Employeenumber

    $password = $ctxWord + $ctxSymbol + $EmployeeNum

    $passwordsArr += $password
}

Write-Output $passwordsArr

#creating accounts
function Create-NewLocalAdmin {
    [CmdletBinding()]
    param (
        [string] $NewLocalAdmin,
        [securestring] $Password
    )    
    begin {
    }    
    process {
        New-LocalUser "$NewLocalAdmin" -Password $Password -FullName "$NewLocalAdmin" -Description "Temporary local admin"
        Write-Verbose "$NewLocalAdmin local user created"
        Add-LocalGroupMember -Group "Administrators" -Member "$NewLocalAdmin"
        Write-Verbose "$NewLocalAdmin added to the local administrator group"
    }    
    end {
    }
}

for($i=0; $i -lt $identifiersArr.count; $i++){
    #initializing local variables
    if($groupsArr[$i] -eq 'Admin'){
        Create-NewLocalAdmin -NewLocalAdmin $identifiersArr[$i] -Password $passwordsArr[$i]
    } else {
        if($expiryArr[$i] -eq ''){
            $expiry = $false
        } else {
            $expiry = $true
        }
        $Localuseraccount = @{
            Name = $identifiersArr[$i]
            Password = ($passwordsArr[$i] | ConvertTo-SecureString -AsPlainText -Force)
            AccountNeverExpires = $expiry
            PasswordNeverExpires = $expiry
            Verbose = $true
         }

         New-LocalUser @Localuseraccount
    }
}



