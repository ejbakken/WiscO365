#Module variables
$moduleDataPath = "$env:LOCALAPPDATA\WindowsPowerShell\ModuleData\WiscO365"
$credPath = "$moduleDataPath\Connections"
$connectionsFile = "$moduleDataPath\Connections\connections.csv"

$defaultEndpoint = "https://wiscmail.wisc.edu/domainadmin.json"

Function Invoke-HelperO365Function ($body){
    Try{
        #Invoke the JSON command
        $result = Invoke-RestMethod -Method Default -Uri $Global:O365CurrentConnection.endpoint -WebSession $Global:O365Session -Credential $Global:O365CurrentConnection.cred -Body $body
        #Check if the command was successful
        If(!$result.mesg){return $result.result}
        ElseIf($result.mesg -like "Success*") {return $result.result}
        Else{Write-Error $result.mesg}
    }
    Catch{
        #If an exception resulted, write the error
        ($_.ErrorDetails.Message | ConvertFrom-Json).error.result |
        Foreach{if($_ -eq 'authcheck failed'){Write-Error $_} Else {Write-Error $_}}
    }
}

Function Test-HelperO365Connection (){
    If(!$Global:O365CurrentConnection){
        Write-Error "The connection to the endpoint is not set. Please set the connection using the method Set-O365Connection."
        break
    }
}

Function Create-HelperO365ModuleData () {
    #Check if path exists
    If(!(Test-Path -Path "$moduleDataPath\Connections")){
        Try{
            New-Item "$moduleDataPath\Connections" -type directory -Force -ErrorAction Stop | Out-Null
        }
        Catch{
            Write-Error $_
            break
        }
    }

    #Check if path exists
    If(!(Test-Path -Path "$moduleDataPath\Old API Functions")){
        Try{
            New-Item "$moduleDataPath\Old API Functions" -type directory -Force -ErrorAction Stop | Out-Null
        }
        Catch{
            Write-Error $_
            break
        }
    }

    #Check if path exists
    If(!(Test-Path -Path "$moduleDataPath\Preferences")){
        Try{
            New-Item "$moduleDataPath\Preferences" -type directory -Force -ErrorAction Stop | Out-Null
        }
        Catch{
            Write-Error $_
            break
        }
    }

        #Check if path exists
    If(!(Test-Path -Path "$moduleDataPath\Preferences\Startup.txt")){
        Try{
            New-Item "$moduleDataPath\Preferences\Startup.txt" -type file -Force -ErrorAction Stop | Out-Null
            1 | Out-File -FilePath "$moduleDataPath\Preferences\Startup.txt" -Force -ErrorAction Stop
        }
        Catch{
            Write-Error $_
            break
        }
    }
}

Function Show-HelperO365StartupMenu (){
    $startupPath = "$moduleDataPath\Preferences\Startup.txt"
    
    Try{
        $startupPreference = Get-Content $startupPath -ErrorAction SilentlyContinue
    }
    Catch{
        Write-Error $_
    }

    If($startupPreference -like "*1*"){
        Write-Host "`n"
        Write-Host "`t Welcome to the WiscO365 PowerShell Module!" -ForegroundColor Cyan
        Write-Host "`n"
        Write-Host "`t`t To get information and help on this module, enter Get-O365Help." -ForegroundColor White
        Write-Host "`t`t To turn off this menu in future sessions, enter Set-O365StartupPreferences." -ForegroundColor White
        Write-Host "`n"
    }
}

Function Update-HelperO365InitialHelp () {
    Invoke-Expression "$PSScriptRoot\Update\psDoc-master\src\psDoc.ps1 -moduleName WiscO365 -outputDir $PSScriptRoot -fileName 'WiscO365 Initial Help.html'"
}

Create-HelperO365ModuleData
Show-HelperO365StartupMenu

#Check for the API Functions file
    If(Test-Path -Path "$moduleDataPath\APIFunctions.ps1"){
        #If it exists, import the API Functions
        . "$moduleDataPath\APIFunctions.ps1"
    }
    Else{
        #If it doesn't exist, import the function to get the other API functions
        . "$PSScriptRoot\Update\Get-O365DomainAdminDoc.ps1"
    }