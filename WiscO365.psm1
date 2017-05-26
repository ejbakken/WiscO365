#Import the helper functions
. "$PSScriptRoot\HelperFunctions.ps1"

<#
.Synopsis
    Creates a new connection for the Office 365 API and sets it as the current connection. 
.DESCRIPTION
    This function allows a user to add connection properties for the API, giving them the option to save the connection for later use. The module first imports the existing connections using Get-O365Connections, then checks if that connection ID already exists. If the ID already exists, an error is thrown and processing stops. Next, a new PSCredential object is created to store the credentials. All the connection information is then added $Global:O365Connections. If the save option is selected, the connection information is also exported to Connections.csv and the PSCredential object is exported to an XML file with a title identical to the specified ID. Both files are stored in %LOCALAPPDATA%\WindowsPowerShell\ModuleData\WiscO365\Connections. A stored password can only be used on the computer and by the user account on which it was originally created. Finally, this connection is set as the current connection using Set-O365Connection. 
.OUTPUTS
    connections.csv: Connection details for saved connections.
	$id.xml: Encrypted password file for saved connection.
    $Global:O365Connections: An array of all connections, both saved and temporary.
#>
Function Add-O365Connection{
    [CmdletBinding()]
    Param
    (
        # The unique id for the connection. This can be whatever the user prefers.
        [Parameter(Mandatory=$true,Position=0)]
        [string]
        $id,

        # The user name for the connection. This is obtained from the DoIT Mail Team.
        [Parameter(Mandatory=$true,Position=1)]
        [string]
        $userName,

        # The password for the connection. This is obtained from the DoIT Mail Team.
        [Parameter(Mandatory=$true,Position=2)]
        [string]
        $password,

        # The domain for the connection. The specified credentials must have administrative access to this domain. This is also the value that will be used by default by functions with a "domain" parameter.
        [Parameter(Mandatory=$true,Position=3)]
        [string]
        $domain,

        # The endpoint for the connection. If not specified, the default endpoint is used.
        [Parameter(Mandatory=$false,Position=4)]
        [string]
        $endpoint = $defaultEndpoint,

        # If specified, the connection will be saved for use in later sessions. The connection can only be used on the computer and by the user account on which it was originally created. 
        # If not specified, the connection can only be used during the session in which it is entered.
        [Parameter(Mandatory=$false,Position=5)]
        [switch]
        $save = $false
    )

    Begin{
    }

    Process{
        #Import the existing connections
        $existConns = Get-O365Connections
    
        #Check if the ID already exists
        If ($existConns.id -contains $id){
            #Throw an error if it exists
            Write-Error -Message "ID already exists. Please enter a different ID."
        }
        Else{
            #Create the credential object
            $cred = New-Object System.Management.Automation.PSCredential($userName, (ConvertTo-SecureString -AsPlainText -Force $password)) 
        
            #Create a new object for storing the existing connections and add the existing connections
            $savedConns = @()
            $allConns = @()
            $savedConns = $existConns | Where-Object {$_.saved -eq $True}
            $allConns = $existConns

            #Set the file path for the password file
            If ($save){$pwFile = "$credPath\$id.xml"}
            Else{$pwFile = $null}

            #Create an object for storing the new connection and add the properties and values
            $newConn = $null
            $newConn = New-Object System.Object
            $newConn | Add-Member -MemberType NoteProperty -Name "id" -Value $id
            $newConn | Add-Member -MemberType NoteProperty -Name "userName" -Value $userName
            $newConn | Add-Member -MemberType NoteProperty -Name "domain" -Value $domain
            $newConn | Add-Member -MemberType NoteProperty -Name "endpoint" -Value $endpoint
            $newConn | Add-Member -MemberType NoteProperty -Name "cred" -Value $cred
            $newConn | Add-Member -MemberType NoteProperty -Name "pwFile" -Value $pwFile

            #Add the new connections to the existing connections and set the global variables
            $Global:O365Connections = @()
            Foreach($allConn in $allConns){$Global:O365Connections += $allConn}
            $Global:O365Connections += $newConn

            #Save connections if specified
            If ($save){
                #Created new array for storing the connections
                $newSavedConns = @()
                #Add the existing connections to the array
                Foreach ($savedConn in $savedConns){
                    $newSavedConns += $savedConn
                }
                #Add the new connection to the array
                $newSavedConns += $newConn
                
                Try{
                    #Export credentials to $id.xml
                    $cred | Export-CliXml "$pwFile" -ErrorAction Stop
                }
                Catch{
                    Write-Error $_
                    break
                }

                #Select the properties to export to CSV
                $newSavedConns = $newSavedConns | Select-Object -Property * -ExcludeProperty cred, saved

                Try{
                    #Export connection details to connections.csv
                    $newSavedConns | Export-Csv -Path $connectionsFile -Force -NoTypeInformation -ErrorAction Stop
                }
                Catch{
                    Write-Error $_
                    break
                }
            }

            #Set as the current connection
            Set-O365Connection -id $newConn.id
        }
    }
}

<#
.Synopsis
    Gets all the existing connections for the Office 365 API, both saved and in-memory. 
.DESCRIPTION
    This function gets all the existing connections for the Office 365 API, both saved and in-memory. It first checks that %LOCALAPPDATA%\WindowsPowerShell\ModuleData\WiscO365\Connections\Connections.csv exists. If the file doesn't exist, it is created with the proper headers. Connections.csv and any password files are imported and added to $GlobalO365Connections with any temporary, in-memory connections. 
.OUTPUTS
    An array of all existing connections is returned.
#>
Function Get-O365Connections(){
    Begin{
        #Check if the connections file exists
        if(!(Test-Path -Path $connectionsFile)){
            #If it doesn't exists, create it
            $header = '"id","userName","domain","endpoint","pwFile"'
            
            Try{
                $header | Out-File $connectionsFile -Encoding ASCII -ErrorAction Stop
            }
            Catch{
                Write-Error $_
                break
            }
        }
    }

    Process{
        #Import the existing connections
        $importConns = @()
        Try{
            $importConns = Import-Csv -Path $connectionsFile -ErrorAction Stop
        }
        Catch{
            Write-Error $_
            break
        }
        
        #Create an array to store all the connections
        $allConns = @()
        #Add each imported connection to the array
        Foreach($importConn in $importConns){

            Try{
                #Import the credential file
                $cred = Import-Clixml -Path $importConn.pwFile -ErrorAction Stop
                #Add the credential object to the array
                $importConn | Add-Member -MemberType NoteProperty -Name "cred" -Value $cred -Force
                #Add the connection details to the array
                $allConns += $importConn
            }
            Catch{
                Write-Error $_
                break
            }
        }
            
        #Add each connection in the allConns array to the global connection array
        Foreach($O365Connection in $Global:O365Connections){$allConns += $O365Connection}

        #Check each connection for the password file
        Foreach($allConn in $allConns){
            $saved = $null
            #Check if the password file exists. If yes, then saved is true. Else, false. 
            If($allConn.pwFile){$saved = $true}
            Else{$saved = $false}
            #Add the saved property to the allConn array
            $allConn | Add-Member -MemberType NoteProperty -Name "saved" -Value $saved -Force
        }

        #Remove duplicates from the array
        $allConns = $allConns | select -Unique *

        #Clear the global connections and replace with object in array
        $Global:O365Connections = @()
        Foreach($allConn in $allConns){$Global:O365Connections += $allConn}
          
        return $allConns
    }
}

<#
.Synopsis
    Sets an existing Office 365 API connection as the current connection or clears the current connection. 
.DESCRIPTION
    This function sets an existing Office 365 API connection as the current connection or clears the current connection. If the ID parameter is not specified, the current connection is cleared by setting $Global:O365CurrentConnection and $Global:O365Session as $null. Otherwise, the function first gets all existing connections using Get-O365Connections. If the specified connection does not exist, an error is thrown and processing stops. A new WebRequestSession object is created and stored in $Global:O365Session. The selected connection details are then stored in $Global:O365CurrentConnection..
.OUTPUTS
    $Global:O365Session: Session information for the API.
    $Global:O365CurrentConnection: Connection information for the selected connection.
#>
Function Set-O365Connection{
    [CmdletBinding()]
    Param
    (
        # The unique id for the connection. If not specified, the current connection will be cleared.
        [Parameter(Mandatory=$false,Position=0)]
        [string]
        $id
    )

    Process{
        #If the id is null (not specified), clear the global current connection & session
        If(!$id){
            $Global:O365CurrentConnection = $null
            $Global:O365Session = $null
        }
        Else{
            #Get all existing connections
            $allConns = Get-O365Connections

            #Check if the ID exists
            If ($allConns.id -notcontains $id){
                #Throw an error if it exists
                Write-Error -Message "Connection does not exist." -ErrorAction Stop
            }
            Else{    
                $selectedConn = $null

                #Get the details for the selected connections
                $selectedConn = $allConns | Where-Object {$_.id -eq $id}

                #Set the global variables
                #Create a web session object
                $Global:O365Session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
                #Set the global current connection
                $Global:O365CurrentConnection = $selectedConn
            }
        }
    }
}

<#
.Synopsis
    Removes an existing Office 365 API connection from memory and disk storage.
.DESCRIPTION
    The function removes an existing Office 365 API connection from memory and disk storage. First, the function gets all existing connections using Get-O365Connections. If the specified connection ID doesn't exist, an error is thrown and processing stops. The specified connection is removed from $Global:O365Connections. If the connection was saved, it is also removed from Connections.csv and its corresponding password file is deleted. If the specified connection was set as the current connection, the current connection is cleared using Set-O365Connection.
.OUTPUTS
    Connections.csv: Connection details for saved connections (specified connection removed).
	$id.xml: Encrypted password file for saved connection (file for specified connection removed).
    $Global:O365Connections: An array of all connections, both saved and temporary (specified connection removed)..
#>
Function Remove-O365Connection{
    [CmdletBinding()]
    Param
    (
        # The unique id for the connection.
        [Parameter(Mandatory=$true,Position=0)]
        [string]
        $id
    )

    Process{
        #Create arrays for storing connections
        $allConns = @()
        $savedConns = @()

        #Get existing connections
        $allConns = Get-O365Connections
        
        #Check if the ID exists
        If ($allConns.id -notcontains $id){
            #Throw an error if it exists
            Write-Error -Message "Connection does not exist." -ErrorAction Stop
        }
        Else{     
            #Filter out the saved connections
            $savedConns = $allConns | Where-Object {$_.saved -eq $True}

            #Determine if this connection is saved
            $isSaved = $null
            $isSaved = ($allConns | Where-Object {$_.id -eq $id}).saved

            #Get the password file location for the specified connections
            $pwFile = $null
            $pwFile = ($allConns | Where-Object {$_.id -eq $id}).pwFile

            #Remove the connection details for the specified connection
            $allConns = $allConns | Where-Object {$_.id -ne $id}
            $savedConns = $savedConns | Where-Object {$_.id -ne $id}
        
            #Remove blank lines     
            $allConns = $allConns | Where-Object {$_}
            $savedConns = $savedConns | Where-Object {$_}

            #Remove duplicate connections
            $allConns = $allConns | select -Unique *
            $savedConns = $savedConns | select -Unique *

            #Delete the password file if it exists
            If($pwFile){
                Try{
                    Remove-Item -Path $pwFile -Force -ErrorAction Continue
                }
                Catch{
                    Write-Error $_
                }
            }

            #Export the CSV with the saved connections, if the connection being removed is a saved connection
            If($isSaved){
                Try{
                    #Export connection details to connections.csv. This removes the specified connection. 
                    $savedConns | Export-Csv $connectionsFile -NoTypeInformation -Force -ErrorAction Stop
                }
                Catch{
                    Write-Error $_
                    break
                }
            }

            #If the specified connection is the current connection, clear the current connection.
            If($Global:O365CurrentConnection.id -eq $id){
                Set-O365Connection
            }

            #Clear the global connections and replace with the existing connections
            $Global:O365Connections = @()
            Foreach($allConn in $allConns){
                $Global:O365Connections += $allConn
            }
        }
    }
}

<#
.Synopsis
    Gets the currently selected connection.
.DESCRIPTION
    Gets the currently selected connection by returning the values in $Global:O365CurrentConnection.
.OUTPUTS
    The details of the currently selection connection are returned.
#>
Function Get-O365CurrentConnection(){
    Process{
        #Return the global currently selected connection
        $Global:O365CurrentConnection
    }
}

<#
.Synopsis
    Updates the API functions and documentation in the module by getting the methods from the Domain Admin API Documentation.
.DESCRIPTION
    This function updates the API functions and documentation in the module by getting the methods from the Domain Admin API Documentation. It first calls the script $PSScriptRoot\Update\UpdateAPIFunctions.ps1, where most of the processing takes place. This gets the latest Domain Admin API Documentation from the web using the function Get-O365DomainAdminDoc. Each method found in the documentation is parsed. 
	
	First, the method name is normalized. A list of verbs is imported from "$PSScriptRoot\actionWords.csv" and the method name is matched to one of the verbs in the originalVerbs column. If the verb is not found in the list, the verb is assumed to be the method name string until the first capital letter and is added to actionWords.csv for later use. If the verb is found in the list and the corresponding value in the approvedVerbs column of actionWords.csv is not null, this indicates that the original verb is an unapproved PowerShell verb. This value in the approvedVerbs column is substituted as the verb, but the original verb is still retained to be used in a function alias. If the original verb is an unapproved verb but has no approved verb specified, this unapproved verb is used (this has the potential to cause conflicts, but this is unlikely).  The original verb is trimmed from the original method name, with the remaining string assumed to be the noun. A list of unallowed words is imported from "$PSScriptRoot\unallowedWords.txt".  Any words in this list are trimmed from the noun. This is to prevent conflicts with functions from other modules (especially Microsoft's official Office 365 module) and to keep the naming convention consistent throughout all functions in the module. Finally, the function name is generated in the form Verb-PrefixNoun. The prefix for all functions is O365. If an unapproved verb existed, a name with the original verb is also generated in this form.
	
	Next, a header is generated for the function, which includes documentation, the function name, and any aliases. The Synopsis section is set as the method purpose from the documentation. The Outputs section is set as the method return from the documentation. The Link section is set to the web address for the Domain Admin API Documentation. The function name is the name value that was normalized earlier. If an unapproved verb was found, a normalized name with this verb is included as an alias. If the includeAlias switch is selected, the original method name from the documentation is added as an alias.
	
	Next, parameters are generated for the function. The parameter description is set as the parameter description from the documentation. If "required" in the documentation is "1", the Mandatory property is set to $True. The exception to this is any function that contains a "domain" parameter. Regardless of the "required" value in the documentation, a "domain" parameter's Mandatory property is always set to $False. This allows the user to avoid entering the "domain" value, since the "domain" value is set in the connection properties. Any "domain" parameter is set to $Global:O365CurrentConnection.domain by default, unless overridden by the user specifying a different value. The parameter position is determined by the order in which the parameters appear in the documentation, with the first parameter's position set as "0", the second parameter's position set as "1", etc. If the documentation contains a pre-defined set of values for the parameter, the parameter is set to validate this set of values. These values are determined by looking for "|" in the parameter description and splitting the string on the "|".
	
	Next, a JSON string is generated for the body of the API method call. This includes the original method name, parameter names, and variables for the parameter value. 
	
	The function is then formatted as an Advanced PowerShell Function, with all the previously described values included. Each function includes a call to the helper functions Test-HelperO365Connection (checks that a connection has been set before calling the API method) and Invoke-HelperO365APIFunction (calls the API method using Invoke-RestMethod and processes the return values). Both of these helper functions can be found at "$PSScriptRoot\HelperFunctions.ps1". A ps1 file is generated with all the functions and stored as "%LOCALAPPDATA%\WindowsPowerShell\ModuleData\WiscO365\New APIFunctions.ps1". Control is then returned to the original function, Update-O365APIFunctions.
	
    If a previous API Functions file exists as "%LOCALAPPDATA%\WindowsPowerShell\ModuleData\WiscO365\APIFunctions.ps1", this is compared to the new API Functions file. If the new API Functions file is different, the old API Functions file and Help File are moved to "%LOCALAPPDATA%\WindowsPowerShell\ModuleData\WiscO365\Old API Functions" and renamed for archiving. If the new API Functions file is different or if a previous file didn't exist, the new API Functions file is renamed "APIFunctions.ps1".  The module is then reloaded using the function "Import-Module WiscO365 -Force -DisableNameChecking". Finally, a new help file is generated for the module as "%LOCALAPPDATA%\WindowsPowerShell\ModuleData\WiscO365\WiscO365 Help.html". This is generated using the psDoc script ("$PSScriptRoot\Update\psDoc-master\src\psDoc.ps1"). This has been modified to include UW branding and extra sections as defined by the files docFooter.html, docTitle.html, extraBody.html, and extraNav.html in "$PSScriptRoot\Update\psDoc-master\src". If the new functions file is not different, it is deleted.
.OUTPUTS
    If no new updates: No Updates | No new API Functions were found.
    If new updates: Updates | New API Functions were found and added.
        APIFunctions.ps1
        WiscO365 Help.html
    
#>
Function Update-O365APIFunctions{
    [CmdletBinding()]
    Param
    (
        # Includes the original API method name as an alias.
        [switch]
        $includeAlias = $false
    )

    Process{
        $oldFunctionsPath = "$moduleDataPath\APIFunctions.ps1"
        $newFunctionsPath = "$moduleDataPath\New APIFunctions.ps1"
        $archiveDir = "$moduleDataPath\Old API Functions"
        $helpPath = "$moduleDataPath\WiscO365 Help.html"

        $oldFunctionsExist = Test-Path -Path $oldFunctionsPath
        $helpExists = Test-Path -Path $helpPath

        #Invoke script to update the API Functions
        Try{
            Invoke-Expression "$PSScriptRoot\Update\UpdateAPIFunctions.ps1 -outputFile '$newFunctionsPath' -includeAlias `$includeAlias" -ErrorAction Stop
        }
        Catch{
            Write-Error $_
            break
        }

        If($oldFunctionsExist){
            Try{
                #Compare the files
                $oldFunctions = Get-Content -Path $oldFunctionsPath -ErrorAction SilentlyContinue
                $newFunctions = Get-Content -Path $newFunctionsPath -ErrorAction SilentlyContinue
                $difference = Compare-Object -ReferenceObject $oldFunctions -DifferenceObject $newFunctions -ErrorAction SilentlyContinue
            }
            Catch{
                Write-Error $_
            }
        }
        Else{
            $difference = $true
        }

        #If the function files are different
        If($difference){
            #Get the date and time for the function archive
            $dateTime = Get-Date -Format "yyyyMMddHHmmss"

            Try{
                #Move the original function file to an archive in case we need it later
                If($oldFunctionsExist){Move-Item -Path $oldFunctionsPath -Destination "$archiveDir\APIFunctions $dateTime.ps1" -Force -ErrorAction SilentlyContinue}
                #Move the new function file to the module data root
                Move-Item -Path $newFunctionsPath -Destination $oldFunctionsPath -Force -ErrorAction SilentlyContinue
                #Move the original HTML help file to an archive in case we need it later
                If($helpExists){Move-Item -Path $helpPath -Destination "$archiveDir\WiscO365 Help $dateTime.html" -Force -ErrorAction SilentlyContinue}
            }
            Catch{
                Write-Error $_
            }
            
            Try{
                #Get the original startup preference
                $startupPath = "$moduleDataPath\Preferences\Startup.txt"
                $origStartupPref = Get-Content $startupPath -ErrorAction SilentlyContinue
                #If the startup menu was enabled
                If($origStartupPref -like "*1*"){
                    #Disable the startup menu so it doesn't show when we update
                    Set-O365StartupPreferences
                }
            }
            Catch{
                Write-Error $_
            }
            
            #Reload the module
            Import-Module WiscO365 -Force -DisableNameChecking

            Try{
                #Generate HTML documentation
                Invoke-Expression "$PSScriptRoot\Update\psDoc-master\src\psDoc.ps1 -moduleName WiscO365 -outputDir $moduleDataPath -fileName 'WiscO365 Help.html'" -ErrorAction SilentlyContinue
            }
            Catch{
                Write-Error $_
            }

            Try{
                #If the startup menu was enabled
                If($origStartupPref -like "*1*"){
                    #Re-enable the startup menu so it doesn't show when we update
                    Set-O365StartupPreferences -showStartupMenu
                }
            }
            Catch{
                Write-Error $_
            }

            Write-Output "Updates | New API Functions were found and added."
        }

        #If the function files are not different
        Else{
            Try{
                #Remove the updated functions file since it isn't different
                Remove-Item -Path $newFunctionsPath -Force -ErrorAction SilentlyContinue
            }
            Catch{
                Write-Error $_
            }

            Write-Output "No Updates | No new API Functions were found."
        }
    }
}

<#
.Synopsis
    Launches the HTML help documentation for the module.
.DESCRIPTION
    If documentation has been generated using Update-O365APIFunctions ("%LOCALAPPDATA%\WindowsPowerShell\ModuleData\WiscO365\WiscO365 Help.html"), this is launched. Otherwise, the initial help document is launched ("$PSScriptRoot\WiscO365 Initial Help.html").   
.OUTPUTS
    WiscO365 Help.html or WiscO365 Initial Help.html
#>
Function Get-O365Help{
    If(Test-Path "$moduleDataPath\WiscO365 Help.html"){
        Start-Process -FilePath "$moduleDataPath\WiscO365 Help.html"
    }
    Else{
        Start-Process -FilePath "$PSScriptRoot\WiscO365 Initial Help.html"
    }
}

<#
.Synopsis
    Sets the preference to determine if the startup menu will be shown when importing the module.
.DESCRIPTION
    If the switch "showStartupMenu" is included, the startup menu will be displayed when importing the module. Otherwise, the menu will not be displayed. This preference is saved in the file "%LOCALAPPDATA%\WindowsPowerShell\ModuleData\WiscO365\Preferences\Startup.txt" as either "0" (don't show) or "1" (show). The startup menu is defined by the function Show-HelperO365StartupMenu in "$PSScriptRoot\HelperFunctions.ps1".
.OUTPUTS
    Startup.txt
#>
Function Set-O365StartupPreferences{
    [CmdletBinding()]
    Param
    (
        # Determines if the startup menu will be shown
        [Parameter(Mandatory=$false,Position=0)]
        [switch]
        $showStartupMenu = $false
    )

    Process{
        #Startup preferences location
        $startupPath = "$moduleDataPath\Preferences\Startup.txt"

        If($showStartupMenu){
            $startupPref = 1
        }
        Else{
            $startupPref = 0
        }

        Try{
            $startupPref | Out-File -FilePath $startupPath -Force -ErrorAction Stop
        }
        Catch{
            Write-Error $_
            break
        }
    }
}