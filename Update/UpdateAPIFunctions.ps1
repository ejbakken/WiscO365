param(
    #Defines the output location for the file
    [Parameter(Mandatory=$false,Position=0)]
    [string]
    $outputFile = "$moduleDataPath\New APIFunctions.ps1",

    #Defines if the alias is included
    [Parameter(Mandatory=$false,Position=1)]
    $includeAlias = $false
    )

$actionWordsPath = "$PSScriptRoot\actionWords.csv"
$unallowedWords = Get-Content "$PSScriptRoot\unallowedWords.txt"

If(Test-Path $outputFile){Remove-Item $outputFile -Force}

$doc = Get-O365DomainAdminDoc -domain $Global:O365CurrentConnection.domain
$methods = $doc.methods

$methodNames = $methods | Get-Member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'

Function Create-Function($methodName, $method, $outputFile, $actionWordsPath, $unallowedWords, $includeAlias){
    $headers = Create-Headers -method $method -methodName $methodName -includeAlias $includeAlias
    $params = Create-Parameters -parameters $method.parameters
    $JSON = Create-JSON -parameters $method.parameters -methodName $methodName
    

Add-Content $outputFile `
$headers
Add-Content $outputFile `
"    Param
    (
$params
    )
"
Add-Content $outputFile `
"    Begin{
        #Check to ensure that a connection has been set.
        Test-HelperO365Connection
    }

    Process{
        #Create hash table for JSON command
        `$body = @{
$JSON
        }
"
Add-Content $outputFile `
"        #Invoke the function 
        Invoke-HelperO365Function -body `$body
    }
}
"
}

Function Create-Headers($method, $methodName, $includeAlias){
$outputs = (Normalize-Outputs -outputs $method.return).trim()
$names = Normalize-FunctionName -origName $methodName -actionWordsPath $actionWordsPath -unallowedWords $unallowedWords

$normalizedName = $names.normalizedName
$normalizedOriginalName = $names.normalizedOriginalName
$originalName = $methodName

$string = ""

$string+=
"<#
.Synopsis
    $($method.purpose)
.OUTPUTS
    $outputs
.LINK
    https://wiscmail.wisc.edu/admin/index.php?action=domain-domainadmin_api
#>
Function $normalizedName{
    [CmdletBinding()]
"

If($normalizedOriginalName){
$string+=
"    [Alias(`"$normalizedOriginalName`")]
"
}

If($includeAlias){
$string+=
"    [Alias(`"$originalName`")]
"
}

return $string
}

Function Create-Parameters($parameters){
        $paramNames = $parameters | Get-Member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'
        $paramString = ""

        $i=0
        $j=$paramNames.count
        Foreach($paramName in $paramNames){
            $isRequired = ($parameters.$paramName.required -eq '1')
            if($paramName -eq "domain"){$isRequired = $false}

            if($parameters.$paramName.description -like "*|*") {
                $predefVals = $parameters.$paramName.description.Split("|") | % {"'$_'"}
                $predefVals = $predefVals -join ","
            }
            else {$predefVals = $null}

            $paramString += "`t`t" + "# " + "$($parameters.$paramName.description)" + "`n"
            $paramString += "`t`t" + "[Parameter(Mandatory=$" + "$isRequired,Position=$i)]" + "`n"
            $paramString += if($predefVals){"`t`t"+ "[ValidateSet($predefVals)]" + "`n"}
            $paramString += "`t`t" + "$" + "$paramName" 
            $paramString += if($paramName -eq "domain"){" = `$Global:O365CurrentConnection.domain"}
            $paramString += if($i+1 -lt $j){",`n`n"}

            $i++
        }
        return $paramString
}

Function Create-JSON($parameters, $methodName){
    $paramNames = $parameters | Get-Member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'
    $JSONString = ""
        
    $JSONString += "`t`t`t" + '"action" = ' + '"' + $methodName + '";' + "`n"

    $i=0
    $j=$paramNames.count
    Foreach ($paramName in $paramNames){
        $JSONString += "`t`t`t" + '"' + $paramName + '"' + " = " + '"$' + $paramName + '"'
        $JSONString += if($i+1 -lt $j){";`n"}

        $i++
    }
    return $JSONString
}

Function Normalize-FunctionName($origName, $actionWordsPath, $unallowedWords){
    #Import the action words
    $actionWords = Import-Csv $actionWordsPath

    #Use PowerShell convention "Verb-Noun"
    $origVerb = $null
    $verb = $null
    $noun = $null

    #Find the verb
    $match = $false
    Foreach ($actionWord in $actionWords){
        #Compare the method name to the pre-defined action words (verbs)
        If($origName -like "$($actionWord.originalVerb)*"){
            $match = $true 
            #Check if the original verb has an alternate approved verb (which means the original is unapproved)
            If($actionWord.approvedVerb){
                #If so, the verb will be the alternate approved verb, not the original verb
                $verb = $actionWord.approvedVerb
                $origVerb = $actionWord.originalVerb
            }
            Else{
                #Otherwise, the verb is the original verb
                $verb = $actionWord.originalVerb
                $origVerb = $actionWord.originalVerb
            }
        }
    }
            
    #If the original name isn't like any of the verbs in the list
    If(!$match){
        #Separate at the first capital letter and remove all except the first word
        $verb = ($origName -creplace '([A-Z\W_]|\d+)(?<![a-z])',' $&')
        $verb = $verb.Substring(0, $verb.IndexOf(' '))
        $origVerb = $verb

        #Add the verb to the list
        #Create an object to store the new verb information and add the values
        $newVerb = New-Object System.Object
        $newVerb | Add-Member -MemberType NoteProperty -Name "originalVerb" -Value $origVerb
        $newVerb | Add-Member -MemberType NoteProperty -Name "approvedVerb" -Value ""

        #Create an object for storing all the verbs and add the existing verbs
        $allVerbs = @()
        Foreach ($actionWord in $actionWords){
            $allVerbs += $actionWord
        }
        #Add the new verb to the list
        $allVerbs += $newVerb

        #Export all verbs to the existing CSV file
        Try{
            $allVerbs | Export-Csv -Path $actionWordsPath -NoTypeInformation -Force -ErrorAction SilentlyContinue
        }
        Catch{
            Write-Error $_
        }
    }

    #Trim the original word to get the noun
    $noun = $origName -replace $origVerb

    #Check for unallowed words in the noun. If present, at the beginning of the noun, trim them out.
    Foreach ($unallowedWord in $unallowedWords){
        If($noun -like "$unallowedWord*"){
            $noun = $noun -replace $unallowedWord
        }
    }

    #Check for special characters in the noun. If present, get rid of them, capitalize the next word, and combine.
    $noun = $noun -replace '[^a-zA-Z0-9]', ' '
    $noun = $noun.Trim()
    if($noun -like '* *'){
        $noun = (Get-Culture).TextInfo.ToTitleCase($noun)
        $noun = $noun.Replace(' ','')
    }

    #Capitalize the first letter of the verbs and noun
    if($verb) {$verb = $verb.substring(0,1).toupper()+$verb.substring(1)}
    if($origVerb) {$origVerb = $origVerb.substring(0,1).toupper()+$origVerb.substring(1)}
    if($noun) {$noun = $noun.substring(0,1).toupper()+$noun.substring(1)}

    #Create an object to store the names
    $names = New-Object System.Object

    #Combine to create the normalized name and add it to the names object
    $normalizedName = $verb + "-" + "O365" + $noun
    $names | Add-Member -MemberType NoteProperty -Name normalizedName -Value $normalizedName

    #If the verb was changed, create the normalized original name and add it to the names object
    If($verb -ne $origVerb){
        $normalizedOriginalName = $origVerb + "-" + "O365" + $noun
        $names | Add-Member -MemberType NoteProperty -Name normalizedOriginalName -Value $normalizedOriginalName
    }

    return $names
}

Function Normalize-Outputs ($outputs){
    $string = ""

    #Get the names of the outputs
    $outputNames = $outputs | Get-Member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'
    
    #Get the members of each of the outputs
    Foreach ($outputName in $outputNames){
        $outputMembers = $outputs.$($outputName)

        #Check if the members have more members
        $outputSubMemberNames = $outputMembers | Get-Member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'
        If($outputSubMemberNames){
            $string += $tab + $outputName
            $string += "`n`t"
            $tab += "`t"
            $string += Normalize-Outputs -outputs $outputMembers
        }
        Else{
            $string += $tab + $outputName + ": " + $outputMembers
            $string += "`n`t"
        }
    }
    return $string
}

Foreach ($methodName in $methodNames){
    Create-Function -methodName $methodName -method $methods.$methodName -outputFile $outputFile -actionWordsPath $actionWordsPath -unallowedWords $unallowedWords -includeAlias $includeAlias
}
