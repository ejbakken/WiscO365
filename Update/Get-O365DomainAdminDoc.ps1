<#
.Synopsis
    Gets the documentation for the Domain Admin API, which is what you are viewing right now
.OUTPUTS
    failure: undef
	success: JSON array with all Domain Admin API documentation
.LINK
    https://wiscmail.wisc.edu/admin/index.php?action=domain-domainadmin_api
#>
Function Get-O365DomainAdminDoc{
    [CmdletBinding()]

    Param
    (
		# Domain name (e.g. bar.wisc.edu)
		[Parameter(Mandatory=$False,Position=0)]
		$domain = $Global:O365CurrentConnection.domain
    )

    Begin{
        #Check to ensure that a connection has been set.
        Test-HelperO365Connection
    }

    Process{
        #Create hash table for JSON command
        $body = @{
			"action" = "getDomainAdminDoc";
			"domain" = "$domain"
        }

        #Invoke the function 
        Invoke-HelperO365Function -body $body
    }
}