$connected = $false
try {
	Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
	$pwd = ConvertTo-SecureString -string $SharePointAdminPWD -AsPlainText -Force
	$cred = New-Object System.Management.Automation.PSCredential $SharePointAdminUser, $pwd
	$null = Connect-SPOService -Url $SharePointBaseUrl -Credential $cred
    HID-Write-Status -Message "Connected to Microsoft SharePoint" -Event Information
    HID-Write-Summary -Message "Connected to Microsoft SharePoint" -Event Information
	$connected = $true
}
catch
{	
    HID-Write-Status -Message "Could not connect to Microsoft SharePoint. Error: $($_.Exception.Message)" -Event Error
    HID-Write-Summary -Message "Failed to connect to Microsoft SharePoint" -Event Failed
}

if ($connected)
{
	try {
		Remove-SPOSite -Identity $spSiteUrl -NoWait -Confirm:$false
		HID-Write-Status -Message "Removed Site [$spSiteTitle] with url [$spSiteUrl]" -Event Success
		HID-Write-Summary -Message "Successfully removed Site [$spSiteTitle] with url [$spSiteUrl]" -Event Success
	}
	catch
	{
		HID-Write-Status -Message "Could not remove Site [$spSiteTitle]. Error: $($_.Exception.Message)" -Event Error
		HID-Write-Summary -Message "Failed to remove Site [$spSiteTitle]" -Event Failed
	}
    finally
    {        
        Disconnect-SPOService
        Remove-Module -Name Microsoft.Online.SharePoint.PowerShell
    }
}
