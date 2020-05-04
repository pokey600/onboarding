<# email attacment portion borrowed from https://community.spiceworks.com/topic/2142884-powershell-script-to-extract-email-attachments-having-trouble
Used for ad and email creation portion https://adamtheautomator.com/powershell-import-csv-foreach/
Name of the mailbox to pull attachments from #>
$address = 'onbarding@lumicor.com'
$TargetFolderName = 'C:\Users\Alex\Documents'
$Subject = 'onbarding'
$UserCredential = Get-Content 'C:\Users\Alex\Documents\usercredtial.txt'
Connect-ExchangeOnline -Credential $UserCredential
$minLength = 8 ## characters
$maxLength = 16 ## characters
$length = Get-Random -Minimum $minLength -Maximum $maxLength
$nonAlphaChars = 5
Import-Module "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll" 
$EWS = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService 'Exchange2013',([timezoneinfo]::Utc)
$EWS.AutodiscoverUrl($address)
$folderID = new-object Microsoft.Exchange.WebServices.Data.FolderId 'Inbox', $address
$folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($EWS, $folderID)
Write-Host "found $($folder.DisplayName) folder"

$folderview = New-Object Microsoft.Exchange.WebServices.Data.FolderView 100
$targetFolder = $null
foreach ($f in $folder.FindFolders($folderview)) {
	if ($f.DisplayName -eq $TargetFolderName) {
		$targetFolder = $f
		break;
	}
}
if ($targetFolder) {
	Write-Host "found $($targetFolder.DisplayName) folder"
}
else {
	$targetFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder $EWS
	$targetFolder.DisplayName = $TargetFolderName
	Write-Host "Create $($targetFolder.DisplayName) folder"
	$targetFolder.Save($folderID)
}

$filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject, $Subject)
$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView 100
$view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet
$mailProperties = New-Object Microsoft.Exchange.WebServices.Data.PropertySet (
						[Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::TextBody,
						[Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject,
						[Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender
					)

$size = 0; $view.Offset = 0; $req = 0;
do {
	$req++
	$MailItems = $folder.FindItems($filter, $view)
	if ($view.Offset -eq 0) { Write-Host ('Messages Total: {0}' -f $MailItems.TotalCount) }
	$view.Offset += $MailItems.Items.Count
	foreach ($item in $MailItems.Items) {
		$size += $item.Size
		$item2 = $item.Move($targetFolder.Id)
		$item2.Load($mailProperties)
    # [...]
    Import-Csv "C:\Users\Alex\Documents\onboarding.csv" | ForEach-Object {
      $password = [System.Web.Security.Membership]::GeneratePassword($length, $nonAlphaChars)
      $Username = $_.FirstInital + $_.LastName + $_.BirthYear
        New-ADUser `
          -Name $($Username) `
          -GivenName $_.FirstName `
          -Surname $_.LastName `
          -AccountPassword $password `
          -Enabled $True `
          -ChangePasswordAtLogon $True `
          -Department $_.Department `
        Set-Mailbox { 
          -Identity $Username 
          -EmailAddresses @{add="smtp:" + $Username + "@lumicor.com"} 
          -AzureADAuthorizationEndpointUri $password 
        }
    $_.FirstName, $_.LastName, $_.BirthYear, $Username, $password, $_.Department > "C:\Users\Alex\Documents\reply.csv" 
    set-GPPermission{ 
      -Name $_.Department
      -TargetName $_.samaccountname 
      -PermissionLevel GpoApply 
      -TargetType User 
      }
    }
    Send-MailMessage {
      -To 'managers@lumicor.com', 'apokrandt2000@lumicor.com' 
      -From 'onboarding@lumicor.com' 
      -Attachments 'C:\Users\Alex\Documents\reply.csv' 
      -Subject 'Onbarding' 
      -UseSsl $True `
    }
    Remove-Item {
      -Path 'C:\Users\Alex\Documents\onbardin.csv, C:\Users\Alex\Documents\reply.csv' 
    }
	}
} while ($MailItems.MoreAvailable)
Write-Host ('Size: {0}, Requests: {1}' -f $size, $req)
