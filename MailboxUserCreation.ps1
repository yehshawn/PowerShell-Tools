[CmdletBinding()]
param(
	[Parameter(Position=0, Mandatory=$true)]
	[System.String]
	$CSVFile,
	[Parameter(Position=1, Mandatory=$false)]
	[System.String]
	$PasswordLength = 12,
	[Parameter(Position=2, Mandatory=$false)]
	[System.String]
	$NonAlphanumericCharacters = 0
	)

begin {
	$DateTime = Get-Date -f "yyyyMMdd HHmmss"
	$ReportFile = "C:\AccountUpdate\MailboxUser Log "+$DateTime+".csv"
	$i = 1
	$mailboxes = Import-Csv $CSVFile
	$props = $mailboxes | gm -MemberType NoteProperty | select -ExpandProperty Name
	$report = @()
}

process {
	foreach($mailbox in $mailboxes) {
		Write-Progress -Id 0 -Activity "Creating Mailboxes..." -Status "$i of $($mailboxes.count)" -PercentComplete (($i / ($mailboxes.count+1))*100)
		$properties = @{}
		$props | %{
			$properties.add($_,$mailbox.$_)
		}
		$types = @{
			uppers = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
			lowers = 'abcdefghjkmnpqrstuvwxyz'
			digits = '23456789'
			symbols = '@$%'
		}
		$four = foreach($thisType in $types.Keys){
			Get-Random -Count 1 -InputObject ([char[]]$types[$thisType])
		}
		[char[]]$allSupportedChars = $types.Values -join ''
		$theRest = Get-Random -count ($PasswordLength - ($types.Keys.Count)) -InputObject $allSupportedChars
		[string]$password = ($four + $theRest -join '') | Sort-Object {Get-Random}
		$securepassword = ConvertTo-SecureString -AsPlainText $password -Force
		$properties.add("Password",$securepassword)
		
		if(New-Mailbox @properties) {
			$mailbox | Add-Member -MemberType NoteProperty -Name Password -Value $Password
			$report += $mailbox
		}
		$properties.clear()
		$i++
	}
}

end {
	$report | Export-Csv -Path $ReportFile -NoTypeInformation
}
