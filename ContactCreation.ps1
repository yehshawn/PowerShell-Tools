$CSVFile = import-csv "C:\AccountUpdate\Contact.csv"
$DateTime = Get-Date -f "yyyyMMdd HHmmss"
$ReportFile = "C:\AccountUpdate\Contact Log "+$DateTime+".csv"


foreach ($line in $CSVFile) {
	New-MailContact -Name $line.Name -FirstName $line.FirstName -LastName $line.LastName -ExternalEmailAddress $line.ExternalEmailAddress -OrganizationalUnit $line.OrganizationalUnit  -alias $line.alias
}

