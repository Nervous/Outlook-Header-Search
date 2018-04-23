$MatchString = "Matchingstring"
Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
function Get-MailboxFolder($folder)
{
    $Headers =
    foreach ( $MailItem in $folder.items ) { 
        if (($MailItem -eq $null) -Or ($MailItem.PropertyAccessor -eq $null)) {continue}
        $MailItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E") 
    }
        $MatchingHeaders = $Headers | where { $_.contains( $MatchString ) }
        $MatchingHeaders | Select-Object -First 1000 >> result.txt

    foreach ($f in $folder.folders)
    {
        Get-MailboxFolder $f
    }
}

$ol = new-object -com Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$mailbox = $ns.stores | where {$_.ExchangeStoreType -eq 0}
$mailbox.GetRootFolder().folders | foreach { Get-MailboxFolder $_}