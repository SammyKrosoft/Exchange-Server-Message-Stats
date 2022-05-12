if ($args[0] -eq $null)      {
                $startoffset = 7
                }
                ELSE       {
                $startoffset = $args[0]
                }
$today = get-date 
$rundate = $($today).tostring("MM\/dd\/yyyy")
$startdate = $($today.adddays(-$startoffset)).tostring("MM\/dd\/yyyy")

$outfile_date = ([datetime]$rundate).tostring("yyyy_MM_dd") 
#$outfile = "email_stats_" + $outfile_date + ".csv" 
 
#$dl_stat_file = "DL_stats.csv" 
 
$accepted_domains = Get-AcceptedDomain |% {$_.domainname.domain} 
[regex]$dom_rgx = "`(?i)(?:" + (($accepted_domains |% {"@" + [regex]::escape($_)}) -join "|") + ")$" 

$mbx_servers = Get-ExchangeServer |? {$_.serverrole -match "Mailbox"}|% {$_.fqdn} 
[regex]$mbx_rgx = "`(?i)(?:" + (($mbx_servers |% {"@" + [regex]::escape($_)}) -join "|") + ")\>$" 

 
$msgid_rgx = "^\<.+@.+\..+\>$" 

$hts = get-exchangeserver |? {$_.serverrole -match "hubtransport" -or ($_.serverrole -match "mailbox" -and $_.admindisplayversion -match "15.")} |% {$_.name} 


$exch_addrs = @{} 
 
$msgrec = @{} 
$bytesrec = @{} 
 
$msgrec_exch = @{} 
$bytesrec_exch = @{} 
 
$msgrec_smtpext = @{} 
$bytesrec_smtpext = @{} 
 
$total_msgsent = @{} 
$total_bytessent = @{} 
$unique_msgsent = @{} 
$unique_bytessent = @{} 
 
$total_msgsent_exch = @{} 
$total_bytessent_exch = @{} 
$unique_msgsent_exch = @{} 
$unique_bytessent_exch = @{} 
 
$total_msgsent_smtpext = @{} 
$total_bytessent_smtpext = @{} 
$unique_msgsent_smtpext=@{} 
$unique_bytessent_smtpext = @{} 
 
$dl = @{} 
 
 
$obj_table = { 
@" 
Date = $rundate 
User = $($address.split("@")[0]) 
Domain = $($address.split("@")[1]) 
SentTotal = $(0 + $total_msgsent[$address]) 
SentMBTotal = $("{0:F2}" -f $($total_bytessent[$address]/1mb)) 
ReceivedTotal = $(0 + $msgrec[$address]) 
ReceivedMBTotal = $("{0:F2}" -f $($bytesrec[$address]/1mb)) 
SentInternal = $(0 + $total_msgsent_exch[$address]) 
SentInternalMB = $("{0:F2}" -f $($total_bytessent_exch[$address]/1mb)) 
SentExternal = $(0 + $total_msgsent_smtpext[$address]) 
SentExternalMB = $("{0:F2}" -f $($total_bytessent_smtpext[$address]/1mb)) 
ReceivedInternal = $(0 + $msgrec_exch[$address]) 
ReceivedInternalMB = $("{0:F2}" -f $($bytesrec_exch[$address]/1mb)) 
ReceivedExternal = $(0 + $msgrec_smtpext[$address]) 
ReceivedExternalMB = $("{0:F2}" -f $($bytesrec_smtpext[$address]/1mb)) 
SentUniqueTotal = $(0 + $unique_msgsent[$address]) 
SentUniqueMBTotal = $("{0:F2}" -f $($unique_bytessent[$address]/1mb)) 
SentInternalUnique  = $(0 + $unique_msgsent_exch[$address])  
SentInternalUniqueMB = $("{0:F2}" -f $($unique_bytessent_exch[$address]/1mb)) 
SentExternalUnique = $(0 + $unique_msgsent_smtpext[$address]) 
SentExternalUniqueMB = $("{0:F2}" -f $($unique_bytessent_smtpext[$address]/1mb)) 
"@ 
} 

$summary_stats = New-Object psobject -Property @{
"TotalMessagesSentpday"=0
"TotalMBSentpday"=0
"TotalMessagesReceivedpday"=0
"TotalMBReceivedpday"=0
"TotalInternalSentpday"=0
"TotalInternalMBpday"=0
"TotalExternalSentpday"=0
"TotalExternalSentMBpday"=0
"TotalExternalReceivedpday"=0
"TotalExternalReceivedMBpday"=0
"TotalInternalSentMBpday"=0
"TotalInternalReceivedMBpday"=0
"TotalExternalMBpday" = 0
"MessagesPerUserPerDay" = 0
"AverageMessageSizeKB" = 0
"UniqueUsers" = 0
"TotalMBReceived" = 0
"TotalMBSent" = 0
"TotalMessagesSent" = 0
"TotalMessagesReceived" = 0
}

$props = $obj_table.ToString().Split("`n")|% {if ($_ -match "(.+)="){$matches[1].trim()}} 
 
$stat_recs = @() 
 
function time_pipeline { 
param ($increment  = 1000) 
begin{$i=0;$timer = [diagnostics.stopwatch]::startnew()} 
process { 
    $i++ 
    if (!($i % $increment)){Write-host “`rProcessed $i in $($timer.elapsed.totalseconds) seconds” -nonewline} 
    $_ 
    } 
end { 
    write-host “`rProcessed $i log records in $($timer.elapsed.totalseconds) seconds” 
    Write-Host "   Average rate: $([int]($i/$timer.elapsed.totalseconds)) log recs/sec." 
    } 
} 

 
foreach ($ht in $hts){ 

trap { 

       write-host $("Error: " + $_.Exception.Message); 
       continue; 
    }


 
    Write-Host "`nStarted processing $ht" 
 
    get-messagetrackinglog -Server $ht -Start "$startdate" -End "$rundate 11:59:59 PM" -resultsize unlimited -ErrorAction "Stop"| 
    time_pipeline | %{ 
    
    
     
     
    if ($_.eventid -eq "DELIVER" -and $_.source -eq "STOREDRIVER"){ 
     
        if ($_.messageid -match $mbx_rgx -and $_.sender -match $dom_rgx) { 
             
            $total_msgsent[$_.sender] += $_.recipientcount 
            $total_bytessent[$_.sender] += ($_.recipientcount * $_.totalbytes) 
            $total_msgsent_exch[$_.sender] += $_.recipientcount 
            $total_bytessent_exch[$_.sender] += ($_.totalbytes * $_.recipientcount) 
         
            foreach ($rcpt in $_.recipients){ 
            $exch_addrs[$rcpt] ++ 
            $msgrec[$rcpt] ++ 
            $bytesrec[$rcpt] += $_.totalbytes 
            $msgrec_exch[$rcpt] ++ 
            $bytesrec_exch[$rcpt] += $_.totalbytes 
            } 
             
        } 
         
        else { 
            if ($_messageid -match $messageid_rgx){ 
                    foreach ($rcpt in $_.recipients){ 
                        $msgrec[$rcpt] ++ 
                        $bytesrec[$rcpt] += $_.totalbytes 
                        $msgrec_smtpext[$rcpt] ++ 
                        $bytesrec_smtpext[$rcpt] += $_.totalbytes 
                    } 
                } 
         
            } 
                 
    } 
     
     
    if ($_.eventid -eq "RECEIVE" -and $_.source -eq "STOREDRIVER"){ 
        $exch_addrs[$_.sender] ++ 
        $unique_msgsent[$_.sender] ++ 
        $unique_bytessent[$_.sender] += $_.totalbytes 
         
            if ($_.recipients -match $dom_rgx){ 
                $unique_msgsent_exch[$_.sender] ++ 
                $unique_bytessent_exch[$_.sender] += $_.totalbytes 
                } 
 
            if ($_.recipients -notmatch $dom_rgx){ 
                $ext_count = ($_.recipients -notmatch $dom_rgx).count 
                $unique_msgsent_smtpext[$_.sender] ++ 
                $unique_bytessent_smtpext[$_.sender] += $_.totalbytes 
                $total_msgsent[$_.sender] += $ext_count 
                $total_bytessent[$_.sender] += ($ext_count * $_.totalbytes) 
                $total_msgsent_smtpext[$_.sender] += $ext_count 
                 $total_bytessent_smtpext[$_.sender] += ($ext_count * $_.totalbytes) 
                } 
                               
        } 
         
    if ($_.eventid -eq "expand"){ 
        $dl[$_.relatedrecipientaddress] ++ 
        }
       
    }     
  
}   

 
foreach ($address in $exch_addrs.keys){ 
 
$stat_rec = (new-object psobject -property (ConvertFrom-StringData (&$obj_table))) 
$stat_recs += $stat_rec | select $props 
$summary_stats.TotalMessagesSentpday = [Math]::Ceiling([decimal]($summary_stats.TotalMessagesSentpday + $stat_rec.SentTotal)/$startoffset)
$summary_stats.TotalMessagesReceivedpday = [Math]::Ceiling([decimal]($summary_stats.TotalMessagesReceivedpday + $stat_rec.ReceivedTotal)/$startoffset)
$summary_stats.TotalMessagesSent = ($summary_stats.TotalMessagesSentpday + $stat_rec.SentTotal)
$summary_stats.TotalMessagesReceived = ($summary_stats.TotalMessagesReceivedpday + $stat_rec.ReceivedTotal)
$summary_stats.TotalMBSentpday = [Math]::Ceiling([decimal]($summary_stats.TotalMBSentpday + $stat_rec.SentMBTotal)/$startoffset)
$summary_stats.TotalMBReceivedpday = [Math]::Ceiling([decimal]($summary_stats.TotalMBReceivedpday + $stat_rec.ReceivedMBTotal)/$startoffset)
$summary_stats.TotalInternalSentpday = [Math]::Ceiling([decimal]($summary_stats.TotalInternalSentpday + $stat_rec.SentInternal)/$startoffset)
$summary_stats.TotalExternalSentpday = [Math]::Ceiling([decimal]($summary_stats.TotalExternalSentpday + $stat_rec.SentExternal)/$startoffset)
$summary_stats.TotalInternalSentMBpday = [Math]::Ceiling([decimal]($summary_stats.TotalInternalSentMBpday + $stat_rec.SentInternalMB)/$startoffset)
$summary_stats.TotalInternalReceivedMBpday = [Math]::Ceiling([decimal]($summary_stats.TotalInternalReceivedMBpday + $stat_rec.ReceivedInternalMB)/$startoffset)
$summary_stats.TotalExternalReceivedMBpday = [Math]::Ceiling([decimal]($summary_stats.TotalExternalReceivedMBpday + $stat_rec.ReceivedExternalMB)/$startoffset)
$summary_stats.TotalExternalSentMBpday = [Math]::Ceiling([decimal]($summary_stats.TotalExternalSentMBpday + $stat_rec.SentExternalMB)/$startoffset)
$summary_stats.TotalMBSent = $summary_stats.TotalMBSent + $stat_rec.SentMBTotal
$summary_stats.TotalMBReceived = $summary_stats.TotalMBReceived + $stat_rec.ReceivedMBTotal

} 
 
#$stat_recs | export-csv $outfile -notype  
$summary_stats.TotalInternalMBpday = $summary_stats.TotalInternalSentMBpday + $summary_stats.TotalExternalSentMBpday
$summary_stats.TotalExternalMBpday = $summary_stats.TotalExternalReceivedMBpday + $summary_stats.TotalExternalSentMBpday
$summary_stats.AverageMessageSizeKB = [Math]::Ceiling([decimal]((($summary_stats.TotalMBReceived + $summary_stats.TotalMBSent)/($summary_stats.TotalMessagesSent + $summary_stats.TotalMessagesReceived + 1))*1024))
$summary_stats.MessagesPerUserPerDay = [Math]::Ceiling([decimal](($summary_stats.TotalMessagesSentpday + $summary_stats.TotalMessagesReceivedpday)/($exch_addrs.Keys.Count + 1)))
$summary_stats.UniqueUsers = $exch_addrs.Keys.Count	

return $summary_stats

 

 
 

