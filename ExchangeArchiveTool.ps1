[Void]
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
# import xaml function used for all wpf forms
try {
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline
	[xml]$Global:Xaml2 = Get-Content -Path "$PSScriptRoot\archiveTool2.xaml"
	$Reader2 = New-Object System.Xml.XmlNodeReader $Xaml2
    $Global:Window2 = [Windows.Markup.XamlReader]::Load($Reader2)
}
catch{
    Write-Error "Error building Xaml data.`n$_"
    exit 
}
$Xaml2.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name "WPF$($_.Name)" -Value $Window2.FindName($_.Name)}
$Window.Close()

$WPFarchiveDisabled.Visibility = "Hidden"
$WPFarchiveEnabled.Visibility = "Hidden"
$WPFdisableArch.Visibility = "Hidden"
$WPFenableArch.Visibility = "Hidden"
$WPFarchiveToggle.Visibility = "Hidden"
$WPFarchiveToggle.IsEnabled ="False"

$WPFprocessButton.Add_Click({
if($WPFclientEmail -ne $null){
$Global:archiveStatus = Get-Mailbox -identity $WPFclientEmail.Text | Select-object -ExpandProperty ArchiveStatus
if ($archiveStatus -eq "Active") {
    $WPFarchiveEnabled.Visibility = "Visible"
    $WPFarchiveDisabled.Visibility = "Hidden"
    $WPFdisableArch.Visibility = "Visible"
    $WPFenableArch.Visibility = "Hidden"
    $WPFarchiveToggle.Visibility = "Visible"
    $WPFarchiveToggle.IsEnabled = "True"
}else{
    $WPFarchiveEnabled.Visibility = "Hidden"
    $WPFarchiveDisabled.Visibility = "Visible"
    $WPFdisableArch.Visibility = "Hidden"
    $WPFenableArch.Visibility = "Visible"
    $WPFarchiveToggle.Visibility = "Visible"
    $WPFarchiveToggle.IsEnabled = "True"
}
if ($WPFretentionToggle.IsChecked -eq "True") {
    Start-ManagedFolderAssistant -Identity $WPFclientEmail.Text
    }else{}
if ($WPFrulesToggle.IsChecked -eq "True") {
        Write-Host "RulesWorking"
        $Log = Export-MailboxDiagnosticLogs -Identity $WPFclientEmail.Text -ExtendedProperties
        $xml = [xml]($Log.MailboxLog)
        $Data= $xml.Properties.MailboxTable.Property | ? {$_.Name -like "ELC*"}
        $WPFOutput.Text = $Data | Out-String
    }else{}

if ($WPFenableArch.Visibility -eq "Visible" -and $WPFarchiveToggle.IsChecked -eq "True" ) {
    Enable-Mailbox -Identity $WPFclientEmail.Text -Archive
}elseif ($WPFenableArch.Visibility -eq "Hidden" -and $WPFarchiveToggle.IsChecked -eq "True" ) {
    Disable-Mailbox -Identity $WPFclientEmail.Text -Archive -Confirm:$false
}else{}
}else{}

$WPFarchiveToggle.IsChecked = "False"
$WPFretentionToggle.IsChecked = "False"
$WPFrulesToggle.IsChecked = "False"
})

#used to display size of archive folder if active. 
$WPFarchiveSizeButton.Add_Click({
    if($WPFclientEmail -ne $null){
        $archiveSizeResult=@() 
        $mailboxes = Get-Mailbox -Identity $WPFclientEmail.Text -ResultSize Unlimited
        $totalmbx = $mailboxes.Count
        $i = 1 
        $mailboxes | ForEach-Object {
        $i++
        $mbx = $_
        $size = $null
         
        Write-Progress -activity "Processing $mbx" -status "$i out of $totalmbx completed"
         
        if ($mbx.ArchiveStatus -eq "Active"){
        $mbs = Get-MailboxStatistics $mbx.UserPrincipalName -Archive
         
        if ($mbs.TotalItemSize -ne $null){
        $size = [math]::Round(($mbs.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',','')/1GB),2)
        }else{
        $size = 0 }
        }
         
        $archiveSizeResult += New-Object -TypeName PSObject -Property $([ordered]@{ 
        UserName = $mbx.DisplayName
        UserPrincipalName = $mbx.UserPrincipalName
        ArchiveStatus =$mbx.ArchiveStatus
        ArchiveName =$mbx.ArchiveName
        ArchiveState =$mbx.ArchiveState
        ArchiveMailboxSizeInGB = $size
        ArchiveWarningQuota=if ($mbx.ArchiveStatus -eq "Active") {$mbx.ArchiveWarningQuota} Else { $null} 
        ArchiveQuota = if ($mbx.ArchiveStatus -eq "Active") {$mbx.ArchiveQuota} Else { $null} 
        AutoExpandingArchiveEnabled=$mbx.AutoExpandingArchiveEnabled
        })
        }
        $WPFOutput.Text = $archiveSizeResult | Out-String
    }else{}
})

$WPFmailSizeButton.Add_Click({
    if($WPFclientEmail -ne $null){
        $mailBoxResult=@() 
        $mailboxes = Get-Mailbox  -Identity $WPFclientEmail.Text -ResultSize Unlimited
        $totalmbx = $mailboxes.Count
        $i = 1 
        $mailboxes | ForEach-Object {
        $i++
        $mbx = $_
        $mbs = Get-MailboxStatistics $mbx.UserPrincipalName
          
        Write-Progress -activity "Processing $mbx" -status "$i out of $totalmbx completed"
          
        if ($mbs.TotalItemSize -ne $null){
        $size = [math]::Round(($mbs.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',','')/1GB),2)
        }else{
        $size = 0 }
         
        $mailBoxResult += New-Object PSObject -property @{ 
        Name = $mbx.DisplayName
        UserPrincipalName = $mbx.UserPrincipalName
        TotalSizeInGB = $size
        SizeWarningQuota=$mbx.IssueWarningQuota
        StorageSizeLimit = $mbx.ProhibitSendQuota
        StorageLimitStatus = $mbs.ProhibitSendQuota
        }
        }
        $WPFOutput.Text = $mailBoxResult | Out-String
    }else{}
})

$Window2.ShowDialog() | Out-Null

$WPFcancelButton.Add_Click({
    Disconnect-ExchangeOnline -Confirm:$false
})

$Window.ShowDialog() | out-null