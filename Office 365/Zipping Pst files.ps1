
$Folder = "PST File Destenation should end with \*.pst"

$Backup = "Zip file destenation"

$compress = @{
  Path = $Folder
  CompressionLevel = "Fastest"
  DestinationPath = $Backup
}


# Closing all instances of outlook before proceeding

Get-Process Outlook |   Foreach-Object { $_.CloseMainWindow() | Out-Null } | stop-process –force

# Zipping all pst files to specified destination to one zip file

Compress-Archive @compress

#Starting outlook again

Start-Process Outlook