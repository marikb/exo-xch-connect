$username = 'aviv'
$password = 'Ss123321'
$LocalFolder = 'C:\maap\test\mariel'
$LocalSentFolder  = ($LocalFolder.Split('\')[0..($LocalFolder.Split('\').Count - 2)] -join ('\')) + '\PreviouslySent'
if((Test-Path $LocalSentFolder) -eq $false)
{
    New-Item -ItemType Directory $LocalSentFolder
}

$RemoteFolder = 'ftp://matomo.DSAM.Test/'

$Files = Get-ChildItem $LocalFolder

foreach($file in $files)
{
    $FTPRequest = [System.Net.FtpWebRequest]::Create("$RemoteFolder/$($file.name)")
    $FTPRequest = [System.Net.FtpWebRequest]$FTPRequest
    $FTPRequest.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
    $FTPRequest.Credentials = New-Object System.Net.NetworkCredential($username, $password)
    $FTPRequest.UseBinary = $true
    $FTPRequest.UsePassive = $true

    $FileContent = Get-Content -Encoding byte $file.FullName
    $FTPRequest.ContentLength = $FileContent.Length

    $Run = $FTPRequest.GetRequestStream()
    $Run.Write($FileContent, 0, $FileContent.Length)
    # Cleanup
    $Run.Close()
    $Run.Dispose()

    $SentMessagesFolder
    Move-Item $file.FullName -Destination $LocalSentFolder
}
