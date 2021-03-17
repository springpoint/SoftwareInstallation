$email = "@email@"
$apikey = "@apikey@"
$fqdn = "@fqdn@"
$installdir = 'C:\Auvik'
$outfile = 'C:\Windows\LTSvc\packages\Auvik\AuvikService.exe'
$AuvikInstallerURI = 'https://apt.my.auvik.com/native-binaries/latest-master/MinGW-i686/AuvikService.exe'

If ($fqdn -notmatch 'https://') {
    $fqdn = "https://$fqdn"
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

If ($email -and $apikey -and $fqdn) {
    $null = New-Item -Path 'C:\Windows\LTSvc\packages\' -Name 'Auvik' -ItemType Directory -Force
    Invoke-WebRequest -Uri $AuvikInstallerURI -OutFile $outfile -ErrorAction Stop
    $params = @{
        install = $true
        dir    = $installdir   
        tenant = $fqdn
        user   = $email
        password = $apikey
    }

    Start-Process $outfile @params -Wait

    
} Else {
    Write-Output 'Not all parameters were provided, please run again and provide email, apikey, and fqdn.'
}