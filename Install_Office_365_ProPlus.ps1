Start-Transcript "C:\Temp\SetupLogs\$($MyInvocation.MyCommand.Name).log"

New-Item -ItemType 'Directory' -Path 'C:\Temp\Office'

#Defines the contents of the SPdownload.xml
$downloadxml = "<Configuration>
  <Add SourcePath=`"c:\Temp\Office`" OfficeClientEdition=`"32`">
  <Product ID=`"O365ProPlusRetail`">
  <Language ID=`"en-us`" />
  </Product>
  </Add>
</Configuration>"

#Sends the above data to an XML file
$downloadxml | Out-File -FilePath C:\Temp\Office\SPdownload.xml -Force

#Defines the contents of the SPconfiguration.xml
$configxml = "<Configuration>
  <Add OfficeClientEdition=`"32`">
  <Product ID=`"O365ProPlusRetail`">
  <Language ID=`"en-us`" />
  </Product>
  </Add>
  <Updates Enabled=`"TRUE`" />
  <Display AcceptEULA=`"TRUE`" Level=`"None`" />
  <Logging Level=`"Standard`" Path=`"%temp%`" />
  <!--Silent install of 32-Bit Office 365 ProPlus with Updates and Logging enabled-->
</Configuration>"

#Sends the above data to an XML file
$configxml | Out-File -FilePath C:\Temp\Office\SPconfiguration.xml -Force

#Determines whether C:\Temp exists
$tempdirstatus = Test-Path -Path C:\Temp\Office

#Creates C:\Temp\Office if necessary
If ($tempdirstatus -like 'False') 
{
  #Creates C:\Temp\Office Directory
  New-Item C:\Temp\Office -type directory
}


If ($PSVersionTable.PSVersion.Major -ge '3') 
{
  $ODTLink = (((Invoke-WebRequest -Uri 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117' -UseBasicParsing).Links |
      Where-Object -FilterScript {
        $_.href -like 'https://download.microsoft.com*'
      } |
      Where-Object -FilterScript {
        $_.class -like 'mscom-link failoverLink'
  }).href)
  

  #Downloads Microsoft Office 365 Deployment Tool and saves it to C:\Temp\Office
  Invoke-WebRequest -Uri "$ODTLink" -OutFile 'C:\Temp\Office\ODT.exe'
}


Else
{
  #Saves the appropriate patch to the directory
  $url = 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11107-33602.exe' 
  $path = 'C:\Temp\Office\ODT.exe' 
  # param([string]$url, [string]$path) 
      
  if(!(Split-Path -Parent -Path $path) -or !(Test-Path -PathType Container (Split-Path -Parent -Path $path))) 
  {
    $path = Join-Path -Path $pwd -ChildPath (Split-Path -Leaf -Path $path)
  } 
      
  "Downloading [$url]`nSaving at [$path]" 
  $client = New-Object -TypeName System.Net.WebClient 
  $client.DownloadFile($url, $path) 
  #$client.DownloadData($url, $path) 
      
  $path
  &$installexe
}



#Extracts the Office 365 Deployment Tool to the C:\Temp\Office Directory   
Start-Process -FilePath 'C:\Temp\Office\ODT.exe' -ArgumentList '/extract:C:\Temp\Office /quiet /passive' -Wait

#Installs Office
Start-Process -FilePath 'C:\Temp\Office\setupodt.exe' -ArgumentList '/download C:\Temp\Office\SPdownload.xml' -Wait
Start-Process -FilePath 'C:\Temp\Office\setupodt.exe' -ArgumentList '/configure C:\Temp\Office\SPconfiguration.xml' -Wait

#Cleanup
Remove-Item -Path C:\Temp\Office -Recurse

Stop-Transcript

Exit