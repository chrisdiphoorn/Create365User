if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -WindowStyle hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

write-host $Env:UserName

$DebugPreference = 'Continue'   		#Turn On Debug
#$DebugPreference = 'SilentlyContinue'   #Turn Off Debug

#Update These Values if Required to Run a Different Script or Change the Source Location
$FileName = 'Create365User'
	
Function Get-CurrentPath {
	$currentPath = $PSScriptRoot                                                                                                     # AzureDevOps, Powershell
	if (!$currentPath) { $currentPath = Split-Path $pseditor.GetEditorContext().CurrentFile.Path -ErrorAction SilentlyContinue }     # VSCode
	if (!$currentPath) { $currentPath = Split-Path $psISE.CurrentFile.FullPath -ErrorAction SilentlyContinue }                       # PsISE
	return $currentPath + '\'
}

$Default = $null
$cfolder = Get-CurrentPath 

$CurrentFolder = Get-CurrentPath 
$IniFile = "$($CurrentFolder)$($FileName).ini"

IF([System.IO.File]::Exists($IniFile) -eq $true) {
	write-Debug "Reading INI File '$($IniFile)'"
	$Default = Get-Content $IniFile | ConvertFrom-StringData
}
 
$PFXFile = $null
$SPOUser = $null

if($Default) {

		$RunPassword = $Default.ActiveDirectoryPassword
		$SPOUser = $Default.ConnectSPOServiceUser
		
		if($SPOUser) {
			$PFXFile = "$($cFolder)$($SPOUser).pfx"
		}

		$secureRunPassword = ConvertTo-SecureString $RunPassword -AsPlainText -Force

		#Ensure that the APP Certificate has been installed. This is needed my Exchange-Online
		$FindCert = (Get-ChildItem -Path Cert:\LocalMachine\my| Where-Object {$_.Subject -eq "CN=$SPOUser"})
		if (!$FindCert) {
				IF([System.IO.File]::Exists($PFXFile) -eq $true) {
					write-host "Installing Certificate from $($PFXFile)"
					try {
							Import-PfxCertificate -FilePath $PFXFile -CertStoreLocation Cert:\LocalMachine\My -Password $secureRunPassword
					} catch {
						write-host $_.Exception.Message
						pause
						exit 1
					}
				} else {
					Write-host "Missing Application Certificate File '$PFXFile' (CN=$SPOUser)"
					pause
					exit 1
				}
		} else {
				write-host "Missing ActiveDirectoryPassword"
				Write-host "Missing Application Certificate File '$PFXFile'"
				pause
				exit 1
		}
}

pause
