# remote
$dvd = "C:\Temp\Dynamics.365.BC.18056.IT.DVD"
$dvdVersion = "240"

$upgradeDev = "$dvd\ModernDev\program files\Microsoft Dynamics NAV\$dvdVersion\AL Development Environment"
$upgradeWeb = "$dvd\WebClient\Microsoft Dynamics NAV\$dvdVersion\Web Client"
$upgradeService = "$dvd\ServiceTier\program files\Microsoft Dynamics NAV\$dvdVersion\Service"
$upgradeWebPublish = "$dvd\WebClient\Microsoft Dynamics NAV\$dvdVersion\Web Client\WebPublish"

#local
$targetDev = "C:\Program Files (x86)\Microsoft Dynamics 365 Business Central\240\AL Development Environment"
$targetWeb = "C:\Program Files\Microsoft Dynamics 365 Business Central\240\Web Client"
$targetSites = ,"C:\inetpub\wwwroot\WDEV"
$targetService = "C:\Program Files\Microsoft Dynamics 365 Business Central\240\Service"

$mgmtPath = "C:\Program Files\Microsoft Dynamics 365 Business Central\240\Service\Admin"
$instance = "WDEV"

# DVD provided apps
$apps = "ModernDev\program files\Microsoft Dynamics NAV\240\AL Development Environment\System.app",
		"Applications\system application\source\Microsoft_System Application.app",
		"Applications\BusinessFoundation\source\Microsoft_Business Foundation.app",
		"Applications\BaseApp\Source\Microsoft_Base Application.app",
		"Applications\Application\Source\Microsoft_Application.app"

Import-Module "$mgmtPath\NavAdminTool.ps1"

Write-Host ""
Write-Host "1) Uninstall all apps"
Write-Host "2) Publish and install new Microsoft apps"
Write-Host "3) Clean up old Microsoft apps"
Write-Host "4) Re-install other apps"
Write-Host ""
Write-Host "C) Convert database"
Write-Host "F) Fix addins"
Write-Host "N) App status"
Write-Host "S) Sync tenant"
Write-Host "U) Service upgrade"
Write-Host ""

$choice = Read-Host "Enter step number"

if (("n", "N").contains($choice))
{
	Get-NAVAppInfo -ServerInstance $instance -TenantSpecificProperties -Tenant default | Format-Table -Property Name, Publisher, Version, IsInstalled
	exit
}

if (("c", "C").contains($choice))
{
	[xml]$serviceConf = Get-Content $targetService\Instances\$instance\CustomSettings.config 
	
	$node = Select-Xml -Xml $serviceConf -XPath "/appSettings/add[@key='DatabaseServer']"
	$dbServer = $node.Node.value.ToString()
	
	$node = Select-Xml -Xml $serviceConf -XPath "/appSettings/add[@key='DatabaseName']"
	$dbName = $node.Node.value.ToString()
	
	Invoke-NAVApplicationDatabaseConversion -DatabaseName $dbName -DatabaseServer $dbServer
	exit 
}

if (("s", "S").contains($choice))
{
	Sync-NAVTenant -ServerInstance $instance -Mode Sync
	exit
}

if (("u", "U").contains($choice))
{
	Copy-Item $upgradeDev\* $targetDev -Recurse -Force
	Copy-Item $upgradeWeb\* $targetWeb -Recurse -Force

	$serviceConf = Get-Content $targetService\CustomSettings.config
	Copy-Item $upgradeService\* $targetService -Recurse -Force	
	Set-Content -Path $targetService\CustomSettings.config -Value $serviceConf	

	foreach ($ts in $targetSites) {
		$webConf = Get-Content $ts\NavSettings.json
		Copy-Item $upgradeWebPublish\* $ts -Recurse -Force
		Set-Content -Path $ts\NavSettings.json -Value $webConf
	}

	exit
}

if (("f", "F").contains($choice))
{
	Set-NAVAddIn -ServerInstance $instance -AddinName 'Microsoft.Dynamics.Nav.Client.BusinessChart' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $targetService 'Add-ins\BusinessChart\Microsoft.Dynamics.Nav.Client.BusinessChart.zip')
	Set-NAVAddIn -ServerInstance $instance -AddinName 'Microsoft.Dynamics.Nav.Client.FlowIntegration' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $targetService 'Add-ins\FlowIntegration\Microsoft.Dynamics.Nav.Client.FlowIntegration.zip')
	Set-NAVAddIn -ServerInstance $instance -AddinName 'Microsoft.Dynamics.Nav.Client.OAuthIntegration' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $targetService 'Add-ins\OAuthIntegration\Microsoft.Dynamics.Nav.Client.OAuthIntegration.zip')
	Set-NAVAddIn -ServerInstance $instance -AddinName 'Microsoft.Dynamics.Nav.Client.PageReady' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $targetService 'Add-ins\PageReady\Microsoft.Dynamics.Nav.Client.PageReady.zip')
	Set-NAVAddIn -ServerInstance $instance -AddinName 'Microsoft.Dynamics.Nav.Client.PowerBIManagement' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $targetService 'Add-ins\PowerBIManagement\Microsoft.Dynamics.Nav.Client.PowerBIManagement.zip')
	Set-NAVAddIn -ServerInstance $instance -AddinName 'Microsoft.Dynamics.Nav.Client.RoleCenterSelector' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $targetService 'Add-ins\RoleCenterSelector\Microsoft.Dynamics.Nav.Client.RoleCenterSelector.zip')
	Set-NAVAddIn -ServerInstance $instance -AddinName 'Microsoft.Dynamics.Nav.Client.SatisfactionSurvey' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $targetService 'Add-ins\SatisfactionSurvey\Microsoft.Dynamics.Nav.Client.SatisfactionSurvey.zip')
	Set-NAVAddIn -ServerInstance $instance -AddinName 'Microsoft.Dynamics.Nav.Client.VideoPlayer' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $targetService 'Add-ins\VideoPlayer\Microsoft.Dynamics.Nav.Client.VideoPlayer.zip')
	Set-NAVAddIn -ServerInstance $instance -AddinName 'Microsoft.Dynamics.Nav.Client.WebPageViewer' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $targetService 'Add-ins\WebPageViewer\Microsoft.Dynamics.Nav.Client.WebPageViewer.zip')
	Set-NAVAddIn -ServerInstance $instance -AddinName 'Microsoft.Dynamics.Nav.Client.WelcomeWizard' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $targetService 'Add-ins\WelcomeWizard\Microsoft.Dynamics.Nav.Client.WelcomeWizard.zip')

	exit
}

$o1 = Get-NAVAppInfo -ServerInstance $instance -TenantSpecificProperties -Tenant default
$o2 = Get-NAVAppInfo -ServerInstance $instance -SymbolsOnly
$oldApps = $o1 + $o2

if ($choice -eq "1")
{
	foreach ($old in $oldApps)
	{
		if ($old.IsInstalled)
		{
			Write-Output "Uninstalling $($old.Name)"
			Uninstall-NAVApp -ServerInstance $instance -Name $old.Name -Version $old.Version -Publisher $old.Publisher -Force
		}
	}
}

if ($choice -eq "2")
{
	# publish
	foreach ($a in $apps)
	{
		$nfo = Get-NAVAppInfo -Path "$dvd\$a"
		
		if ($nfo.Name -eq "System")
		{
			Write-Output "Publishing $($nfo.Name)"
     		Publish-NAVApp -ServerInstance $instance -Path "$dvd\$a" -PackageType SymbolsOnly
		}
		else
		{
			Write-Output "Publishing $($nfo.Name)"
			Publish-NAVApp -ServerInstance $instance -Path "$dvd\$a"
		}
	}
	
	#sync
	foreach ($a in $apps)
	{
		$nfo = Get-NAVAppInfo -Path "$dvd\$a"
		$nfo = Get-NAVAppInfo -ServerInstance $instance -Name $nfo.Name -Publisher $nfo.Publisher -Version $nfo.Version -TenantSpecificProperties -Tenant default
		if ($nfo.Count -eq 0)
		{
			continue
		}
			
		$nfo = $nfo[0]

		Write-Output "Syncing $($nfo.Name)"
		Sync-NAVApp -ServerInstance $instance -Name $nfo.Name -Publisher $nfo.Publisher -Version $nfo.Version
	}
	
	Sync-NAVTenant -ServerInstance $instance -Mode Sync -Force
	Start-NavDataUpgrade -ServerInstance $instance -FunctionExecutionMode Serial -Force
	
	#data upgrade
	foreach ($a in $apps)
	{
		$nfo = Get-NAVAppInfo -Path "$dvd\$a"
		$nfo = Get-NAVAppInfo -ServerInstance $instance -Name $nfo.Name -Publisher $nfo.Publisher -Version $nfo.Version -TenantSpecificProperties -Tenant default
		if ($nfo.Count -eq 0)
		{
			continue
		}
			
		$nfo = $nfo[0]

		Write-Output "Data upgrading $($nfo.Name)"
		Start-NAVAppDataUpgrade -ServerInstance $instance -Name $nfo.Name -Publisher $nfo.Publisher -Version $nfo.Version
	}
	
	#install
	foreach ($a in $apps)
	{
		$nfo = Get-NAVAppInfo -Path "$dvd\$a"
		$nfo = Get-NAVAppInfo -ServerInstance $instance -Name $nfo.Name -Publisher $nfo.Publisher -Version $nfo.Version -TenantSpecificProperties -Tenant default
		if ($nfo.Count -eq 0)
		{
			continue
		}
			
		$nfo = $nfo[0]
		
		Write-Output "Installing $($nfo.Name)"
		Install-NAVApp -ServerInstance $instance -Name $nfo.Name -Version $nfo.Version -Publisher $nfo.Publisher -Force
	}
	
	#application Version
	foreach ($a in $apps)
	{
		$nfo = Get-NAVAppInfo -Path "$dvd\$a"

		if ($nfo.Name -eq "Base Application")
		{
			Write-Output "Setting application version to $($nfo.Version)"
			Set-NAVApplication -ServerInstance $instance -Force -ApplicationVersion $nfo.Version
		}
	}
	
	Sync-NAVTenant -ServerInstance $instance -Mode Sync -Force
	Start-NavDataUpgrade -ServerInstance $instance -FunctionExecutionMode Serial -Force
}

if ($choice -eq "3")
{
	foreach ($a in $apps)
	{
		$nfo = Get-NAVAppInfo -Path "$dvd\$a"
		
		foreach ($old in $oldApps)
		{
			if (($nfo.Name -eq $old.Name) -and ($nfo.Publisher -eq $old.Publisher) -and ($nfo.Version -gt $old.Version))
			{
				Write-Output "Unpublishing $($nfo.Name)"
				Unpublish-NAVApp -ServerInstance $instance -Name $old.Name -Publisher $old.Publisher -Version $old.Version
			}
		}
	}	
}

if ($choice -eq "4")
{
	foreach ($old in $oldApps)
	{
		if (($old.IsInstalled -ne $True) -and ($old.SyncState -ne $null) -and ($old.SyncState.ToString() -eq "Synced"))
		{
			Write-Output "Installing $($old.Name)"
			Install-NAVApp -ServerInstance $instance -Name $old.Name -Version $old.Version -Publisher $old.Publisher -Force
		}
	}
}
