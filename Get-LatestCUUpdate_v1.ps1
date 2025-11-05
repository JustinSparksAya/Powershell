#========Todo list========#
# - Add error handling and logging
# - Optimize module imports and checks
# - Add functionality to clean up old updates and WIMs
#=========================#


function Send-TeamsNotification {
    param (
        [string]$message
    )

     $webhookUrl = "https://ayahealthcare.webhook.office.com/webhookb2/0401bdb0-24da-4278-8e0e-ac4720a271d5@c32ce235-4d9a-4296-a647-a9edb2912ac9/IncomingWebhook/914c0121daba4aa89f0865bd4bbaaa12/9097b75c-a3c4-4eb5-ba52-f9cb324658d4/V2AOH2IPQJmbOFgwzqu-ynGzY6hlYFbWsUtOXvppuAoZ01"
     $title = "Get-LatestCUUpdate.ps1 Notification"

    $payload = @{
        title = $title
        text  = $message
    } | ConvertTo-Json

    Invoke-RestMethod -Uri $webhookUrl -Method Post -ContentType 'application/json' -Body $payload
}

function Ensure-ModuleInstalled{
     param (
          [string]$moduleName
     )
     Write-Host "[$(Get-Date -format s)] Ensuring module $moduleName is installed."
     if (-not (Get-Module -ListAvailable -Name $moduleName)) {
          Write-Host "[$(Get-Date -format s)] Module $moduleName not found. Installing..."
          Install-Module -Name $moduleName -Scope CurrentUser -Force
     }else{
          Write-Host "[$(Get-Date -format s)] Module $moduleName is already installed."
     }

     Write-Host "[$(Get-Date -format s)] Importing module $moduleName."
     Import-Module $moduleName

}

function Get-LatestCUUpdate{
     param(
          [string]$version,
          [switch]$isPreview
     )

     $year = (Get-Date).Year.ToString()
     $month = (Get-Date).ToString("MM")
     Write-Host "[$(Get-Date -format s)] Searching for latest CU update for Windows 11 version $version for $year-$month. Preview: $isPreview"
     if($isPreview){
          Write-Host "[$(Get-Date -format s)] Searching for Preview updates."
          $availableUpdates = Get-MSCatalogUpdate -Search "$year-$month Cumulative Update Preview for Windows 11, version $version x64" -LastDays 60 -AllPages -IncludePreview | Where-Object {$_.Title -notmatch ".NET Framework" -and $_.Title -match "x64"}
     }else{
          Write-Host "[$(Get-Date -format s)] Searching for non-Preview updates."
          $availableUpdates = Get-MSCatalogUpdate -Search "$year-$month Cumulative Update for Windows 11, version $version x64" -LastDays 60 -AllPages  | Where-Object {$_.Title -notmatch ".NET Framework" -and $_.Title -match "x64"}
     }

     Write-Host "[$(Get-Date -format s)] Found $($availableUpdates.Count) updates for version $version. Returning the latest one."
     $availableUpdates | Sort-Object LastUpdated -Descending | Select-Object -First 1
}


try {
Start-Transcript -Path "$PSScriptRoot\Get-LatestCUUpdate.log" -Append
Ensure-ModuleInstalled -moduleName "DISM"
Ensure-ModuleInstalled -moduleName "MSCatalogLTS"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

$versions = @("25H2")

foreach($version in $versions){
    
    $latestUpdate = Get-LatestCUUpdate -version $version -isPreview:$false
    $latestPreviewUpdate = Get-LatestCUUpdate -version $version -isPreview:$true

    #extract the update id from the title, string in first parentheses
    $latestupdateId = (($latestUpdate.Title -split "\(")[1] -split "\)")[0]
    $latestPreviewUpdateId = (($latestPreviewUpdate.Title -split "\(")[1] -split "\)")[0]

    $lastUpdateDate = $null
    $lastPreviewUpdateDate = $null
    $lastUpdateDate = $latestUpdate.LastUpdated.ToString("yyyy-MM-dd")
    $lastPreviewUpdateDate = $latestPreviewUpdate.LastUpdated.ToString("yyyy-MM-dd")
  

   if(-not (Get-ChildItem -Path $downloadLocation | Where-Object {$_.Name -match $version -and $_.Name -match $latestupdateId -and $_.Name -notmatch "Preview"})){
        Write-Host "[$(Get-Date -format s)] Downloading $($latestUpdate.Title) to $scriptRoot"
        $latestUpdate | Save-MSCatalogUpdate -Destination $scriptRoot
        start-sleep -s 5
        Rename-Item -Path "$scriptRoot\windows11.0-$($latestupdateId)-x64.msu" -NewName "W11-$version-CU-$($lastPReviewUpdateDate)-$($latestupdateId)-x64.msu"
   }else{
    Write-Host "[$(Get-Date -format s)] $($latestUpdate.Title) already downloaded."
   }


   if(-not (Get-ChildItem -Path $scriptRoot | Where-Object {$_.Name -match $version -and $_.Name -match $latestPreviewUpdateId -and $_.Name -match "Preview"})){
        Write-Host "[$(Get-Date -format s)] Downloading $($latestPreviewUpdate.Title) to $scriptRoot"
        $latestPreviewUpdate | Save-MSCatalogUpdate -Destination $scriptRoot
        start-sleep -s 5
        Rename-Item -Path "$scriptRoot\windows11.0-$($latestPreviewUpdateId)-x64.msu" -NewName "W11-$version-CU-Preview-$($lastPReviewUpdateDate)-$($latestPreviewUpdateId)-x64.msu"
   }else{
    Write-Host "[$(Get-Date -format s)] $($latestPreviewUpdate.Title) already downloaded."
   }

   $allVersionDownloads = Get-ChildItem -Path $scriptRoot | Where-Object {$_.Name -match  $version -and $_.Name -notmatch "Preview"}
   $allPreviewVersionDownloadss = Get-ChildItem -Path $scriptRoot | Where-Object {$_.Name -match  $version -and $_.Name -match "Preview"}

   if($allVersionDownloads.Count -gt 2){
        $filesToRemove = $allVersionDownloads | Sort-Object LastWriteTime -Descending | Select-Object -Skip 2
        foreach($file in $filesToRemove){
            Write-Host "[$(Get-Date -format s)] Removing old update file: $($file.Name)"
            #Remove-Item -Path $file.FullName -Force
        }
   }
   if($allPreviewVersionDownloadss.Count -gt 2){
        $previewFilesToRemove = $allPreviewVersionDownloadss | Sort-Object LastWriteTime -Descending | Select-Object -Skip 2
        foreach($file in $previewFilesToRemove){
            Write-Host "[$(Get-Date -format s)] Removing old preview update file: $($file.Name)"
            #Remove-Item -Path $file.FullName -Force
        }
   }

   #Check if the WIM has been created for this version
   $baseWimPath = "$scriptRoot\WIM\base\W11-$version.wim"
   $previewbaseWimPath = "$scriptRoot\WIM\base\W11-$version-preview.wim"
   $wimPath = Get-ChildItem -Path "$scriptRoot\WIM" | Where-Object {$_.Name -match $version -and $_.Name -match $latestUpdateId -and $_.Name -notmatch "Preview"}
   $previewWimPath = Get-ChildItem -Path "$scriptRoot\WIM" | Where-Object {$_.Name -match $version -and $_.Name -match $latestPreviewUpdateId -and $_.Name -match "Preview"}
   $isoPath = Get-ChildItem -Path "$scriptRoot\ISO" | Where-Object {$_.Name -match $version} | Sort-Object LastWriteTime -Descending | Select-Object -First 1
   $msuUpdatePath = Get-ChildItem -Path $scriptRoot | Where-Object {$_.Name -match $version -and $_.Name -match $latestupdateId -and $_.Name -notmatch "Preview"} | Select -ExpandProperty FullName
   $previewmsuUpdatePath = Get-ChildItem -Path $scriptRoot | Where-Object {$_.Name -match $version -and $_.Name -match $latestPreviewUpdateId -and $_.Name -match "Preview"} | Select -ExpandProperty FullName
   if(-not (Test-Path "$env:Temp\WIMmount")){
        New-Item -Path "$env:Temp\WIMmount" -ItemType Directory | Out-Null
   }

   if(-not (Test-Path -Path $baseWimPath)){
     $img = Mount-DiskImage -ImagePath $isoPath.FullName -PassThru
     Start-Sleep -Seconds 2
     $vol = Get-Volume -DiskImage $img
     $driveLetter = "$($vol.DriveLetter):"
     $wimOutputPath = $baseWimPath
     Write-Host "[$(Get-Date -format s)] Creating WIM base: $wimOutputPath"
     Export-WindowsImage -SourceImagePath "$driveLetter\sources\install.wim" -SourceIndex 3 -DestinationImagePath $wimOutputPath -CheckIntegrity -CompressionType Max
     start-sleep -s 2
     Dismount-DiskImage -ImagePath $isoPath.FullName | Out-Null
   }

   if(-not (Test-Path -Path $previewbaseWimPath)){
     Write-Host "[$(Get-Date -format s)] Creating preview base WIM: $previewbaseWimPath"
     Copy-Item -Path $baseWimPath -Destination $previewbaseWimPath -Force
   }
 
   if(-not $wimPath){
        #Create new WIM with latest CU
        $wimOutputPath = "$scriptRoot\WIM\W11-$version-CU-$($lastUpdateDate)-$($latestupdateId).wim"
        Copy-Item -Path $baseWimPath -Destination $wimOutputPath -Force
        start-sleep -s 2
        Write-Host "[$(Get-Date -format s)] Applying $($latestUpdate.Title) to WIM: $wimOutputPath"
        Mount-WindowsImage -ImagePath $wimOutputPath -Index 1 -Path "$env:Temp\WIMmount"
        Write-Host "[$(Get-Date -format s)] Adding CU package: $msuUpdatePath"
        Add-WindowsPackage -Path "$env:Temp\WIMmount" -PackagePath $msuUpdatePath -IgnoreCheck 
        Write-Host "[$(Get-Date -format s)] Finalizing WIM with CU: $wimOutputPath"
        Dismount-WindowsImage -Path "$env:Temp\WIMmount" -Save
        Copy-Item -Path $wimOutputPath -Destination "$scriptRoot\WIM\base\W11-$version.wim" -Force
        Write-Host "[$(Get-Date -format s)] WIM with latest CU created: $wimOutputPath"
        $teamsMessage = "A new WIM with the latest CU has been created: $wimOutputPath"
        Send-TeamsNotification -message $teamsMessage
}    else{
    Write-Host "[$(Get-Date -format s)] WIM with latest CU already exists: $($wimPath.FullName)"
}   

   if(-not $previewWimPath){
        #Create new WIM with latest Preview CU
        $wimOutputPath = "$scriptRoot\WIM\W11-$version-CU-Preview-$($lastPreviewUpdateDate)-$($latestPreviewUpdateId).wim"
        Copy-Item -Path $previewbaseWimPath -Destination $wimOutputPath -Force
        start-sleep -s 2
        Write-Host "[$(Get-Date -format s)] Applying $($latestPreviewUpdate.Title) to WIM: $wimOutputPath"
        Mount-WindowsImage -ImagePath $wimOutputPath -Index 1 -Path "$env:Temp\WIMmount"
        Write-Host "[$(Get-Date -format s)] Adding Preview CU package: $previewmsuUpdatePath"
        Add-WindowsPackage -Path "$env:Temp\WIMmount" -PackagePath $previewmsuUpdatePath -IgnoreCheck
        write-Host "[$(Get-Date -format s)] Finalizing WIM with Preview CU: $wimOutputPath"
        Dismount-WindowsImage -Path "$env:Temp\WIMmount" -Save
        Copy-Item -Path $wimOutputPath -Destination "$scriptRoot\WIM\base\W11-$version-preview.wim" -Force
        Write-Host "[$(Get-Date -format s)] WIM with latest Preview CU created: $wimOutputPath"
        $teamsMessage = "A new WIM with the latest Preview CU has been created: $wimOutputPath"
        Send-TeamsNotification -message $teamsMessage
   }
}
} catch {
     if(Test-Path -Path "$env:Temp\WIMmount"){
          Dismount-WindowsImage -Path "$env:Temp\WIMmount" -Discard
     }
    Write-Host "[$(Get-Date -format s)] An error occurred: $_"
    $teamsMessage = "An error occurred while processing the latest CU update: $_ Line: $($_.InvocationInfo.ScriptLineNumber)"
    Send-TeamsNotification -message $teamsMessage

} finally {
     Stop-Transcript
}

