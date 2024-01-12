<#
.SYNOPSIS
Deploy any pending update using Windows Update Agent (WUA) API

.DESCRIPTION
Automatically download and install any pending update using the native WUA API.

.EXAMPLE
WUT.PS1 -Reboot

.PARAMETER Reboot
Automatically reboot when it is needed to complete the installation of one or more updates.

.PARAMETER SearchOnly
Search for updates and exit.

.PARAMETER DownloadOnly
Search and Download updates, and exit.

.PARAMETER MillisecondsDelay
Milliseconds of delay while waiting for progress

.PARAMETER RebootDelaySeconds
Seconds to wait before request a computer restart.

#>

[cmdletbinding()]
Param(
  [Parameter(Position = 1)]
  [switch]$Reboot,
  [Parameter(Position = 2)]
  [int]$RebootDelaySeconds = 10,
  [Parameter(Position = 3)]
  [switch]$SearchOnly,
  [Parameter(Position = 4)]
  [switch]$DownloadOnly,
  [Parameter(Position = 5)]
  [switch]$ResetWindowsUpdate,
  [Parameter(Position = 6)]
  [switch]$ShowUpdateHistory  
)

[string]$WsusServer = (Get-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate").WUServer
[int]$UseWUServer = (Get-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU").UseWUServer
[string]$WUTargetGroup = (Get-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate").TargetGroup
[int]$UseWUTargetGroup = (Get-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate").TargetGroupEnabled
[int]$MillisecondsDelay = 100
[string]$LocalDeviceName = ([system.net.dns]::GetHostByName('localhost')).hostname

Write-Host "Device Name: " -NoNewline
Write-Host -ForegroundColor Yellow $LocalDeviceName

$PSVersionText = $PSVersionTable.PSVersion.ToString()
Write-Host "PowerShell Version: " -NoNewline
Write-Host -ForegroundColor Yellow $PSVersionText 


function FormattedSize ($DataSize) {
  $GigaBytes = [Math]::Round($DataSize / (1024 * 1024 * 1024), 2)
  $MegaBytes = [Math]::Round($DataSize / (1024 * 1024), 2)
  $KiloBytes = [Math]::Round($DataSize / 1024, 2)
  $Bytes = [Math]::Round($DataSize, 2)

  if ($GigaBytes -gt 1) {
    $ResultString = $GigaBytes.ToString() + " GB" 
  }
  elseif ($MegaBytes -gt 1) {
    $ResultString = $MegaBytes.ToString() + " MB" 
  }
  elseif ($KiloBytes -gt 1) {
    $ResultString = $KiloBytes.ToString() + " KB" 
  }
  else {
    $ResultString = $Bytes.ToString() + " B" 
  }

  return $ResultString
}


function DoRebootIfNeededAndAllowed ([int]$Seconds2Reboot) {
  # Reboot if needed and allowed
  If ((New-Object -ComObject Microsoft.Update.SystemInfo).RebootRequired) {
    Write-Host "This computer needs to be restarted to complete the installation of one or more updates."
    If ($Reboot) {
      Write-Host "Restarting computer in " $Seconds2Reboot " seconds!"
      Start-Sleep -Seconds $Seconds2Reboot
      Restart-Computer
    }
    Exit # Exit if there is a pending reboot
  }
}


function DoDisplayLine ([string]$Line2Display) {
  Write-Host $Line2Display -NoNewline
  Write-Host "`r" -NoNewline
  #[System.Console]::SetCursorPosition(0, [System.Console]::CursorTop) # This fails when using PSExec
}

function IsCompleted {
  switch ($OperationID) {
    0 { $SearchJob.IsCompleted }
    1 { $DownloadJob.IsCompleted }
    2 { $InstallJob.IsCompleted }
  }  
}

function DoReport2WUServer {
  if (($UseWUServer -eq 1) -and ($UseWUTargetGroup -eq 1)) {   # If a WSUS is in use, request report to WSUS server
    & "C:\Windows\System32\wuauclt.exe" /reportnow
  }  
}

function DoShowErrorAndExit ($Stage, $ResultCode, $HResultCode) {
  $HexResult = ('0x{0:x}h' -f $HResultCode)
  Write-Host $Stage "ResultCode: ["$ResultCode"]" -NoNewline
  switch ($ResultCode) {
    0 { Write-Host "The operation has not started." }
    1 { Write-Host "The operation is in progress."}
    2 { Write-Host "The operation was completed successfully."}
    3 { Write-Host "One or more errors occurred during the operation. The results might be incomplete."}
    4 { Write-Host "The operation failed to complete."}
    5 { Write-Host "The operation was canceled."}
    Default { Write-Host "Unexpected value for ResultCode!" }
  }
  Write-Host "($HexResult): " (New-Object System.ComponentModel.Win32Exception($HResultCode)).Message
  exit
}

# If a WSUS server is in use, exit if it is not responding


If ($UseWUServer -eq 1) {
  Write-Host "Connecting to WSUS server: " -NoNewline
  Write-Host -ForegroundColor Yellow $WsusServer
  $HttpWebResponse = ([System.Net.HttpWebRequest]::Create($WsusServer)).GetResponse();
  if ($HttpWebResponse) {
    If ($HttpWebResponse.StatusCode -ne 200) {
      Write-Host "WSUS server is not responding correctly. Error code: " -NoNewline
      Write-Host -Object $HttpWebResponse.StatusCode.value__;
      #Write-Host -Object $HttpWebResponse.GetResponseHeader("X-Detailed-Error");
      exit
    }
  }
  else {
    Exit
  }
}
if ($UseWUTargetGroup -eq 1) {
  Write-Host "WSUS group: " -NoNewline
  Write-Host -ForegroundColor Yellow $WUTargetGroup
}


# Create the main COM objects

$UnusedCallback = New-Object -ComObject Microsoft.Update.UpdateColl

$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSession.ClientApplicationID = "Windows Update Tool"

while ($true) {
  DoReport2WUServer 
  DoRebootIfNeededAndAllowed($RebootDelaySeconds)
  For ($OperationID=0; $OperationID -le 2; $OperationID++) { #Loop through the three stages: 0=Search, 1=Download and 2=Install
    switch ($OperationID) { # Begining of current operation
      0 { #BeginSearch
        $Searcher = $UpdateSession.CreateUpdateSearcher()
        $SearchJob = $Searcher.BeginSearch("IsInstalled=0", $UnusedCallback, $null) # Search for updates
      } 
      1 { #BeginDownload
        $Downloader = $UpdateSession.CreateUpdateDownloader()
        $Downloader.Updates = $UpdatesToDownload
        $UpdatesCount = $Downloader.Updates.Count
        $DownloadJob = $Downloader.BeginDownload($UnusedCallback, $UnusedCallback, $null) # Begin downloading the updates
      } 
      2 { #BeginInstall
        $Installer = $UpdateSession.CreateUpdateInstaller()
        $Installer.Updates = $UpdatesToInstall
        $UpdatesCount = $Installer.Updates.Count
        $InstallJob = $Installer.BeginInstall($UnusedCallback, $UnusedCallback, $null) # Begin installing the updates
      } 
    } 

    while (-not (IsCompleted)) { # Loop while the current operation is in progress
      switch ($SpinChar) {
        "-" { $SpinChar = "\" }
        "\" { $SpinChar = "|" }
        "|" { $SpinChar = "/" }
        Default { $SpinChar = "-" }
      }

      switch ($OperationID) {
        0 { #SearchInProgress
            $ProgressText = "Searching for updates... "
        } 
        1 { #DownloadInProgress
          $UpdateProgress = $DownloadJob.GetProgress()
          $ProgressText = "Downloading"
        } 
        2 { #InstallInProgress
          $UpdateProgress = $InstallJob.GetProgress()
          $ProgressText = "Installing"
        } 
      } 
      if (($OperationID -eq 1) -or ($OperationID -eq 2)) {
        $CurrentUpdateIndex = $UpdateProgress.CurrentUpdateIndex + 1
        $ProgressText += " " + $CurrentUpdateIndex + " of " + $UpdatesCount
        $ProgressText += " [" + $UpdateProgress.CurrentUpdatePercentComplete + "%] "
      }
      $ProgressText += $SpinChar
      DoDisplayLine($ProgressText)
      Start-Sleep -Milliseconds $MillisecondsDelay
      DoDisplayLine(([string]::new(' ', $ProgressText.Length)))
    }

    switch ($OperationID) { # Ending of current operation
      0 { # EndSearch 
        $SearchResult = $Searcher.EndSearch($SearchJob) # Get the search result
        if ($SearchResult.ResultCode -ne 2) {
          DoShowErrorAndExit("Search",$SearchResult.ResultCode,$SearchResult.HResult)
        }
        if ($SearchResult.Updates.Count -eq 0) { # Exit if no new updates available
          Write-Host "No updates found."
          exit
        }
        # Enumerate the updates available and prepare for next operations
        $ResultsTable = New-Object System.Data.DataTable
        $ResultsTable.Columns.Add("Index") | Out-Null
        $ResultsTable.Columns.Add("Size") | Out-Null
        $ResultsTable.Columns.Add("Title") | Out-Null
        $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
        $index = 1
        foreach ($WUpdate in $SearchResult.Updates) {
          $UpdateSize = FormattedSize $WUpdate.MaxDownloadSize
          $TotalSize = $TotalSize + $WUpdate.MaxDownloadSize
          $ResultsTable.Rows.Add($index, $UpdateSize, $WUpdate.Title) | Out-Null
          $UpdatesToDownload.Add($WUpdate) | Out-Null
          $index++
        }
        $ResultsTable | Format-Table -AutoSize -Wrap
        $SystemFreeSpace = (Get-Volume -DriveLetter "c").SizeRemaining
        if ($SystemFreeSpace -lt $TotalSize) {
          Write-Host "Total size to be downloaded: $TotalSize Bytes" 
          Write-Host "Free space available for download: $SystemFreeSpace Bytes" 
          Write-Host "There is not enough free space on the system drive to continue!"
          exit
        }
        if ($SearchOnly) { Exit }
      } 
      1 { # EndDownload
        $DownloadResult = $Downloader.EndDownload($DownloadJob) # Get the download result
        if ($DownloadResult.ResultCode -ne 2) {
          DoShowErrorAndExit("Download",$DownloadResult.ResultCode,$DownloadResult.HResult)
        }
        else {
          # Write-Host "[All updates were downloaded successfully]"
          $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
          foreach ($WUpdate in $Downloader.Updates) {
            If ($WUpdate.IsDownloaded -eq $true) {
              $UpdatesToInstall.Add($WUpdate) | Out-Null
            } 
          }
        }
                
        If ($DownloadOnly) { 
          Write-Host "Terminating execution after download, as requested!"
          exit } 
        # The download process is completed
      } 
      2 { # EndInstall
        $InstallResult = $Installer.EndInstall($InstallJob) # Get the result of the installation process
        if ($InstallResult.ResultCode -ne 2) {
          DoShowErrorAndExit("Installation",$InstallResult.ResultCode,$InstallResult.HResult)
        }
        else {
          Write-Host -ForegroundColor Yellow "[Installation completed successfully]" 
          # Show the updates that need to reboot to complete the installation
          If ($InstallResult.RebootRequired) {
            Write-Host "A reboot is needed to complete the installation of the following update(s):"
            for ($index=0; $index -lt $Installer.Updates.Count; $index++) {
              $UpdateResult = $InstallResult.GetUpdateResult($index)
              if ($UpdateResult.RebootRequired) {
                Write-Host $Installer.Updates[$index].Title
              }
            }
          }
        }
        # The installation process is completed
      } 
    } 
  }
}