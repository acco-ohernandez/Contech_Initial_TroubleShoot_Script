#
# Script.ps1
#
$TimeString = (Get-Date).ToString("mmddyyyy_HHmmss")
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = Split-Path -Parent -Path $script:MyInvocation.MyCommand.Path
$Log = "$env:TEMP\Log_$ScriptName`_$TimeString.txt"
Start-Transcript $Log

$InitalDriveInfoScriptBlock = {
    Write-Host "                   -*******************************************************************-" -ForegroundColor Green      
    Write-Host "                   -***                      Inital Drive Info:                     ***-" -ForegroundColor Green
    Write-Host "                   -*******************************************************************-" -ForegroundColor Green
    $drive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
    $StartingUsedSpace = $drive.Size - $drive.FreeSpace  # <-- This variable is used at the end to Get the spece clened up on the drvie.
} 

#region 1- Get-CDriveSpace
$Get_CDriveSpace_ScriptBlock = 
{
    function Get-CDriveSpace {
        #NWrite-Output "============= Hard Drive Info:"
        $drive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
        $usedSpace = $drive.Size - $drive.FreeSpace
        $availableSpace = $drive.FreeSpace
        $totalSpace = $drive.Size
    
        $usedSpaceGB = [math]::Round($usedSpace / 1GB, 2)
        $availableSpaceGB = [math]::Round($availableSpace / 1GB, 2)
        $totalSpaceGB = [math]::Round($totalSpace / 1GB, 2)
    
        $driveSpace = [PSCustomObject]@{
            C_Drive_UsedSpace = $usedSpaceGB
            AvailableSpace    = $availableSpaceGB
            TotalSpace        = $totalSpaceGB
        }
        Start-Sleep -Seconds 2
        return $driveSpace
    }
    Get-CDriveSpace | Out-String | Write-Host -ForegroundColor Yellow
}
    

#endregion ==================================================================================================

#region 2- Get-GPUDriverInfo
$GetGPUDriverInfo_ScriptBlock = 
{
#https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/dpi-related-apis-and-registry-settings?view=windows-11
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class NativeMethods
{
    [DllImport("user32.dll")]
    public static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("gdi32.dll")]
    public static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

    [DllImport("user32.dll")]
    public static extern int ReleaseDC(IntPtr hWnd, IntPtr hdc);
}

public class ResolutionUtils
{
    public static void GetNativeResolution(out int horizontalDPI, out int verticalDPI)
    {
        IntPtr desktopDc = NativeMethods.GetDC(IntPtr.Zero);
        horizontalDPI = NativeMethods.GetDeviceCaps(desktopDc, 88); // LOGPIXELSX
        verticalDPI = NativeMethods.GetDeviceCaps(desktopDc, 90); // LOGPIXELSY
        NativeMethods.ReleaseDC(IntPtr.Zero, desktopDc);
    }
}
"@

# Declare variables for DPI values
$horizontalDPI = 0
$verticalDPI = 0


    function CalculateScaling ($CurrentVerticalResolution)
    {
       [Math]::Round( ($CurrentVerticalResolution / [int][System.Windows.Forms.SystemInformation]::VirtualScreen.Height) * 100, 2 ).ToString("0.##")
    }

    function Get-GPUDriverInfo {
        Write-Output "                   -*************************************************************-"
        Write-Output "                   -***                  Video Driver Info:                   ***-"
        Write-Output "                   -*************************************************************-"
       
        $VideoResults +=
        Get-CimInstance -ClassName CIM_VideoController | 
            select DeviceID,
            Name,
            DriverDate,
            DriverVersion,
            CurrentVerticalResolution,
            CurrentHorizontalResolution,
            CurrentNumberOfColors,
            CurrentRefreshRate,
            MaxRefreshRate,
            @{Name = "Screen_Scale_Factor"; Expression = {(CalculateScaling $_.CurrentVerticalResolution)} },
            PNPDeviceID | % { $_ ; "---"} | Out-String
        
        Start-Sleep -Seconds 2
        # Call the GetNativeResolution function
        [ResolutionUtils]::GetNativeResolution([ref]$horizontalDPI, [ref]$verticalDPI)
        
        # Display the results
        $VideoResults +="Horizontal DPI: $horizontalDPI `n"
        $VideoResults +="Vertical DPI: $verticalDPI"


        return $VideoResults
    }
    Get-GPUDriverInfo | Out-String | Select-Object -ExcludeProperty RunspaceId 
}
#endregion ==================================================================================================

#region 3- Get-AdskLicensingVersion
$Get_AdskLicensingVersion_ScriptBlock = 
{
    function Get-AdskLicensingVersion {
        Write-Output "                   -********************************************************************-"
        Write-Output "                   -***                 Adsk Licensing Version Info:                 ***-"
        Write-Output "                   -********************************************************************-"
        $filePath = "C:\Program Files (x86)\Common Files\Autodesk Shared\AdskLicensing\version.ini"
    
        if (Test-Path $filePath) {
            $content = Get-Content $filePath -Raw
            $versionPattern = "version=(\d+\.\d+\.\d+\.\d+)"
            $version = $content | Select-String -Pattern $versionPattern | ForEach-Object { $_.Matches.Groups[1].Value }
    
            if ($version) {
                $result = [PSCustomObject]@{
                    AdskLicensingVersion = $version
                }
            }
            else {
                $result = [PSCustomObject]@{
                    AdskLicensingVersion = "No version information found in the file."
                }
            }
        }
        else {
            $result = [PSCustomObject]@{
                AdskLicensingVersion = "No AdskLicensing file found: $filePath"
            }
        }
        Start-Sleep -Seconds 2
        return $result
    }
    Get-AdskLicensingVersion | Out-String | Select-Object -ExcludeProperty RunspaceId 
}
#endregion ==================================================================================================

#region 4- Get-DisplayCount
$Get_DisplayCount_ScriptBlock = 
{
    function Get-DisplayCount {
        Write-Output "                   -*****************************************************************-"
        Write-Output "                   -***                    Displays List Info:                    ***-"
        Write-Output "                   -*****************************************************************-"
        $Displays = Get-CimInstance -Namespace "root\wmi" -ClassName WmiMonitorConnectionParams | select -ExpandProperty InstanceName
        $count = 0
        $Displays | % { $count++; Write-Output "$($count)- $_.InstanceName" } 
        Start-Sleep -Seconds 2
    }
    Get-DisplayCount | Out-String | Select-Object -ExcludeProperty RunspaceId 
}
#endregion ==================================================================================================

#region 5-  Formated-RevitFolderSizes And Get-RevitFolderSizes
$Formated_RevitFolderSizes_scriptBlock = {
    function Get-RevitFolderSizes {
        $folders = @(
            "$env:LOCALAPPDATA\Autodesk\Revit",
            "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2018\CollaborationCache",
            "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2019\CollaborationCache",
            "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2020\CollaborationCache",
            "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2021\CollaborationCache",
            "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2022\CollaborationCache",
            "$env:LOCALAPPDATA\Autodesk\Revit\Autodesk Revit 2023\CollaborationCache",
            "$env:USERPROFILE\accdocs",
            "$env:USERPROFILE\Downloads",
            "$env:USERPROFILE\Box\Box",
            "C:\Windows\IMECache",
            "C:\Autodesk"
            "C:\Autodesk\WI"
    
    
    
        )
    
        $folderSizes = foreach ($folder in $folders) {
            if (Test-Path -Path $folder) {
                $sizeInBytes = Get-ChildItem -Path $folder -Recurse | Measure-Object -Property Length -Sum | Select-Object -ExpandProperty Sum
                $sizeInMB = [math]::Round($sizeInBytes / 1MB, 2)
                $sizeInGB = [math]::Round($sizeInBytes / 1GB, 2)
                [PSCustomObject]@{
                    Folder = $folder
                    SizeMB = $sizeInMB
                    SizeGB = $sizeInGB
                }
            }
            else {
                [PSCustomObject]@{
                    SizeMB = "N/A"
                    SizeGB = "N/A"
                    Folder = $folder
                }
            }
        }
    
        return $folderSizes
    }
    # Example usage:
    # Get-RevitFolderSizes
    
    
    
    function Formated-RevitFolderSizes {
        Write-Output "                   -**************************************************************-"
        Write-Output "                   -***                  Folders Sizes Info:                   ***-"
        Write-Output "                   -**************************************************************-"
        $sizes = Get-RevitFolderSizes
        $sizes | Out-String | Write-Output
        
        $totalMB = 0
        $totalGB = 0
        
        $sizes | Where-Object { $_.SizeMB -ne "N/A" } | ForEach-Object {
            $totalMB += $_.SizeMB
            $totalGB += $_.SizeGB
        }
        
        $totalMB = [math]::Round($totalMB, 2)
        $totalGB = [math]::Round($totalGB, 2)
        
        Write-Output "Total Space used by folders"
        Write-Output "In MB: $totalMB"
        Write-Output "In GB: $totalGB"
        Start-Sleep -Seconds 2
    }
    # Example usage:
    Formated-RevitFolderSizes | Out-String | Select-Object -ExcludeProperty RunspaceId 
}
#endregion ==================================================================================================

#region 6- Check-AutodeskOdisRegistry
$Check_AutodeskOdisRegistry_scriptBlock = 
{
    function Check-AutodeskOdisRegistry {
        Write-Output "                   -********************************************************************-"
        Write-Output "                   -***                  Autodesk ODIS Registry Info:                ***-"
        Write-Output "                   -********************************************************************-"
        $registryPath = "HKCU:\SOFTWARE\Autodesk\ODIS"
        $registryValueName = "DisableManualUpdateInstall"
        $registryValueData = 1
        
        if (Test-Path -Path $registryPath) {
            $existingValue = Get-ItemProperty -Path $registryPath -Name $registryValueName -ErrorAction SilentlyContinue
            if ($existingValue -and $existingValue.$registryValueName -eq $registryValueData) {
                Write-Output "Existing value: HKCU\Software\Autodesk\ODIS /V DisableManualUpdateInstall /D 1."
            }
            else {
                Set-ItemProperty -Path $registryPath -Name $registryValueName -Value $registryValueData
                Write-Output "Value updated: HKCU\Software\Autodesk\ODIS /V DisableManualUpdateInstall /D 1."
            }
        }
        else {
            New-Item -Path $registryPath -Force | Out-Null
            New-ItemProperty -Path $registryPath -Name $registryValueName -Value $registryValueData -PropertyType DWORD -Force | Out-Null
            Write-Output "Registry key and value added. `nHKCU\Software\Autodesk\ODIS /V DisableManualUpdateInstall /D 1"
        }
        Start-Sleep -Seconds 2
    }
    # Example usage:
    Check-AutodeskOdisRegistry | Out-String | Select-Object -ExcludeProperty RunspaceId 
}
#endregion ==================================================================================================

#region 7- PowerCFG_Off
$PowerCFG_Off_ScriptBlock = 
{
    function PowerCFG-Off {
        Write-Output "                   -*******************************************************************-"
        Write-Output "                   -***                  Hibernation status Info:                   ***-"
        Write-Output "                   -*******************************************************************-"
        # Check if hibernation is currently enabled in the registry
        $hibernationEnabled = Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Power" -Name "HibernateEnabled"

        if ($hibernationEnabled -eq 1) {
            Write-Output "Hibernation is currently turned on."

            # Disable hibernation using powercfg
            Write-Output "Disabling hibernation..."
            powercfg /h off | Out-Null
            Start-Sleep -Milliseconds 200

            # Check hibernation status again
            $hibernationEnabled = Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Power" -Name "HibernateEnabled"
    
            if ($hibernationEnabled -eq 0) {
                Write-Output "Hibernation is now turned off."
            }
            else {
                Write-Output "Failed to turn off hibernation."
            }
        }
        else {
            Write-Output "Hibernation is currently turned off."
        }

    }
    PowerCFG-Off | Out-String | Select-Object -ExcludeProperty RunspaceId 
}
    

#endregion ==================================================================================================

#region 8- Delete C:\Autodesk\WI
$DeleteAutodeskWI_ScriptBlock = 
{
    function DeleteAutodeskWI {
        $outputString = "                    -*******************************************************************-`n"
        $outputString += "                   -***                    Delete 'C:\Autodesk\WI'                  ***-`n"
        $outputString += "                   -*******************************************************************-`n"
        $FolderToDelete = "C:\Autodesk\WI - Copy"
        $DelThisFolder = Get-Item -Path $FolderToDelete -ErrorAction SilentlyContinue                 # Get the folder item
        if ($DelThisFolder) {
            $outputString += "Found: $FolderToDelete `n"
            $outputString += "Deleting: $FolderToDelete `n"
            $DelThisFolder.Delete($true)                                        # Delete folder item with .NET Method
        }
        else {
            $outputString += "$FolderToDelete - NOT FOUND! `n"
        }
        return $outputString
    }
    DeleteAutodeskWI #| Out-String | Select-Object -ExcludeProperty RunspaceId 
}
    
#endregion ==================================================================================================

#region 9- FinalDriveSpaceCheck
$FinalDriveSpaceCheck_ScriptBlock = 
{
    function FinalDriveSpaceCheck {
        $outputString = "$scriptCount $('='*80) `n"
                
        $outputString += "                   -*************************************************************-`n"
        $outputString += "                   -***                  Final Drive Info:                    ***-`n"
        $outputString += "                   -*************************************************************-`n"
        $drive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
        $EndingUsedSpace = $drive.Size - $drive.FreeSpace
        $crearedSpace = $($StartingUsedSpace - $EndingUsedSpace)

        $CleanedSpace = [PSCustomObject]@{
            ClearedSpace_MB = [math]::Round($crearedSpace / 1MB, 2)
            ClearedSpace_GB = [math]::Round($crearedSpace / 1GB, 2)
        }

        $outputString += $CleanedSpace | Out-String
        $outputString

    }
    FinalDriveSpaceCheck | Out-String | Write-Host -ForegroundColor Yellow
}
    
#endregion ==================================================================================================
#############

$ProcessAllScriptBlocksInBackgroundJobs = {
    #==================== Process all Background Jobs process ====================
    $jobs = & {
        # Start jobs for each script block
        #$job1 = Start-Job -ScriptBlock $Get_CDriveSpace_ScriptBlock            
        $job2 = Start-Job -ScriptBlock $Get_DisplayCount_ScriptBlock
        $job3 = Start-Job -ScriptBlock $GetGPUDriverInfo_ScriptBlock
        $job4 = Start-Job -ScriptBlock $Get_AdskLicensingVersion_ScriptBlock
        $job5 = Start-Job -ScriptBlock $Formated_RevitFolderSizes_scriptBlock
        $job6 = Start-Job -ScriptBlock $Check_AutodeskOdisRegistry_scriptBlock
        $job7 = Start-Job -ScriptBlock $PowerCFG_Off_ScriptBlock
        $job8 = Start-Job -ScriptBlock $DeleteAutodeskWI_ScriptBlock

        # Return all the jobs
        $job1, $job2, $job3, $job4, $job5, $job6, $job7, $job8
    }

    $colors = @('Green', 'Cyan')
    $currentColorIndex = 0
    $scriptCount = 0
    Write-Host "Please wait while all scriptblocks are processed..." -ForegroundColor Yellow

    while ($jobs.Count -gt 0) {
        foreach ($job in $jobs) {
            if ($job.State -eq 'Completed') {
                # Retrieve the result from the completed job
                $result = Receive-Job -Job $job
                $color = $colors[$currentColorIndex]
                $scriptCount++

                "$scriptCount |$('-'*100)|" | Out-String | Write-Host -ForegroundColor $color
                $result | Out-String | Write-Host -ForegroundColor $color

                Write-Host "Next..." -ForegroundColor $color

                # Remove the completed job from the job list
                $jobs = $jobs | Where-Object { $_.Id -ne $job.Id }
                $job | Remove-Job | Out-Null

                # Switch to the next color
                $currentColorIndex = ($currentColorIndex + 1) % $colors.Count
            }
            elseif ($job.State -eq 'Running') {
                # Get the current progress of the job
                $progress = $job.Progress

                # Update the progress if available
                if ($progress) {
                    Write-Host "Job $($job.Id): $progress%"
                }

                # Wait for the job to complete
                $job | Wait-Job -Timeout 1 | Out-Null
            }
        }

        # Pause for a moment before checking again
        Start-Sleep -Milliseconds 500
    }
}
$PromptTheUserToStartTheLogInMicrosoftEdge_ScriptBlock = {


    # Prompt the user to start the log in Microsoft Edge
    do {
        Write-Host "Do you want to start the log file in Microsoft Edge? (y/n)" -ForegroundColor Yellow
        $choice = Read-Host "Enter your answer here -> "
        if ($choice -eq 'y') {
            # Start Microsoft Edge with the log file
            Start-Process msedge $log
            Write-Host "Script complete." -ForegroundColor Green
            break
        }
        elseif ($choice -eq 'n') {
            # User chose not to start the log file
            Write-Host "Script complete." -ForegroundColor Green
            break
        }
        else {
            # Invalid choice
            Write-Host "Invalid choice. `nPlease enter 'y' or 'n'." -ForegroundColor Yellow
        }
    } while ($true)
}

cls 
#region ==================== Process all the scriptblocks ====================
  
& $InitalDriveInfoScriptBlock

& $Get_CDriveSpace_ScriptBlock

& $ProcessAllScriptBlocksInBackgroundJobs # <-- Background Jobs process

& $FinalDriveSpaceCheck_ScriptBlock

& $Get_CDriveSpace_ScriptBlock

Stop-Transcript | Out-Null
"=== DONE! ==="
& $PromptTheUserToStartTheLogInMicrosoftEdge_ScriptBlock

#endregion