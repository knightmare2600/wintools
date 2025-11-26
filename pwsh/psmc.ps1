#!/usr/bin/env pwsh

<#
=============================== VERSION HISTORY ===============================

v1.0.0 - Initial Release
  - Basic two-pane file manager with Terminal.Gui
  - Directory navigation and file listing
  - File operations: copy, move, delete, rename
  - Basic search functionality

v1.1.0 - Enhanced File Operations
  - Added multi-file selection support
  - Implemented pattern-based selection (e.g., *.txt)
  - Added file size display with proper formatting
  - Improved keyboard shortcuts

v1.2.0 - Progress Dialog Implementation
  - Added progress dialog for multi-file operations (3+ files)
  - Text-based progress bar with visual feedback
  - Real-time status updates during file operations
  - Overwrite confirmation dialogs (Yes/No/Yes to All/No to All)

v1.2.1 - Progress Dialog Bug Fixes
  - Fixed progress dialog not displaying (added Application.Run)
  - Fixed progress bar not updating (added RunIteration for UI refresh)
  - Fixed timer events not firing (switched to MainLoop.AddTimeout)
  - Added comprehensive debug logging for troubleshooting
  - Progress now updates correctly during file copy operations

v1.5.9 - Progrss Refinements
  - Refined progress update mechanism
  - Improved UI responsiveness during operations
  - Enhanced error handling in progress dialogs

v1.6.0 - Cleanup Code
  - Merge Pane switching and focus code
  - Rework phrasing in menus
  - Declutter status bar

v1.6.1 - 
  - Add Rename function on F3 to menus
  - Fix some typos
  - Improve Menu and Statusbar flow

v1.6.2 - 
  - Add Error action code to track down silent failures

================================================================================
#>

#Requires -Version 7.0

<#
.SYNOPSIS
    PowerShell Commander (PSMC) v1.6.0 STABLE
.NOTES
    Version: $($Global:PSMC_Version)
    Terminal.Gui: v1.16.0
#>

$ErrorActionPreference = 'Continue'
$Global:PSMC_Version = '1.6.2'

param([switch]$Verbose)

function Debug-Log {
    param([string]$Message)
    if ($Verbose) {
        $ts = (Get-Date).ToString('HH:mm:ss')
        Write-Host "[$ts] LOG: $Message" -ForegroundColor Cyan
    }
}

# Load Terminal.Gui
if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Terminal.Gui' })) {
    $mod = Get-Module Microsoft.PowerShell.ConsoleGuiTools -List
    if (-not $mod) { throw "ConsoleGuiTools module not found" }
    $dll = Join-Path $mod.ModuleBase 'Terminal.Gui.dll'
    Add-Type -Path $dll
    
    if ($Verbose) {
        $asmVer = [System.Reflection.AssemblyName]::GetAssemblyName($dll).Version
        Debug-Log "Terminal.Gui version: $asmVer"
        Debug-Log "Module version: $($mod.Version)"
    }
}

[Terminal.Gui.Application]::Init()

function Get-Theme {
    param([string]$mode)
    $cs = [Terminal.Gui.ColorScheme]::new()
    if ($mode -eq 'light') {
        $cs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Gray)
        $cs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Cyan)
        $cs.HotNormal = $cs.Normal
        $cs.HotFocus  = $cs.Focus
    } else {
        $cs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Black)
        $cs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::DarkGray)
        $cs.HotNormal = $cs.Normal
        $cs.HotFocus  = $cs.Focus
    }
    return $cs
}

function Build-DisplayNames {
    param([System.IO.FileSystemInfo[]]$items, [string]$path)
    $names = [System.Collections.Generic.List[string]]::new()
    
    try {
        $parent = [System.IO.Directory]::GetParent($path)
        if ($parent -ne $null) {
            $names.Add("..")
            Debug-Log "Added '..' for: $path"
        }
    } catch {
        Debug-Log "No parent for: $path"
    }
    
    foreach ($d in ($items | Where-Object PSIsContainer | Sort-Object Name)) { 
        $names.Add("[DIR] $($d.Name)") 
    }
    
    foreach ($f in ($items | Where-Object {-not $_.PSIsContainer} | Sort-Object Name)) { 
        $names.Add($f.Name) 
    }
    
    Debug-Log "Built $($names.Count) items for $path"
    return $names
}

function Handle-ListView-Enter {
    param(
        $pane,
        [Terminal.Gui.KeyEventEventArgs] $args
    )

    # Detect Enter key (Return or Enter, both appear as \r or 13)
    if ($args.KeyEvent.Key -ne [Terminal.Gui.Key]::Enter) {
        return
    }

    # Prevent Terminal.Gui from bubbling this event
    $args.Handled = $true

    # Get selected index
    $idx = $pane.ListView.SelectedItem
    if ($idx -lt 0) { return }

    # Get the underlying list
    $sourceList = $pane.ListView.Source.ToList()
    $item = $sourceList[$idx]

    # ---------------------------- " .. "  → go UP directory ----------------------------
    if ($item -eq "..") {
        try {
            $parent = [System.IO.Directory]::GetParent($pane.Path)
            if ($parent -ne $null) {
                $pane.Path = $parent.FullName
                Refresh-Pane $pane
            }
        } catch {
            Show-Modal "Error" "Cannot move to parent directory."
        }
        return
    }

    # ---------------------------- "[DIR] foldername" → go INTO directory ----------------------------
    if ($item.StartsWith("[DIR]")) {
        # Strip prefix
        $folder = $item.Substring(5).Trim()

        $newPath = Join-Path $pane.Path $folder
        if (-not (Test-Path ${newPath})) {
            Show-Modal "Error" "Folder not found: ${folder}"
            return
        }

        try {
            $pane.Path = ${newPath}
            Refresh-Pane $pane
        } catch {
            Show-Modal "Error" "Unable to enter directory: $($_.Exception.Message)"
        }
        return
    }

    # ---------------------------- No action: clicked file ----------------------------
    # Later you can add: Open viewer, file preview, etc.
}


function Navigate-ToItem {
    param($pane, [int]$selectedIndex)
    
    if ($pane -eq $null -or $pane.ListView -eq $null -or $pane.ListView.Source -eq $null) {
        Debug-Log "Navigate: Invalid pane state"
        Show-Modal "Error" "Navigation failed"
        return
    }
    
    $sourceList = $pane.ListView.Source.ToList()
    $count = $sourceList.Count
    
    Debug-Log "Navigate: pane=$($pane.Name), idx=$selectedIndex, count=$count"
    
    if ($selectedIndex -lt 0 -or $selectedIndex -ge $count) { 
        return 
    }
    
    $selected = $sourceList[$selectedIndex]
    Debug-Log "Navigate: item='$selected'"
    
    if ($selected -eq "..") {
        try {
            $parent = [System.IO.Directory]::GetParent($pane.Path)
            if ($parent -ne $null) {
                $pane.Path = $parent.FullName
                Refresh-Pane $pane
                Show-Modal "Navigation" "Moved UP to:`n$($parent.FullName)"
            } else {
                Show-Modal "Info" "Already at root"
            }
        } catch {
            Show-Modal "Error" "Cannot navigate to parent"
        }
        return
    }
    
    if ($selected.StartsWith("[DIR] ")) {
        $dirName = $selected.Substring(6)
        $newPath = Join-Path $pane.Path $dirName
        
        if (Test-Path -LiteralPath $newPath -PathType Container) {
            $pane.Path = (Resolve-Path -LiteralPath $newPath).Path
            Refresh-Pane $pane
            Show-Modal "Navigation" "Entered: $dirName"
        } else {
            Show-Modal "Error" "Directory not found"
        }
        return
    }
    
    $filePath = Join-Path $pane.Path $selected
    if (Test-Path -LiteralPath $filePath -PathType Leaf) {
        $info = Get-Item -LiteralPath $filePath
        $sizeKB = [math]::Round($info.Length / 1KB, 2)
        $msg = "File: $($info.Name)`nSize: $($info.Length) bytes ($sizeKB KB)`nModified: $($info.LastWriteTime)"
        Show-Modal "File Info" $msg
    }
}

function Show-Modal { 
    param($title, $msg) 
    [Terminal.Gui.MessageBox]::Query($title, $msg, @("OK")) | Out-Null 
}

function New-FilePane {
    param([string]$initialPath, [string]$themeMode='dark', [string]$paneName='')
    
    $pane = [PSCustomObject]@{ 
        Path = ''
        Items = @()
        Frame = $null
        ListView = $null
        Name = $paneName
    }
    
    $pane.Path = (Resolve-Path $initialPath).Path
    $pane.Items = Get-ChildItem -LiteralPath $pane.Path -Force -ErrorAction SilentlyContinue
    $displayNames = Build-DisplayNames $pane.Items $pane.Path

    $frame = [Terminal.Gui.FrameView]::new()
    $frame.Title = "$paneName : $($pane.Path)"
    $frame.Y = 0
    $frame.Height = [Terminal.Gui.Dim]::Fill()
    $frame.ColorScheme = Get-Theme $themeMode

    $list = [Terminal.Gui.ListView]::new()
    $list.X = 0
    $list.Y = 0
    $list.Width = [Terminal.Gui.Dim]::Fill()
    $list.Height = [Terminal.Gui.Dim]::Fill()
    $list.CanFocus = $true
    $list.AllowsMarking = $false
    
    $sourceList = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $displayNames) {
        $sourceList.Add($item)
    }
    $list.SetSource($sourceList)
    
    Debug-Log "[$paneName] Created with $($sourceList.Count) items"

    $frame.Add($list)
    $pane.Frame = $frame
    $pane.ListView = $list
    return $pane
}

function Refresh-Pane {
    param($pane)
    
    Debug-Log "[$($pane.Name)] Refresh: $($pane.Path)"
    
    $pane.Items = Get-ChildItem -LiteralPath $pane.Path -Force -ErrorAction SilentlyContinue
    $displayNames = Build-DisplayNames $pane.Items $pane.Path
    $pane.Frame.Title = "$($pane.Name) : $($pane.Path)"
    
    $sourceList = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $displayNames) {
        $sourceList.Add($item)
    }
    $pane.ListView.SetSource($sourceList)
    
    if ($sourceList.Count -gt 1) {
        $pane.ListView.SelectedItem = 1
    } else {
        $pane.ListView.SelectedItem = 0
    }
    
    Debug-Log "[$($pane.Name)] Refreshed: $($sourceList.Count) items"
}

function Get-CurrentFile {
    param($pane)
    
    $idx = $pane.ListView.SelectedItem
    $sourceList = $pane.ListView.Source.ToList()
    
    if ($idx -lt 0 -or $idx -ge $sourceList.Count) {
        return $null
    }
    
    $item = $sourceList[$idx]
    
    if ($item -eq ".." -or $item.StartsWith("[DIR] ")) { 
        return $null
    }
    
    $filePath = Join-Path $pane.Path $item
    if (Test-Path -LiteralPath $filePath -PathType Leaf) {
        return $filePath
    }
    
    return $null
}

function Select-FilesByPattern {
    param($pane)
    
    Debug-Log "=== Select by Pattern ==="
    
    $dlg = [Terminal.Gui.Dialog]::new("Select Files by Pattern - $($pane.Name)", 70, 14)
    $lbl1 = [Terminal.Gui.Label]::new(1, 1, "Enter pattern to match files:")
    $lbl2 = [Terminal.Gui.Label]::new(1, 2, "Examples: *.ps1  test*.*  *report*  file?.txt")
    $txt = [Terminal.Gui.TextField]::new(1, 4, 66, "*.ps1")
    $lbl3 = [Terminal.Gui.Label]::new(1, 6, "Matched files will be highlighted in title")
    $btnOK = [Terminal.Gui.Button]::new(18, 9, "Select")
    $btnCancel = [Terminal.Gui.Button]::new(36, 9, "Cancel")
    
    $btnOK.add_Clicked({
        $pattern = $txt.Text.ToString().Trim()
        
        if ([string]::IsNullOrWhiteSpace($pattern)) {
            [Terminal.Gui.MessageBox]::ErrorQuery("Invalid Pattern", "Pattern cannot be empty", @("OK")) | Out-Null
            return
        }
        
        Debug-Log "Pattern: $pattern"
        
        # Initialize selection array if not exists
        if (-not $pane.PSObject.Properties['SelectedFiles']) {
            $pane | Add-Member -NotePropertyName SelectedFiles -NotePropertyValue ([System.Collections.Generic.List[string]]::new())
        }
        
        # Get all items from current pane
        $sourceList = $pane.ListView.Source.ToList()
        $matchedCount = 0
        
        for ($i = 0; $i -lt $sourceList.Count; $i++) {
            $item = $sourceList[$i]
            
            # Skip parent directory and directories
            if ($item -eq ".." -or $item.StartsWith("[DIR] ")) {
                continue
            }
            
            # Match against pattern using -like operator
            if ($item -like $pattern) {
                # Only add if not already in selection
                if (-not $pane.SelectedFiles.Contains($item)) {
                    $pane.SelectedFiles.Add($item)
                    $matchedCount++
                    Debug-Log "Added to selection: $item"
                } else {
                    Debug-Log "Already selected (skipped): $item"
                }
            }
        }
        
        [Terminal.Gui.Application]::RequestStop()
        
        if ($matchedCount -eq 0) {
            Show-Modal "No New Matches" "No new files matched pattern:`n$pattern`n`nTotal selected: $($pane.SelectedFiles.Count)"
        } else {
            # Update pane title to show selection count
            $pane.Frame.Title = "$($pane.Name) : $($pane.Path) [$($pane.SelectedFiles.Count) selected]"
            
            $msg = "Added $matchedCount file(s) matching:`n$pattern`n`nTotal selected: $($pane.SelectedFiles.Count)"
            Show-Modal "Files Selected" $msg
            Debug-Log "Total selected files in pane: $($pane.SelectedFiles.Count)"
        }
    })
    
    $btnCancel.add_Clicked({ 
        [Terminal.Gui.Application]::RequestStop() 
    })
    
    $dlg.Add($lbl1)
    $dlg.Add($lbl2)
    $dlg.Add($txt)
    $dlg.Add($lbl3)
    $dlg.AddButton($btnOK)
    $dlg.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($dlg)
}

# Also add an Unselect All function:

function Unselect-AllFiles {
    param($pane)
    
    Debug-Log "=== Unselect All ==="
    
    if ($pane.PSObject.Properties['SelectedFiles']) {
        $count = $pane.SelectedFiles.Count
        $pane.SelectedFiles.Clear()
        $pane.Frame.Title = "$($pane.Name) : $($pane.Path)"
        Show-Modal "Unselect All" "Cleared $count selected file(s)"
        Debug-Log "Cleared all selections"
    } else {
        Show-Modal "Unselect All" "No files were selected"
    }
}

# And a Select All function:

function Select-AllFiles {
    param($pane)
    
    Debug-Log "=== Select All ==="
    
    # Initialize selection array if not exists
    if (-not $pane.PSObject.Properties['SelectedFiles']) {
        $pane | Add-Member -NotePropertyName SelectedFiles -NotePropertyValue ([System.Collections.Generic.List[string]]::new())
    } else {
        $pane.SelectedFiles.Clear()
    }
    
    # Get all items and add files only
    $sourceList = $pane.ListView.Source.ToList()
    $count = 0
    
    for ($i = 0; $i -lt $sourceList.Count; $i++) {
        $item = $sourceList[$i]
        
        # Skip parent directory and directories
        if ($item -eq ".." -or $item.StartsWith("[DIR] ")) {
            continue
        }
        
        $pane.SelectedFiles.Add($item)
        $count++
    }
    
    $pane.Frame.Title = "$($pane.Name) : $($pane.Path) [ALL $count selected]"
    Show-Modal "Select All" "Selected ALL $count file(s) in $($pane.Name) pane"
    Debug-Log "Selected all $count files"
}

function Copy-CurrentFile {
    param($sourcePane, $destPane)
    
    Debug-Log "=== COPY ==="
    
    # Check if there are selected files
    $filesToCopy = @()
    
    if ($sourcePane.PSObject.Properties['SelectedFiles'] -and $sourcePane.SelectedFiles.Count -gt 0) {
        # Multi-file copy
        Debug-Log "Multi-file copy: $($sourcePane.SelectedFiles.Count) files selected"
        
        foreach ($fileName in $sourcePane.SelectedFiles) {
            $filePath = Join-Path $sourcePane.Path $fileName
            if (Test-Path -LiteralPath $filePath -PathType Leaf) {
                $filesToCopy += $filePath
            }
        }
        
        if ($filesToCopy.Count -eq 0) {
            Show-Modal "Copy" "No valid files in selection"
            return
        }
    } else {
        # Single file copy
        $file = Get-CurrentFile $sourcePane
        
        if ($file -eq $null) {
            Show-Modal "Copy" "No file selected.`n`nPlace cursor on a file and press F6."
            return
        }
        
        $filesToCopy = @($file)
    }
    
    Debug-Log "Files to copy: $($filesToCopy.Count)"
    
    # If copying to same directory and single file, prompt for new name
    if ($sourcePane.Path -eq $destPane.Path -and $filesToCopy.Count -eq 1) {
        Debug-Log "Same directory, single file - prompting for new name"
        
        $fileName = Split-Path -Leaf $filesToCopy[0]
        
        $dlg = [Terminal.Gui.Dialog]::new("Copy to Same Directory", 70, 12)
        $lbl1 = [Terminal.Gui.Label]::new(1, 1, "Copying in same directory requires a new name:")
        $lbl2 = [Terminal.Gui.Label]::new(1, 2, "Original: $fileName")
        $lbl3 = [Terminal.Gui.Label]::new(1, 4, "New name:")
        $txt = [Terminal.Gui.TextField]::new(1, 5, 66, $fileName)
        $btnOK = [Terminal.Gui.Button]::new(18, 8, "Copy")
        $btnCancel = [Terminal.Gui.Button]::new(36, 8, "Cancel")
        
        $script:DialogResult = $null
        
        $btnOK.add_Clicked({
            $newName = $txt.Text.ToString().Trim()
            
            if ([string]::IsNullOrWhiteSpace($newName)) {
                [Terminal.Gui.MessageBox]::ErrorQuery("Invalid Name", "File name cannot be empty", @("OK")) | Out-Null
                return
            }
            
            if ($newName -eq $fileName) {
                [Terminal.Gui.MessageBox]::ErrorQuery("Same Name", "New name must be different from original", @("OK")) | Out-Null
                return
            }
            
            $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
            if ($newName.IndexOfAny($invalidChars) -ge 0) {
                [Terminal.Gui.MessageBox]::ErrorQuery("Invalid Name", "File name contains invalid characters", @("OK")) | Out-Null
                return
            }
            
            $script:DialogResult = $newName
            [Terminal.Gui.Application]::RequestStop()
        })
        
        $btnCancel.add_Clicked({ 
            $script:DialogResult = $null
            [Terminal.Gui.Application]::RequestStop() 
        })
        
        $dlg.Add($lbl1)
        $dlg.Add($lbl2)
        $dlg.Add($lbl3)
        $dlg.Add($txt)
        $dlg.AddButton($btnOK)
        $dlg.AddButton($btnCancel)
        
        [Terminal.Gui.Application]::Run($dlg)
        
        if ($script:DialogResult -eq $null) {
            Debug-Log "User cancelled rename"
            return
        }
        
        $destFileName = $script:DialogResult
        $dest = Join-Path $destPane.Path $destFileName
        
        if (Test-Path -LiteralPath $dest) {
            $result = [Terminal.Gui.MessageBox]::Query("File Exists", "Destination file already exists:`n$dest`n`nOverwrite?", @("Yes", "No"))
            if ($result -ne 0) {
                return
            }
        }
        
        $msg = "Copy file:`n  FROM: $($filesToCopy[0])`n  TO: $dest"
        $result = [Terminal.Gui.MessageBox]::Query("Confirm Copy", $msg, @("Yes", "No"))
        
        if ($result -eq 0) {
            try {
                Copy-Item -LiteralPath $filesToCopy[0] -Destination $dest -Force
                Refresh-Pane $destPane
                Show-Modal "Success" "File copied to:`n$destFileName"
            } catch {
                Show-Modal "Error" "Copy failed:`n$($_.Exception.Message)"
            }
        }
        
        return
    }
    
    # Multi-file copy or different directory
    if ($sourcePane.Path -eq $destPane.Path -and $filesToCopy.Count -gt 1) {
        Show-Modal "Copy Error" "Cannot copy multiple files to same directory!`n`nUse different destination pane."
        return
    }
    
    # Build confirmation message
    if ($filesToCopy.Count -eq 1) {
        $fileName = Split-Path -Leaf $filesToCopy[0]
        $msg = "Copy 1 file:`n  $fileName`n`nTO: $($destPane.Path)"
    } else {
        $fileList = ($filesToCopy | ForEach-Object { "  • $(Split-Path -Leaf $_)" } | Select-Object -First 5) -join "`n"
        if ($filesToCopy.Count -gt 5) {
            $fileList += "`n  ... and $($filesToCopy.Count - 5) more"
        }
        $msg = "Copy $($filesToCopy.Count) files:`n$fileList`n`nTO: $($destPane.Path)"
    }
    
    $result = [Terminal.Gui.MessageBox]::Query("Confirm Copy", $msg, @("Yes", "No"))
    
    if ($result -ne 0) {
        return
    }
    
    # Show progress bar for multiple files (3 or more)
    if ($filesToCopy.Count -ge 3) {
        $copiedCount = 0
        $errorCount = 0
        $overwriteAll = $false
        $skipAll = $false
        
        Show-ProgressDialog -Title "Copying Files" -Total $filesToCopy.Count -Operation {
            for ($i = 0; $i -lt $filesToCopy.Count; $i++) {
                $file = $filesToCopy[$i]
                $fileName = Split-Path -Leaf $file
                $dest = Join-Path $destPane.Path $fileName
                
                Update-Progress -Current ($i + 1) -Status "Copying: $fileName"
                
                # Check if destination exists
                if (Test-Path -LiteralPath $dest) {
                    if (-not $overwriteAll -and -not $skipAll) {
                        $result = [Terminal.Gui.MessageBox]::Query("File Exists", "File exists:`n$fileName`n`nOverwrite?", @("Yes", "Yes to All", "No", "No to All"))
                        
                        if ($result -eq 1) { $overwriteAll = $true }
                        elseif ($result -eq 3) { $skipAll = $true; continue }
                        elseif ($result -eq 2) { continue }
                    } elseif ($skipAll) {
                        continue
                    }
                }
                
                try {
                    Copy-Item -LiteralPath $file -Destination $dest -Force
                    $copiedCount++
                    Debug-Log "Copied: $fileName"
                } catch {
                    $errorCount++
                    Debug-Log "Error copying $fileName : $($_.Exception.Message)"
                }
                
                Start-Sleep -Milliseconds 50  # Small delay to see progress
            }
        }
        
        Refresh-Pane $destPane
        
        if ($sourcePane.PSObject.Properties['SelectedFiles']) {
            $sourcePane.SelectedFiles.Clear()
            $sourcePane.Frame.Title = "$($sourcePane.Name) : $($sourcePane.Path)"
        }
        
        if ($errorCount -eq 0) {
            Show-Modal "Success" "Copied $copiedCount file(s) successfully"
        } else {
            Show-Modal "Partial Success" "Copied: $copiedCount`nFailed: $errorCount"
        }
        
        return
    }
    
    # Single or few files - no progress bar needed
    $copiedCount = 0
    $errorCount = 0
    $overwriteAll = $false
    $skipAll = $false
    
    foreach ($file in $filesToCopy) {
        $fileName = Split-Path -Leaf $file
        $dest = Join-Path $destPane.Path $fileName
        
        if (Test-Path -LiteralPath $dest) {
            if (-not $overwriteAll -and -not $skipAll) {
                $result = [Terminal.Gui.MessageBox]::Query("File Exists", "File exists:`n$fileName`n`nOverwrite?", @("Yes", "Yes to All", "No", "No to All"))
                
                if ($result -eq 1) { $overwriteAll = $true }
                elseif ($result -eq 3) { $skipAll = $true; continue }
                elseif ($result -eq 2) { continue }
            } elseif ($skipAll) {
                continue
            }
        }
        
        try {
            Copy-Item -LiteralPath $file -Destination $dest -Force
            $copiedCount++
            Debug-Log "Copied: $fileName"
        } catch {
            $errorCount++
            Debug-Log "Error copying $fileName : $($_.Exception.Message)"
        }
    }
    
    Refresh-Pane $destPane
    
    if ($sourcePane.PSObject.Properties['SelectedFiles']) {
        $sourcePane.SelectedFiles.Clear()
        $sourcePane.Frame.Title = "$($sourcePane.Name) : $($sourcePane.Path)"
    }
    
    if ($errorCount -eq 0) {
        Show-Modal "Success" "Copied $copiedCount file(s) successfully"
    } else {
        Show-Modal "Partial Success" "Copied: $copiedCount`nFailed: $errorCount"
    }
}

function Move-CurrentFile {
    param($sourcePane, $destPane)
    
    Debug-Log "=== Move ==="
    
    # Check if there are selected files
    $filesToMove = @()
    
    if ($sourcePane.PSObject.Properties['SelectedFiles'] -and $sourcePane.SelectedFiles.Count -gt 0) {
        # Multi-file Move
        Debug-Log "Multi-file Move: $($sourcePane.SelectedFiles.Count) files selected"
        
        foreach ($fileName in $sourcePane.SelectedFiles) {
            $filePath = Join-Path $sourcePane.Path $fileName
            if (Test-Path -LiteralPath $filePath -PathType Leaf) {
                $filesToMove += $filePath
            }
        }
        
        if ($filesToMove.Count -eq 0) {
            Show-Modal "Move" "No valid files in selection"
            return
        }
    } else {
        # Single file Move
        $file = Get-CurrentFile $sourcePane
        
        if ($file -eq $null) {
            Show-Modal "Move" "No file selected.`n`nPlace cursor on a file and press F6."
            return
        }
        
        $filesToMove = @($file)
    }
    
    Debug-Log "Files to Move: $($filesToMove.Count)"
    
    # If Moving to same directory and single file, prompt for new name
    if ($sourcePane.Path -eq $destPane.Path -and $filesToMove.Count -eq 1) {
        Debug-Log "Same directory, single file - prompting for new name"
        
        $fileName = Split-Path -Leaf $filesToMove[0]
        
        $dlg = [Terminal.Gui.Dialog]::new("Move to Same Directory", 70, 12)
        $lbl1 = [Terminal.Gui.Label]::new(1, 1, "Moving in same directory requires a new name:")
        $lbl2 = [Terminal.Gui.Label]::new(1, 2, "Original: $fileName")
        $lbl3 = [Terminal.Gui.Label]::new(1, 4, "New name:")
        $txt = [Terminal.Gui.TextField]::new(1, 5, 66, $fileName)
        $btnOK = [Terminal.Gui.Button]::new(18, 8, "Move")
        $btnCancel = [Terminal.Gui.Button]::new(36, 8, "Cancel")
        
        $script:DialogResult = $null
        
        $btnOK.add_Clicked({
            $newName = $txt.Text.ToString().Trim()
            
            if ([string]::IsNullOrWhiteSpace($newName)) {
                [Terminal.Gui.MessageBox]::ErrorQuery("Invalid Name", "File name cannot be empty", @("OK")) | Out-Null
                return
            }
            
            if ($newName -eq $fileName) {
                [Terminal.Gui.MessageBox]::ErrorQuery("Same Name", "New name must be different from original", @("OK")) | Out-Null
                return
            }
            
            $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
            if ($newName.IndexOfAny($invalidChars) -ge 0) {
                [Terminal.Gui.MessageBox]::ErrorQuery("Invalid Name", "File name contains invalid characters", @("OK")) | Out-Null
                return
            }
            
            $script:DialogResult = $newName
            [Terminal.Gui.Application]::RequestStop()
        })
        
        $btnCancel.add_Clicked({ 
            $script:DialogResult = $null
            [Terminal.Gui.Application]::RequestStop() 
        })
        
        $dlg.Add($lbl1)
        $dlg.Add($lbl2)
        $dlg.Add($lbl3)
        $dlg.Add($txt)
        $dlg.AddButton($btnOK)
        $dlg.AddButton($btnCancel)
        
        [Terminal.Gui.Application]::Run($dlg)
        
        if ($script:DialogResult -eq $null) {
            Debug-Log "User cancelled rename"
            return
        }
        
        $destFileName = $script:DialogResult
        $dest = Join-Path $destPane.Path $destFileName
        
        if (Test-Path -LiteralPath $dest) {
            $result = [Terminal.Gui.MessageBox]::Query("File Exists", "Destination file already exists:`n$dest`n`nOverwrite?", @("Yes", "No"))
            if ($result -ne 0) {
                return
            }
        }
        
        $msg = "Move file:`n  FROM: $($filesToMove[0])`n  TO: $dest"
        $result = [Terminal.Gui.MessageBox]::Query("Confirm Move", $msg, @("Yes", "No"))
        
        if ($result -eq 0) {
            try {
                Move-Item -LiteralPath $filesToMove[0] -Destination $dest -Force
                Refresh-Pane $destPane
                Show-Modal "Success" "File Moved to:`n$destFileName"
            } catch {
                Show-Modal "Error" "Move failed:`n$($_.Exception.Message)"
            }
        }
        
        return
    }
    
    # Multi-file Move or different directory
    if ($sourcePane.Path -eq $destPane.Path -and $filesToMove.Count -gt 1) {
        Show-Modal "Move Error" "Cannot Move multiple files to same directory!`n`nUse different destination pane."
        return
    }
    
    # Build confirmation message
    if ($filesToMove.Count -eq 1) {
        $fileName = Split-Path -Leaf $filesToMove[0]
        $msg = "Move 1 file:`n  $fileName`n`nTO: $($destPane.Path)"
    } else {
        $fileList = ($filesToMove | ForEach-Object { "  • $(Split-Path -Leaf $_)" } | Select-Object -First 5) -join "`n"
        if ($filesToMove.Count -gt 5) {
            $fileList += "`n  ... and $($filesToMove.Count - 5) more"
        }
        $msg = "Move $($filesToMove.Count) files:`n$fileList`n`nTO: $($destPane.Path)"
    }
    
    $result = [Terminal.Gui.MessageBox]::Query("Confirm Move", $msg, @("Yes", "No"))
    
    if ($result -ne 0) {
        return
    }
    
    # Show progress bar for multiple files (3 or more)
    if ($filesToMove.Count -ge 3) {
        $MovedCount = 0
        $errorCount = 0
        $overwriteAll = $false
        $skipAll = $false
        
        Show-ProgressDialog -Title "Moving Files" -Total $filesToMove.Count -Operation {
            for ($i = 0; $i -lt $filesToMove.Count; $i++) {
                $file = $filesToMove[$i]
                $fileName = Split-Path -Leaf $file
                $dest = Join-Path $destPane.Path $fileName
                
                Update-Progress -Current ($i + 1) -Status "Moving: $fileName"
                
                # Check if destination exists
                if (Test-Path -LiteralPath $dest) {
                    if (-not $overwriteAll -and -not $skipAll) {
                        $result = [Terminal.Gui.MessageBox]::Query("File Exists", "File exists:`n$fileName`n`nOverwrite?", @("Yes", "Yes to All", "No", "No to All"))
                        
                        if ($result -eq 1) { $overwriteAll = $true }
                        elseif ($result -eq 3) { $skipAll = $true; continue }
                        elseif ($result -eq 2) { continue }
                    } elseif ($skipAll) {
                        continue
                    }
                }
                
                try {
                    Move-Item -LiteralPath $file -Destination $dest -Force
                    $MovedCount++
                    Debug-Log "Moved: $fileName"
                } catch {
                    $errorCount++
                    Debug-Log "Error Moving $fileName : $($_.Exception.Message)"
                }
                
                Start-Sleep -Milliseconds 50  # Small delay to see progress
            }
        }
        
        Refresh-Pane $destPane
        
        if ($sourcePane.PSObject.Properties['SelectedFiles']) {
            $sourcePane.SelectedFiles.Clear()
            $sourcePane.Frame.Title = "$($sourcePane.Name) : $($sourcePane.Path)"
        }
        
        if ($errorCount -eq 0) {
            Show-Modal "Success" "Moved $MovedCount file(s) successfully"
        } else {
            Show-Modal "Partial Success" "Moved: $MovedCount`nFailed: $errorCount"
        }
        
        return
    }
    
    # Single or few files - no progress bar needed
    $MovedCount = 0
    $errorCount = 0
    $overwriteAll = $false
    $skipAll = $false
    
    foreach ($file in $filesToMove) {
        $fileName = Split-Path -Leaf $file
        $dest = Join-Path $destPane.Path $fileName
        
        if (Test-Path -LiteralPath $dest) {
            if (-not $overwriteAll -and -not $skipAll) {
                $result = [Terminal.Gui.MessageBox]::Query("File Exists", "File exists:`n$fileName`n`nOverwrite?", @("Yes", "Yes to All", "No", "No to All"))
                
                if ($result -eq 1) { $overwriteAll = $true }
                elseif ($result -eq 3) { $skipAll = $true; continue }
                elseif ($result -eq 2) { continue }
            } elseif ($skipAll) {
                continue
            }
        }
        
        try {
            Move-Item -LiteralPath $file -Destination $dest -Force
            $MovedCount++
            Debug-Log "Moved: $fileName"
        } catch {
            $errorCount++
            Debug-Log "Error Moving $fileName : $($_.Exception.Message)"
        }
    }
    
    Refresh-Pane $destPane
    
    if ($sourcePane.PSObject.Properties['SelectedFiles']) {
        $sourcePane.SelectedFiles.Clear()
        $sourcePane.Frame.Title = "$($sourcePane.Name) : $($sourcePane.Path)"
    }
    
    if ($errorCount -eq 0) {
        Show-Modal "Success" "Moved $MovedCount file(s) successfully"
    } else {
        Show-Modal "Partial Success" "Moved: $MovedCount`nFailed: $errorCount"
    }
}

function Delete-CurrentFile {
    param($pane)
    
    Debug-Log "=== DELETE ==="
    
    # Check if there are selected files
    $filesToDelete = @()
    
    if ($pane.PSObject.Properties['SelectedFiles'] -and $pane.SelectedFiles.Count -gt 0) {
        # Multi-file delete
        Debug-Log "Multi-file delete: $($pane.SelectedFiles.Count) files selected"
        
        foreach ($fileName in $pane.SelectedFiles) {
            $filePath = Join-Path $pane.Path $fileName
            if (Test-Path -LiteralPath $filePath -PathType Leaf) {
                $filesToDelete += $filePath
            }
        }
        
        if ($filesToDelete.Count -eq 0) {
            Show-Modal "Delete" "No valid files in selection"
            return
        }
    } else {
        # Single file delete
        $file = Get-CurrentFile $pane
        
        if ($file -eq $null) {
            Show-Modal "Delete" "No file selected.`n`nPlace cursor on a file and press F8."
            return
        }
        
        $filesToDelete = @($file)
    }
    
    Debug-Log "Files to delete: $($filesToDelete.Count)"
    
    # Build confirmation message
    if ($filesToDelete.Count -eq 1) {
        $fileName = Split-Path -Leaf $filesToDelete[0]
        $msg = "Delete 1 file?`n  $fileName`n`nCANNOT BE UNDONE!"
    } else {
        $fileList = ($filesToDelete | ForEach-Object { "  • $(Split-Path -Leaf $_)" } | Select-Object -First 8) -join "`n"
        if ($filesToDelete.Count -gt 8) {
            $fileList += "`n  ... and $($filesToDelete.Count - 8) more"
        }
        $msg = "Delete $($filesToDelete.Count) files?`n$fileList`n`nCANNOT BE UNDONE!"
    }
    
    $result = [Terminal.Gui.MessageBox]::Query("Confirm Delete", $msg, @("Yes, Delete", "No, Cancel"))
    
    if ($result -ne 0) {
        Debug-Log "User cancelled delete"
        return
    }
    
    # Perform the delete operation
    $deletedCount = 0
    $errorCount = 0
    $errorMessages = [System.Collections.Generic.List[string]]::new()
    
    foreach ($file in $filesToDelete) {
        $fileName = Split-Path -Leaf $file
        
        try {
            Remove-Item -LiteralPath $file -Force
            $deletedCount++
            Debug-Log "Deleted: $fileName"
        } catch {
            $errorCount++
            $errorMessages.Add("$fileName : $($_.Exception.Message)")
            Debug-Log "Error deleting $fileName : $($_.Exception.Message)"
        }
    }
    
    Refresh-Pane $pane
    
    # Clear selection after delete
    if ($pane.PSObject.Properties['SelectedFiles']) {
        $pane.SelectedFiles.Clear()
        $pane.Frame.Title = "$($pane.Name) : $($pane.Path)"
    }
    
    if ($errorCount -eq 0) {
        Show-Modal "Success" "Deleted $deletedCount file(s) successfully"
    } else {
        # Show first few errors
        $errorSummary = ($errorMessages | Select-Object -First 3) -join "`n"
        if ($errorMessages.Count -gt 3) {
            $errorSummary += "`n... and $($errorMessages.Count - 3) more errors"
        }
        Show-Modal "Partial Success" "Deleted: $deletedCount`nFailed: $errorCount`n`nErrors:`n$errorSummary"
    }
}

function Show-ChangeDirectoryDialog {
    param($pane)
    
    $dlg = [Terminal.Gui.Dialog]::new("Change Directory - $($pane.Name)", 80, 10)
    $lbl = [Terminal.Gui.Label]::new(1, 1, "Enter directory path:")
    $txt = [Terminal.Gui.TextField]::new(1, 3, 76, $pane.Path)
    $btnOK = [Terminal.Gui.Button]::new(20, 6, "OK")
    $btnCancel = [Terminal.Gui.Button]::new(36, 6, "Cancel")
    
    $btnOK.add_Clicked({
        $entered = $txt.Text.ToString().Trim()
        
        if ([string]::IsNullOrWhiteSpace($entered)) { 
            return 
        }
        
        if ($entered.StartsWith("~")) { 
            $entered = Join-Path $HOME ($entered.Substring(1).TrimStart('\', '/'))
        }
        
        $entered = [System.Environment]::ExpandEnvironmentVariables($entered)
        
        if (Test-Path -LiteralPath $entered -PathType Container) {
            $pane.Path = (Resolve-Path -LiteralPath $entered).Path
            Refresh-Pane $pane
            [Terminal.Gui.Application]::RequestStop()
        } else {
            [Terminal.Gui.MessageBox]::ErrorQuery("Invalid", "Directory not found:`n$entered", @("OK")) | Out-Null
        }
    })
    
    $btnCancel.add_Clicked({ 
        [Terminal.Gui.Application]::RequestStop() 
    })
    
    $dlg.Add($lbl)
    $dlg.Add($txt)
    $dlg.AddButton($btnOK)
    $dlg.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($dlg)
}

function New-DirectoryDialog {
    param($pane)
    
    Debug-Log "=== Create Directory ==="
    
    $dlg = [Terminal.Gui.Dialog]::new("Create Directory - $($pane.Name)", 70, 10)
    $lbl = [Terminal.Gui.Label]::new(1, 1, "Enter new directory name:")
    $txt = [Terminal.Gui.TextField]::new(1, 3, 66, "")
    $btnOK = [Terminal.Gui.Button]::new(18, 6, "Create")
    $btnCancel = [Terminal.Gui.Button]::new(36, 6, "Cancel")
    
    $btnOK.add_Clicked({
        $dirName = $txt.Text.ToString().Trim()
        
        if ([string]::IsNullOrWhiteSpace($dirName)) {
            [Terminal.Gui.MessageBox]::ErrorQuery("Invalid Name", "Directory name cannot be empty", @("OK")) | Out-Null
            return
        }
        
        $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
        if ($dirName.IndexOfAny($invalidChars) -ge 0) {
            [Terminal.Gui.MessageBox]::ErrorQuery("Invalid Name", "Directory name contains invalid characters", @("OK")) | Out-Null
            return
        }
        
        $newPath = Join-Path $pane.Path $dirName
        
        if (Test-Path -LiteralPath $newPath) {
            [Terminal.Gui.MessageBox]::ErrorQuery("Already Exists", "A directory with this name already exists", @("OK")) | Out-Null
            return
        }
        
        try {
            New-Item -Path $newPath -ItemType Directory -Force | Out-Null
            Refresh-Pane $pane
            [Terminal.Gui.Application]::RequestStop()
            Show-Modal "Success" "Directory created:`n$dirName"
        } catch {
            [Terminal.Gui.MessageBox]::ErrorQuery("Create Failed", "Failed to create directory:`n$($_.Exception.Message)", @("OK")) | Out-Null
        }
    })
    
    $btnCancel.add_Clicked({ 
        [Terminal.Gui.Application]::RequestStop() 
    })
    
    $dlg.Add($lbl)
    $dlg.Add($txt)
    $dlg.AddButton($btnOK)
    $dlg.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($dlg)
}

function Remove-DirectoryDialog {
    param($pane)
    
    Debug-Log "=== Delete Directory ==="
    
    $idx = $pane.ListView.SelectedItem
    $sourceList = $pane.ListView.Source.ToList()
    
    if ($idx -lt 0 -or $idx -ge $sourceList.Count) {
        Show-Modal "Delete Directory" "No item selected"
        return
    }
    
    $item = $sourceList[$idx]
    
    if ($item -eq ".." -or -not $item.StartsWith("[DIR] ")) {
        Show-Modal "Delete Directory" "Please select a directory.`n`nDirectories are shown as [DIR] name"
        return
    }
    
    $dirName = $item.Substring(6)
    $dirPath = Join-Path $pane.Path $dirName
    
    if (-not (Test-Path -LiteralPath $dirPath -PathType Container)) {
        Show-Modal "Error" "Directory not found"
        return
    }
    
    $contents = Get-ChildItem -LiteralPath $dirPath -Force -ErrorAction SilentlyContinue
    $hasContents = $contents.Count -gt 0
    
    if ($hasContents) {
        $msg = "Directory contains $($contents.Count) item(s):`n  $dirPath`n`nDelete directory and ALL contents?`nTHIS CANNOT BE UNDONE!"
        $result = [Terminal.Gui.MessageBox]::Query("Delete Directory (NOT EMPTY)", $msg, @("Yes, Delete All", "No, Cancel"))
        
        if ($result -ne 0) {
            return
        }
    } else {
        $msg = "Delete empty directory?`n  $dirPath"
        $result = [Terminal.Gui.MessageBox]::Query("Delete Directory", $msg, @("Yes", "No"))
        
        if ($result -ne 0) {
            return
        }
    }
    
    try {
        Remove-Item -LiteralPath $dirPath -Recurse:$hasContents -Force
        Refresh-Pane $pane
        Show-Modal "Success" "Directory deleted"
    } catch {
        Show-Modal "Error" "Failed to delete directory:`n$($_.Exception.Message)"
    }
}

function Show-ProgressDialog {
    param(
        [string]$Title,
        [int]$Total,
        [scriptblock]$Operation
    )
    
    Debug-Log "Show-ProgressDialog: Starting with $Total items"
    
    $dlg = [Terminal.Gui.Dialog]::new($Title, 70, 10)
    
    $lblStatus = [Terminal.Gui.Label]::new(1, 1, "Starting...")
    $lblProgress = [Terminal.Gui.Label]::new(1, 3, "0 / $Total")
    
    $progressBar = [Terminal.Gui.Label]::new(1, 5, "")
    $progressBar.Width = 66
    
    $dlg.Add($lblStatus)
    $dlg.Add($lblProgress)
    $dlg.Add($progressBar)
    
    # Store progress state
    $script:ProgressState = @{
        Current = 0
        Total = $Total
        StatusLabel = $lblStatus
        ProgressLabel = $lblProgress
        ProgressBar = $progressBar
        Dialog = $dlg
    }
    
    # Schedule the operation to run after dialog is displayed
    [Terminal.Gui.Application]::MainLoop.AddTimeout([TimeSpan]::FromMilliseconds(100), {
        Debug-Log "Show-ProgressDialog: Timeout fired, starting operation"
        try {
            & $Operation
            Debug-Log "Show-ProgressDialog: Operation completed"
        } catch {
            Debug-Log "Progress operation error: $($_.Exception.Message)"
            Debug-Log "Stack trace: $($_.ScriptStackTrace)"
        }
        
        # Close dialog after operation completes
        $script:ProgressState.Dialog.Running = $false
        return $false  # Don't repeat timeout
    })
    
    Debug-Log "Show-ProgressDialog: Running dialog"
    [Terminal.Gui.Application]::Run($dlg)
    Debug-Log "Show-ProgressDialog: Dialog closed"
}

function Update-Progress {
    param(
        [int]$Current,
        [string]$Status
    )
    
    Debug-Log "Update-Progress: $Current/$($script:ProgressState.Total) - $Status"
    
    if ($script:ProgressState -eq $null) { 
        Debug-Log "Update-Progress: ProgressState is null!"
        return 
    }
    
    $script:ProgressState.Current = $Current
    $percent = [Math]::Round(($Current / $script:ProgressState.Total) * 100)
    
    $script:ProgressState.StatusLabel.Text = $Status
    $script:ProgressState.ProgressLabel.Text = "$Current / $($script:ProgressState.Total) ($percent%)"
    
    $barWidth = 66
    $filled = [Math]::Floor($barWidth * $Current / $script:ProgressState.Total)
    $empty = $barWidth - $filled
    $bar = ("█" * $filled) + ("░" * $empty)
    $script:ProgressState.ProgressBar.Text = $bar
    
    $script:ProgressState.Dialog.SetNeedsDisplay()
    
    # Process one UI event to allow the display to update
    [Terminal.Gui.Application]::RunIteration([ref]$false)
    
    Debug-Log "Update-Progress: UI updated"
}

function Show-SelectedFiles {
    param($pane)
    
    Debug-Log "=== Show Selection ==="
    
    if (-not $pane.PSObject.Properties['SelectedFiles'] -or $pane.SelectedFiles.Count -eq 0) {
        Show-Modal "Selection" "No files selected in $($pane.Name) pane"
        return
    }
    
    # Create a dialog with a ListView
    $dlg = [Terminal.Gui.Dialog]::new("Selected Files - $($pane.Name)", 80, 25)
    
    $lbl = [Terminal.Gui.Label]::new(1, 1, "$($pane.SelectedFiles.Count) file(s) selected:")
    
    # Create ListView to show selected files
    $listView = [Terminal.Gui.ListView]::new()
    $listView.X = 1
    $listView.Y = 3
    $listView.Width = [Terminal.Gui.Dim]::Fill(1)
    $listView.Height = [Terminal.Gui.Dim]::Fill(3)
    
    # Populate the list
    $fileList = [System.Collections.Generic.List[string]]::new()
    foreach ($file in $pane.SelectedFiles) {
        $fileList.Add($file)
    }
    $listView.SetSource($fileList)
    
    $btnClose = [Terminal.Gui.Button]::new(35, [Terminal.Gui.Pos]::Bottom($dlg) - 2, "Close")
    $btnClose.add_Clicked({
        [Terminal.Gui.Application]::RequestStop()
    })
    
    $dlg.Add($lbl)
    $dlg.Add($listView)
    $dlg.AddButton($btnClose)
    
    [Terminal.Gui.Application]::Run($dlg)
}

function Invoke-Rename {
    # Get current pane
    $p = Get-FocusedPane

    # Get selected index
    $idx = $p.Active.ListView.SelectedItem
    if ($idx -lt 0) {
        Show-Modal "Rename" "No file selected."
        return
    }

    # Get selected name
    $sourceList = $p.Active.ListView.Source.ToList()
    if ($idx -ge $sourceList.Count) { return }
    $oldName = $sourceList[$idx]

    # Ignore ".."
    if ($oldName -eq "..") {
        Show-Modal "Rename" "Cannot rename parent directory shortcut."
        return
    }

    # Build full path
    $dir = $p.Active.Path
    $oldFull = Join-Path ${dir} ${oldName}

    # Check existence
    if (-not (Test-Path ${oldFull})) {
        Show-Modal "Error" "Item no longer exists."
        return
    }

    # ---- Show rename dialog ----
    $dlg = [Terminal.Gui.Dialog]::new("Rename", 60, 10)

    $lbl = [Terminal.Gui.Label]::new(1, 1, "New name:")
    $txt = [Terminal.Gui.TextField]::new(1, 2, 55, $oldName)

    # OK / Cancel buttons
    $okBtn = [Terminal.Gui.Button]::new("OK")
    $cancelBtn = [Terminal.Gui.Button]::new("Cancel")

    $okBtn.add_Click({
        $newName = $txt.Text.ToString().Trim()
        if ([string]::IsNullOrWhiteSpace($newName)) {
            Show-Modal "Rename" "Name cannot be empty."
            return
        }

        $newFull = Join-Path ${dir} ${newName}

        # --- Prevent overwrite if target exists ---
        if (Test-Path ${newFull}) {
            Show-Modal "Rename" "A file or folder with that name already exists.`n`nRename cancelled."
            return
        }

        # --- Attempt rename ---
        try {
            Rename-Item -LiteralPath ${oldFull} -NewName ${newName} -ErrorAction Stop
        }
        catch {
            Show-Modal "Error" "Rename failed: $($_.Exception.Message)"
            return
        }

        # Close dialog
        $dlg.RequestStop()

        # Refresh pane
        Refresh-Pane $p.Active
    })

    $cancelBtn.add_Click({
        $dlg.RequestStop()
    })

    # Add controls
    $dlg.Add($lbl)
    $dlg.Add($txt)
    $dlg.AddButton($okBtn)
    $dlg.AddButton($cancelBtn)

    [Terminal.Gui.Application]::Run($dlg)

    # Refocus pane after closing
    $p = Get-FocusedPane
    $p.Active.ListView.SetFocus()
}


function Get-FileInfo {
    param($pane)

    $file = Get-CurrentFile $pane

    if ($file -eq $null) {
        Show-Modal "File Info" "`u{200C}No file selected."
        return
    }

    try {
        $item = Get-Item -LiteralPath $file -ErrorAction Stop

        # ---- Friendly size conversion ----
        function Convert-FriendlySize {
            param([long]$bytes)

            switch ($bytes) {
                {$_ -ge 1TB} { return "{0:N2} TB" -f ($bytes / 1TB) }
                {$_ -ge 1GB} { return "{0:N2} GB" -f ($bytes / 1GB) }
                {$_ -ge 1MB} { return "{0:N2} MB" -f ($bytes / 1MB) }
                {$_ -ge 1KB} { return "{0:N2} KB" -f ($bytes / 1KB) }
                default      { return "$bytes bytes" }
            }
        }

        $friendlySize = Convert-FriendlySize -bytes $item.Length

        # Left-align hack: prefix each line with ZERO-WIDTH NON-JOINER (U+200C)
        $prefix = "`u{200C}"

        $msg = @(
            "$prefix Name:        $($item.Name)"
            "$prefix Full Path:   $($item.FullName)"
            "$prefix Size:        $($item.Length) bytes ($friendlySize)"
            "$prefix Extension:   $($item.Extension)"
            "$prefix Created:     $($item.CreationTime)"
            "$prefix Modified:    $($item.LastWriteTime)"
            "$prefix Attributes:  $($item.Attributes)"
        ) -join "`n"

        Show-Modal "File Info" $msg

    } catch {
        Show-Modal "Error" "`u{200C}Unable to read file info:`n`u{200C}$($_.Exception.Message)"
    }
}

function Get-FocusedPane {
    if ($script:CurrentFocusPane -eq 'RIGHT') {
        Debug-Log "Focus: RIGHT"
        Update-FocusIndicator
        return @{ Active = $script:RightPane; Other = $script:LeftPane }
    } else {
        Debug-Log "Focus: LEFT"
        Update-FocusIndicator
        return @{ Active = $script:LeftPane; Other = $script:RightPane }
    }
}

function Update-FocusIndicator {
    if ($script:CurrentFocusPane -eq 'RIGHT') {
        $script:FocusStatusItem.Title = " Active: RIGHT "
    } else {
        $script:FocusStatusItem.Title = " Active: LEFT "
    }

    # Force redraw
    [Terminal.Gui.Application]::Refresh()
}

# Mehod to the madness
function Show-PineappleInfo {
    $dlg = [Terminal.Gui.Dialog]::new("Why Pineapple? 🍍", 60, 12)
    
    $message = @"
PSMC is codenamed "Pineapple" because:

- I was drinking pineapple soda when writing the code.
- The Danish word for pineapple is Ananas which I find
  amusing.
- Pizza - with(out) pineapple can be ordered from the
  command line, so this is a call back.
- Every great project needs a tropical mascot!
"@
    
    $label = [Terminal.Gui.Label]::new(1, 1, $message)
    $dlg.Add($label)
    
   [Terminal.Gui.MessageBox]::Query("Why Pineapple? 🍍", $message, @("OK"))
}

# Main UI Setup
$script:ThemeMode = 'dark'
$script:CurrentFocusPane = 'LEFT'

$start = (Get-Location).Path
Debug-Log "=== PSMC Starting ==="

$script:LeftPane = New-FilePane -initialPath $start -themeMode $script:ThemeMode -paneName 'LEFT'
$script:RightPane = New-FilePane -initialPath $start -themeMode $script:ThemeMode -paneName 'RIGHT'

$win = [Terminal.Gui.Window]::new("PSMC v$($Global:PSMC_Version) - Pineapple Build")
$win.X = 0
$win.Y = 1
$win.Width = [Terminal.Gui.Dim]::Fill()
$win.Height = [Terminal.Gui.Dim]::Fill()

$script:LeftPane.Frame.X = 0
$script:LeftPane.Frame.Width = [Terminal.Gui.Dim]::Percent(50)

$divider = [Terminal.Gui.LineView]::new()
$divider.X = [Terminal.Gui.Pos]::Right($script:LeftPane.Frame)
$divider.Y = 0
$divider.Width = 1
$divider.Height = [Terminal.Gui.Dim]::Fill()

$script:RightPane.Frame.X = [Terminal.Gui.Pos]::Right($divider)
$script:RightPane.Frame.Width = [Terminal.Gui.Dim]::Fill()

## Set up code to check current pane in status bar
## Status item to show which pane is active
$script:FocusStatusItem = [Terminal.Gui.StatusItem]::new(
    [Terminal.Gui.Key]::Null,       # No hotkey
    " Active: LEFT ",               # Default (will update)
    $null
)


# Status bar
$statusBar = [Terminal.Gui.StatusBar]::new(@(
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F1, "~F1~ Help", {
        Show-Modal "Shortcuts" "F1 - Help`nTAB - Switch Pane`nF2 - Rename`n/F3 - Mkdir`nF4 - Browse Dir`nF5 - Refresh`nF6 - Copy`nF7 - Change dir`nF8 - Delete`nF10 - Quit" 
    }),

    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::Tab, "~Tab~ Switch Pane", {

    # Toggle the focus pane
    if ($script:CurrentFocusPane -eq 'LEFT') {
        $script:CurrentFocusPane = 'RIGHT'
    } else {
        $script:CurrentFocusPane = 'LEFT'
    }

    # Update indicator
    Update-FocusIndicator

    # Get the pane object
    $p = Get-FocusedPane

    # --- Terminal.Gui 1.16: Set focus on the new pane's ListView ---
    $p.Active.ListView.SetFocus()

   }),

    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F2, "~F2~ Rename", {
        Invoke-Rename
    }),

    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F3, "~F3~ Mkdir", {
        Invoke-Rename
    }),

    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F4, "~F4~ Browse", {
        $p = Get-FocusedPane
        $idx = $p.Active.ListView.SelectedItem
        if ($idx -lt 0) { return }
        
        # Get the selected item
        $sourceList = $p.Active.ListView.Source.ToList()
        if ($idx -ge $sourceList.Count) { return }
        $name = $sourceList[$idx]
        
        # Handle parent directory (..)
        if ($name -eq "..") {
            try {
                $parent = [System.IO.Directory]::GetParent($p.Active.Path)
                if ($parent -ne $null) {
                    $p.Active.Path = $parent.FullName
                    Refresh-Pane $p.Active
                }
            } catch {
                Show-Modal "Error" "Cannot navigate to parent"
            }
            return
        }
        
        # Handle directories (strip [DIR] prefix)
        if ($name.StartsWith("[DIR] ")) {
            $dirName = $name.Substring(6)
            $fullPath = Join-Path $p.Active.Path $dirName
            
            if (Test-Path -LiteralPath $fullPath -PathType Container) {
                $p.Active.Path = (Resolve-Path -LiteralPath $fullPath).Path
                Refresh-Pane $p.Active
            } else {
                Show-Modal "Error" "Directory not found:`n$fullPath"
            }
            return
        }
        
        # If it's a file, show file info
        Get-FileInfo $p.Active
    }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F5, "~F5~ Redraw", {
        $p = Get-FocusedPane
        Refresh-Pane $p.Active
    }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F6, "~F6~ Copy", {
        $p = Get-FocusedPane
        Copy-CurrentFile $p.Active $p.Other
    }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F7, "~F7~ Cwd", {
        $p = Get-FocusedPane
        Show-ChangeDirectoryDialog $p.Active
    }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F8, "~F8~ Del", {
        $p = Get-FocusedPane
        Delete-CurrentFile $p.Active
    }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F10, "~F10~ Quit", {
        [Terminal.Gui.Application]::RequestStop()
    })

    # <-- The dynamic indicator here must be LAST -->
    $script:FocusStatusItem

))

# Menus
$menuFile = [Terminal.Gui.MenuBarItem]::new("_File", @(
    [Terminal.Gui.MenuItem]::new("E_xit (F10)", "Quit", [Action]{ 
        [Terminal.Gui.Application]::RequestStop() 
    })
))

$menuActions = [Terminal.Gui.MenuBarItem]::new("_Actions", @(
    [Terminal.Gui.MenuItem]::new("_Move File", "Move file", [Action]{
        $p = Get-FocusedPane
        Move-CurrentFile $p.Active $p.Other
    }),
    [Terminal.Gui.MenuItem]::new("_Rename File (F2)", "Rename file", [Action]{
        Invoke-Rename
    }),
    [Terminal.Gui.MenuItem]::new("Create _Directory (F3)", "Create directory", [Action]{
        $p = Get-FocusedPane
        New-DirectoryDialog $p.Active
    }),
    [Terminal.Gui.MenuItem]::new("Delete D_irectory", "Delete directory", [Action]{
        $p = Get-FocusedPane
        Remove-DirectoryDialog $p.Active
    }),
    [Terminal.Gui.MenuItem]::new("_Copy File (F6)", "Copy file", [Action]{
        $p = Get-FocusedPane
        Copy-CurrentFile $p.Active $p.Other
    }),
    [Terminal.Gui.MenuItem]::new("_Delete File (F8)", "Delete file", [Action]{
        $p = Get-FocusedPane
        Delete-CurrentFile $p.Active
    }),
    [Terminal.Gui.MenuItem]::new("File _Info", "Show file info", [Action]{
    $p = Get-FocusedPane
    Get-FileInfo $p.active
    }),
    [Terminal.Gui.MenuItem]::new("_Select by Pattern", "Select files by pattern", [Action]{
        $p = Get-FocusedPane
        Select-FilesByPattern $p.Active
    }),
    [Terminal.Gui.MenuItem]::new("Select _All Files", "Select all files in current pane", [Action]{
        $p = Get-FocusedPane
        Select-AllFiles $p.Active
    }),
    [Terminal.Gui.MenuItem]::new("_Unselect All", "Clear all selections", [Action]{
        $p = Get-FocusedPane
        Unselect-AllFiles $p.Active
    }),
    [Terminal.Gui.MenuItem]::new("Sho_w Selection", "View selected files", [Action]{
        $p = Get-FocusedPane
        Show-SelectedFiles $p.Active
    })

))

$menuNav = [Terminal.Gui.MenuBarItem]::new("_Navigate", @(
    [Terminal.Gui.MenuItem]::new("_Browse (F4)", "Browse", [Action]{
        $p = Get-FocusedPane
        try {
            $parent = [System.IO.Directory]::GetParent($p.Active.Path)
            if ($parent) {
                $p.Active.Path = $parent.FullName
                Refresh-Pane $p.Active
                Show-Modal "Navigation" "Moved to: $($parent.FullName)"
            }
        } catch {
            Show-Modal "Error" "Cannot navigate to parent"
        }
    }),
    [Terminal.Gui.MenuItem]::new("_Redraw (F5)", "Refresh pane", [Action]{
        $p = Get-FocusedPane
        Refresh-Pane $p.Active
    }),
    [Terminal.Gui.MenuItem]::new("_Change Directory (F7)", "Change Dir", [Action]{
        $p = Get-FocusedPane
        Show-ChangeDirectoryDialog $p.Active
    })
))

$menuHelp = [Terminal.Gui.MenuBarItem]::new("_Help", @(
    [Terminal.Gui.MenuItem]::new("_Keys (F1)", "Shortcuts", [Action]{ 
        Show-Modal "Shortcuts" "F1 Help`nTab - Switch Pane`nF2/F3 -Mkdir`nF4 - Go up`nF5 - Refresh`nF6 - Copy`nF7 - Change dir`nF8 - Delete`nF10 - Quit" 
    }),
    [Terminal.Gui.MenuItem]::new("_About", "About", [Action]{ 
        Show-Modal "About" "PSMC v$($Global:PSMC_Version) STABLE`nGPL-3 Copyleft`nBy Knightmare2600 (https://github.com/knightmare2600" 
    }),,
    [Terminal.Gui.MenuItem]::new("Why _Pineapple?", "", {
        Show-PineappleInfo
    })

))

$menu = [Terminal.Gui.MenuBar]::new(@($menuFile, $menuActions, $menuNav, $menuHelp))

[Terminal.Gui.Application]::Top.Add($menu)
[Terminal.Gui.Application]::Top.Add($win)
[Terminal.Gui.Application]::Top.Add($statusBar)

$win.Add($script:LeftPane.Frame)
$win.Add($divider)
$win.Add($script:RightPane.Frame)

$script:LeftPane.Frame.ColorScheme = Get-Theme $script:ThemeMode
$script:RightPane.Frame.ColorScheme = Get-Theme $script:ThemeMode

try { 
    [Terminal.Gui.Application]::SetFocus($script:LeftPane.ListView) 
} catch {}

Debug-Log "=== PSMC v$($Global:PSMC_Version) Ready ==="
Debug-Log "File | Actions | Navigate | Help"

[Terminal.Gui.Application]::Run()
[Terminal.Gui.Application]::Shutdown()
