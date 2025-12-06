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

# -------------------- Load Terminal.Gui --------------------
if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Terminal.Gui' })) {
    $mod = Get-Module Microsoft.PowerShell.ConsoleGuiTools -List
    if (-not $mod) { throw "Module Microsoft.PowerShell.ConsoleGuiTools not found. Install it and retry." }
    $dll = Join-Path $mod.ModuleBase 'Terminal.Gui.dll'
    if (-not (Test-Path $dll)) { throw "Terminal.Gui.dll not found at $dll" }
    Add-Type -Path $dll
}

[Terminal.Gui.Application]::Init()

# -------------------- Helpers --------------------
function Debug-Log([string]$m) {
    if ($using:Debug) { Write-Host "DEBUG: $m" -ForegroundColor Cyan }
}

function Show-Modal([string]$title, [string]$msg) {
    [Terminal.Gui.MessageBox]::Query($title, $msg, @("OK")) | Out-Null
}

function Get-Theme([string]$mode) {
    $cs = [Terminal.Gui.ColorScheme]::new()
    if ($mode -eq 'light') {
        $cs.Normal   = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black, [Terminal.Gui.Color]::Gray)
        $cs.Focus    = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black, [Terminal.Gui.Color]::Cyan)
    } else {
        $cs.Normal   = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray, [Terminal.Gui.Color]::Black)
        $cs.Focus    = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black, [Terminal.Gui.Color]::Gray)
    }
    return $cs
}

function Build-DisplayNames([System.IO.FileSystemInfo[]]$items, [string]$path) {
    $names = [System.Collections.Generic.List[string]]::new()
    $parent = Split-Path -Parent $path
    if ($parent -and $parent -ne '') { $names.Add("..") }

    $dirs = $items | Where-Object { $_.PSIsContainer } | Sort-Object Name
    foreach ($d in $dirs) { $names.Add("[DIR] $($d.Name)") }
    $files = $items | Where-Object { -not $_.PSIsContainer } | Sort-Object Name
    foreach ($f in $files) { $names.Add($f.Name) }

    return ,([string[]]$names)
}

function Get-ListViewSourceArray($listView) {
    $src = $listView.Source
    try {
        if ($src -is [Terminal.Gui.ListWrapper]) {
            return ,([string[]]($src.ToList()))
        }
    } catch { }
    try {
        $arr = @()
        foreach ($s in $src) { $arr += [string]$s }
        return ,([string[]]$arr)
    } catch {
        return ,([string[]]@([string]$src))
    }
}

# -------------------- Pane creation and refresh --------------------
function New-FilePane([string]$initialPath, [string]$themeMode='dark') {
    $pane = [PSCustomObject]@{
        Path     = (Resolve-Path $initialPath).Path
        Items    = @()
        Frame    = $null
        ListView = $null
    }

    $pane.Items = Get-ChildItem -LiteralPath $pane.Path -Force -ErrorAction SilentlyContinue
    $display  = Build-DisplayNames $pane.Items $pane.Path

    $frame = [Terminal.Gui.FrameView]::new()
    $frame.Title = $pane.Path
    $frame.Y = 0
    $frame.Height = [Terminal.Gui.Dim]::Fill()
    $frame.ColorScheme = Get-Theme $themeMode

    $list = [Terminal.Gui.ListView]::new([string[]]$display)
    $list.X = 0; $list.Y = 0
    $list.Width  = [Terminal.Gui.Dim]::Fill()
    $list.Height = [Terminal.Gui.Dim]::Fill()
    $list.CanFocus = $true

    $frame.Add($list)
    $pane.Frame = $frame
    $pane.ListView = $list

    Debug-Log "ListView created successfully for path: $($pane.Path)"
    return $pane
}

function Refresh-Pane($pane) {
    try {
        $pane.Items = Get-ChildItem -LiteralPath $pane.Path -Force -ErrorAction SilentlyContinue
        $names = Build-DisplayNames $pane.Items $pane.Path
        $pane.ListView.SetSource([string[]]$names)
        $pane.Frame.Title = $pane.Path
        Debug-Log "Pane refreshed for path: $($pane.Path)"
    } catch {
        Debug-Log "Refresh-Pane error: $($_.Exception.Message)"
    }
}

# -------------------- UI Setup --------------------
$top = [Terminal.Gui.Application]::Top
$win = [Terminal.Gui.Window]::new("PSMC - PowerShell Commander (restored)")
$win.X = 0; $win.Y = 1
$win.Width  = [Terminal.Gui.Dim]::Fill()
$win.Height = [Terminal.Gui.Dim]::Fill()

# Change directory dialog
function Show-ChangeDir {
    $current = (Get-Location).Path
    $input = [Terminal.Gui.TextField]::new($current)
    $input.Width = 50
    $d = [Terminal.Gui.Dialog]::new("Change Directory", 60, 10)
    $lbl = [Terminal.Gui.Label]::new("Enter new path:")
    $lbl.X = 1; $lbl.Y = 1
    $input.X = 1; $input.Y = 2
    $ok = [Terminal.Gui.Button]::new("OK")
    $cancel = [Terminal.Gui.Button]::new("Cancel")
    $ok.X = 10; $ok.Y = 5
    $cancel.X = 30; $cancel.Y = 5
    $d.Add($lbl); $d.Add($input); $d.Add($ok); $d.Add($cancel)
    $ok.add_Click({ $d.RequestStop() })
    $cancel.add_Click({ $input.Text = ""; $d.RequestStop() })
    [Terminal.Gui.Application]::Run($d)
    $path = $input.Text
    if ($path -and (Test-Path $path -PathType Container)) {
        $LeftPane.Path = $path
        $RightPane.Path = $path
        Refresh-Pane $LeftPane
        Refresh-Pane $RightPane
        Debug-Log "Changed directory to $path"
    } elseif ($path -ne "") {
        Show-Modal "Error" "Invalid path: $path"
    }
}

# Menu bar — restored and always visible
$menu = [Terminal.Gui.MenuBar]::new(@(
    [Terminal.Gui.MenuBarItem]::new("_File", @(
        [Terminal.Gui.MenuItem]::new("_Quit","Quit application",[Action]{ [Terminal.Gui.Application]::RequestStop() })
    )),
    [Terminal.Gui.MenuBarItem]::new("_Actions", @(
        [Terminal.Gui.MenuItem]::new("_Copy (F5)","Copy selected",[Action]{ Show-Modal "Copy" "Simulated copy" }),
        [Terminal.Gui.MenuItem]::new("_Move (F6)","Move selected",[Action]{ Show-Modal "Move" "Simulated move" }),
        [Terminal.Gui.MenuItem]::new("_Delete (F8)","Delete selected",[Action]{ Show-Modal "Delete" "Simulated delete" }),
        [Terminal.Gui.MenuItem]::new("_Change Dir (F9)","Change directory",[Action]{ Show-ChangeDir })
    )),
    [Terminal.Gui.MenuBarItem]::new("_View", @(
        [Terminal.Gui.MenuItem]::new("_Toggle Theme","Toggle dark/light theme",[Action]{
            if ($script:ThemeMode -eq 'dark') { $script:ThemeMode = 'light' } else { $script:ThemeMode = 'dark' }
            $LeftPane.Frame.ColorScheme  = Get-Theme $script:ThemeMode
            $RightPane.Frame.ColorScheme = Get-Theme $script:ThemeMode
            [Terminal.Gui.Application]::Refresh()
            Debug-Log "Theme toggled to $script:ThemeMode"
        })
    )),
    [Terminal.Gui.MenuBarItem]::new("_Help", @(
        [Terminal.Gui.MenuItem]::new("_About","About this tool",[Action]{ Show-Modal "About" "PSMC — demo TUI file manager" })
    ))
))

# initial theme
$script:ThemeMode = 'dark'

# Apply full theme to all components <-- do this BEFORE the menus
Apply-Theme -ThemeData $themeData -TopLevel $top -MainWindow $win -Menu $menu -Status $status

Debug-Log "Theme applied successfully: $Theme"

# create panes
$start = (Get-Location).Path
$LeftPane  = New-FilePane -initialPath $start -themeMode:$script:ThemeMode
$RightPane = New-FilePane -initialPath $start -themeMode:$script:ThemeMode

# layout
$LeftPane.Frame.X = 0
$LeftPane.Frame.Width = [Terminal.Gui.Dim]::Percent(50)

$divider = [Terminal.Gui.LineView]::new()
$divider.X = [Terminal.Gui.Pos]::Right($LeftPane.Frame)
$divider.Y = 0
$divider.Width = 1
$divider.Height = [Terminal.Gui.Dim]::Fill()
try {
    if ([Terminal.Gui.LineView].GetMember('Orientation')) {
        $divider.Orientation = [Terminal.Gui.LineView+Orientation]::Vertical
    } elseif ([Terminal.Gui.LineView].GetMember('Direction')) {
        $divider.Direction = [Terminal.Gui.LineView+Orientation]::Vertical
    }
} catch { Debug-Log "LineView orientation assignment skipped." }

$RightPane.Frame.X = [Terminal.Gui.Pos]::Right($divider)
$RightPane.Frame.Width = [Terminal.Gui.Dim]::Fill()

# add menu and window
[Terminal.Gui.Application]::Top.Add($menu)
[Terminal.Gui.Application]::Top.Add($win)
$win.Add($LeftPane.Frame)
$win.Add($divider)
$win.Add($RightPane.Frame)

# status bar
$StatusBar = [Terminal.Gui.StatusBar]::new(@(
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F5, "F5 Copy", { Show-Modal "Copy" "Simulated copy" }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F6, "F6 Move", { Show-Modal "Move" "Simulated move" }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F8, "F8 Delete", { Show-Modal "Delete" "Simulated delete" }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F9, "F9 ChDir", { Show-ChangeDir }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::CtrlMask -bor [Terminal.Gui.Key]::Q, "Exit", { [Terminal.Gui.Application]::RequestStop() })
))
[Terminal.Gui.Application]::Top.Add($StatusBar)

# apply theme
$LeftPane.Frame.ColorScheme  = Get-Theme $script:ThemeMode
$RightPane.Frame.ColorScheme = Get-Theme $script:ThemeMode

# focus and navigation logic
try { [Terminal.Gui.Application]::SetFocus($LeftPane.ListView) } catch {}

function Handle-OpenSelectedItem($pane, $side) {
    $selIndex = $pane.ListView.SelectedItem
    if ($selIndex -lt 0) { return }
    $src = Get-ListViewSourceArray $pane.ListView
    if ($selIndex -ge $src.Length) { return }
    $selText = ($src[$selIndex]).ToString().Trim()
    Debug-Log "OpenSelectedItem fired on ${side}: '${selText}'"

    if ($selText -eq '..') {
        $parent = Split-Path -Parent $pane.Path
        if ($parent -and (Test-Path $parent -PathType Container)) {
            $pane.Path = $parent
            Refresh-Pane $pane
            Debug-Log "${side}: Moved up to ${parent}"
        } else { Debug-Log "${side}: Already at root" }
        return
    }

    if ($selText -match '^\[DIR\]\s*(.+)$') {
        $dirName = $Matches[1]
        $target = Join-Path -Path $pane.Path -ChildPath $dirName
        if (Test-Path $target -PathType Container) {
            $pane.Path = $target
            Refresh-Pane $pane
            Debug-Log "${side}: Entered ${target}"
        }
        return
    }

    Debug-Log "${side}: File selected: ${selText}"
    Show-Modal "File" "Would open file: ${selText}"
}

$LeftPane.ListView.add_OpenSelectedItem({ Handle-OpenSelectedItem $LeftPane 'LEFT' })
$RightPane.ListView.add_OpenSelectedItem({ Handle-OpenSelectedItem $RightPane 'RIGHT' })

# Tab switches focus
$win.add_KeyPress({
    param($args)
    try { $kv = $args.KeyEvent.KeyValue } catch { $kv = $null }
    if ($kv -eq 9) {
        if ([Terminal.Gui.Application]::Focused -eq $LeftPane.ListView) {
            [Terminal.Gui.Application]::SetFocus($RightPane.ListView)
        } else {
            [Terminal.Gui.Application]::SetFocus($LeftPane.ListView)
        }
        $args.Handled = $true
    }
})

# -------------------- Start --------------------
Debug-Log "PSMC ready. Left: $($LeftPane.Path) Right: $($RightPane.Path)"
[Terminal.Gui.Application]::Run()
[Terminal.Gui.Application]::Shutdown()
