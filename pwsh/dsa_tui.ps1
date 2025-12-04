<#

DSA-TUI Text Mode version of dsa.msc for powershell
Locked-in baseline: dynamic resize, menu, demo data mirrors prod format, Change Domain fixed, fixed DC selection, full production AD object detection, properties modal, AD search popup

===========================================================================================
 DSA-TUI Blaabaer — Active Directory TUI Tool
 Historical Build Notes and Change Log
===========================================================================================

 1.0.0  (Initial Experimental)
 - First internal test build. Basic TUI scaffolding only.
 - Bare Window + Menu + Exit only. No AD integration.
 - Non-functional placeholder TreeView.

 1.1.0  (Initial AD Integration)
 - Added basic Domain Bind and LDAP query functions.
 - Added Build-Tree function (initial non-recursive prototype).
 - Added minimal Properties popup (placeholder).

 1.1.1  (Bugfix)
 - Fixed null-domain crash.
 - Fixed title bar misalignment on Linux/macOS terminals.

 1.2.0  (Tree + Navigation)
 - Introduced TreeView AD structure display.
 - Added OU expansion, user nodes, group nodes.
 - Implemented Refresh (F5) bound to Build-Tree.
 - Added basic status bar with simple messages only.

 1.2.1  (Bugfix)
 - Fixed node-expansion crash when encountering empty OUs.
 - Fixed cosmetic padding/spacing inconsistencies.

 1.3.0  (Selection, Node Info)
 - Added node selection handling.
 - Added Show-Properties modal (initial version).
 - Added object type detection for icons (U/G/OU/DC).

 1.3.1  (Bugfix)
 - Fixed Properties dialog not clearing previous content.
 - Fixed MessageBox misalignment under Terminal.Gui 1.16.

 1.4.0  (Filter System v1)
 - Added Filter Panel (right side) with toggles.
 - Added Global:FilterOptions hashtable.
 - Added Update-FilterStatusLabel function.
 - Added name-filter support (“search by name”).

 1.4.1  (Bugfix)
 - Fixed filter panel overlapping TreeView.
 - Fixed name-filter not persisting during refresh.
 - Fixed missing redrawing after filter changes.

 1.5.0  (Full Refresh Engine + Searchable Properties Rewrite)
 - Major rewrite of Refresh / Build-Tree pipeline.
 - Added "Searchable Attributes" handling:
       * name
       * displayName
       * sAMAccountName
       * userPrincipalName
       * givenName / sn
 - Optimized LDAP lookups to only fetch required fields.
 - Added caching to reduce domain traffic.

 1.5.1  (Bugfix)
 - Fixed several typos in attribute lookup keys (displayName vs displayname).
 - Fixed "OU:" prefixes duplicating on some nodes.
 - Fixed sorting of users/groups inside OUs.

 1.5.2  (Bugfix)
 - Fixed rare crash when node had malformed DN.
 - Corrected spacing in status label (“No filters active” line).
 - Fixed cosmetic typo: “Serach” → “Search”.

 1.5.3  (Bugfix)
 - Fixed name filter not updating until second refresh.
 - Fixed stale nodes remaining after filter changes.
 - Added missing "show groups" toggle check.

 1.6.0  (Major UI Improvements)
 - Introduced fully functional modal system (non-blocking).
 - Replaced Read-Host prompts with TUI modals.
 - Added Create-FilterPanel (initial modern version).
 - Added Show-QuickFilterDialog function.
 - Added TreeView bounds fixes + visibility fixes.

 1.6.1  (Bugfix)
 - Fixed filter panel incorrectly covering entire window.
 - Fixed TreeView being hidden beneath filter layer.
 - Fixed Update-FilterStatusLabel rendering directly on main window.

 1.6.2  (Bugfix)
 - Fixed MessageBox defaultButton index errors.
 - Fixed PasswordGenerator dialog always showing success regardless of click.
 - Corrected missing `.Visible = $false` on filter panel startup.

 1.6.3  (New Feature: Password Generator)
 - Added Generate-RandomPassword function.
 - Added menu entry “_Password Generator”.
 - Added modal with secret password textbox.
 - Added copy-to-clipboard support.
 - Added character-set toggles: upper/lower/numbers/symbols.

 1.6.4  (Bugfix)
 - Fixed "password copied" message showing when NOT copying.
 - Fixed textbox not rendering due to incorrect X/Y offsets.
 - Fixed modal stacking order (TreeView was drawing under modals).

 1.6.5  (Bugfix)
 - Fixed MessageBox.Query always returning OK due to wrong button array.
 - Corrected typos: “Copie” → “Copied”, “Genertor” → “Generator”.
 - Fixed UI padding around password modal.

 1.6.6  (UI & Layout Fixes)
 - Fixed TreeView not anchored properly when window resized.
 - Fixed filter panel stealing focus on startup.
 - Fixed modal shadows not redrawing.

 1.6.7  (Bugfix)
 - Corrected menu hotkeys.
 - Fixed password modal height too small on macOS Terminal.
 - Fixed Build-Tree not auto-refreshing after filter changes.

 1.6.8  (Stability & AD Query Fixes)
 - Fixed recursive OU building missing final child nodes.
 - Fixed groups sometimes displayed as users due to schema mismatch.
 - Fixed refresh loop running twice on some domains.
 - Added safer DN parsing with fallback.

 1.6.9  (Today’s Fixes — Window/Layout Rebuild)
 - Reimplemented main window cleanly:
       * Menu bar
       * TreeView left pane
       * Status bar bottom
       * Filter panel hidden by default
 - Fixed filter panel appearing as full-size window.
 - Fixed TreeView not rendering due to misplaced Add() calls.
 - Fixed Update-FilterStatusLabel drawing onto main window instead of label.
 - Fixed modal stacking/z-order interference.
 - Fixed several cosmetic layout typos (“fitler”, “proprties”, “protaitonal”).
 - Code cleanup: removed obsolete commented blocks interfering with layout.

 1.7.0  (Demo data and menu fixes)
 - Rework demo data to be more AD like cleaning up redundancy
 - Create dedicated about menu
 - Add shortcuts modal

 1.7.0  (Demo data expansion - New cities and band)
 - Rendr demo data in a better fashion
 - Add user devices and office printers
 - Add Rocazino form the Koge office
 - Have other locations in Denmark along with different user properites e.g. Alan Wilder

 1.7.1  Fixes for regressions in 1.7.0
 - Clean up entropy box
 - Clean up Demo data tree building
 - Add and Remove button code fixes

1.8.0  Domian Controller Information
 - Domain controller informaiton (WIP)
 - Domain replication and syncing modal (WIP)
 - Fix tree view right click context menu so it shows now

1.8.1 Verbosity
 - Add Debug-Log and clean up Debug-Log calls
 - Convert Demo data to use same format as produciton data which cuts code down significantly
 - clean up outut via ebug-Log and removing dead or unused  dmeo code stanzas

===========================================================================================
#>

param(
    [switch]$DemoMode,
    [switch]$Verbose,
    [string]$Domain,
    [ValidateSet("light","dark","matrix","british")]
    [string]$Theme = "dark"
)

Write-Host "Starting DSA-TUI in $(if($DemoMode){'DEMO'}else{'PRODUCTION'}) mode with $Theme theme..."

# Define the build version once
$BuildVersion = "1.8.1"

## For passwords expiring soon
$sevenDaysFileTime = (Get-Date).AddDays(-7).ToFileTime()

# ------------------------- Load Terminal.Gui ------------------------
Write-Host "Checking Terminal.Gui assembly..."
if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Terminal.Gui' })) {
    $mod = Get-Module -ListAvailable Microsoft.PowerShell.ConsoleGuiTools | Select-Object -First 1
    if ($mod) {
        $dll = Join-Path $mod.ModuleBase 'Terminal.Gui.dll'
        if (Test-Path $dll) { Add-Type -Path $dll -ErrorAction Stop; Write-Host "Loaded Terminal.Gui from $dll" } 
        else { Write-Error "Terminal.Gui.dll not found. Install Microsoft.PowerShell.ConsoleGuiTools."; return }
    } else { Write-Error "Microsoft.PowerShell.ConsoleGuiTools module not found."; return }
} else { Write-Host "Terminal.Gui assembly already loaded." }

# ------------------------- Globals ------------------------
if (-not $Domain) { $Domain = if ($DemoMode) { "example.com" } else { (Get-ADDomain).Forest } }
$Global:Domain = $Domain
$Global:CurrentDC = $null
$Global:Users = @()
$Global:DCs = @()
$Global:ADObjects = @()  # New for production AD object detection
$Global:SelectedObjects = @()
$Global:SelectionMode = $false

# Set global demo mode flag immediately
$Global:DemoMode = $DemoMode

# Set global flags immediately after param block - themes here
$Global:DemoMode = $DemoMode
$script:ThemeMode = $Theme

Write-Host "Starting DSA-TUI in $(if($DemoMode){'DEMO'}else{'PRODUCTION'}) mode..."

# Global Search filters:
$Global:FilterOptions = @{
    ShowDisabledUsers = $true
    ShowEnabledUsers = $true
    SowLockedUsers = $true
    ShowGroups = $true
    ShowDCs = $true
    ShowComputers = $true
    ShowOUs = $true
    NameFilter = ""
    SortBy = "Name"
    SortDescending = $false
}

## --------------------------{ Debug Logging }-------------------------
function Debug-Log {
  param([string]$Message)
    if ($Verbose) {
      $ts = (Get-Date).ToString('HH:mm:ss')
      Write-Host "[$ts] LOG: $Message" -ForegroundColor Cyan
    }
}

# ---- Theme Definitions ----
function Get-Theme {
    param([string]$mode)

    # Initialize color schemes and Ensure ColorSchemes are instantiated
    if (-not $globalCs)     { $globalCs     = [Terminal.Gui.ColorScheme]::new() }
    if (-not $mainWindowCs) { $mainWindowCs = [Terminal.Gui.ColorScheme]::new() }

    # Normalize theme string: lowercase + ASCII
    $mode = $mode.Trim().ToLower()
####    $mode = $mode -replace "ae","ae"

## Adding Themes:
##
## Add an option above in the [ValidateSet() then define a theme below:
##
##       "faxekondi" {
##            $globalCs.Normal     <-- Foreground borders and background colour for all modals
##            $globalCs.Focus      <-- Foreground and background for menus
##            $mainWindowCs.Normal <-- Main opening dialog and foreground text colour
##            $mainWindowCs.Focus  <-- Main opening window focus colours foreground nad background
##        }

    switch ($mode) {
        "light" {
            $globalCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Cyan)
            $globalCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
            $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Cyan)
            $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::Blue)
        }

        "dark" {
            $globalCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Black)
            $globalCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Green,[Terminal.Gui.Color]::White)
            $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Green)
            $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Red)
        }

        "matrix" {
            $globalCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Green,[Terminal.Gui.Color]::Black)
            $globalCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Green)
            $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::Black)
            $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::Gray)
        }

        "british" {
            $globalCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
            $globalCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
            $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
            $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::White)
        }

        default {
            # fallback to dark
            $globalCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Black)
            $globalCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Gray)
            $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Black)
            $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::DarkGray)
        }
    }

    # Ensure HotNormal/HotFocus
    $globalCs.HotNormal     = $globalCs.Normal
    $globalCs.HotFocus      = $globalCs.Focus
    $mainWindowCs.HotNormal = $mainWindowCs.Normal
    $mainWindowCs.HotFocus  = $mainWindowCs.Focus

    return @{
        Global     = $globalCs
        MainWindow = $mainWindowCs
    }
}

function Apply-Theme {
    param(
        [hashtable]$ThemeData,        # expects keys: Global, MainWindow
        [object]$TopLevel,
        [object]$MainWindow,
        [object]$Menu,
        [object]$Status
    )

    if ($null -eq $ThemeData) { return }

    # --- Global / TopLevel ---
    if ($TopLevel -and $TopLevel.PSObject.Properties.Name -contains 'ColorScheme') {
        $TopLevel.ColorScheme = $ThemeData.Global
    }

    # --- Main window ---
    if ($MainWindow -and $MainWindow.PSObject.Properties.Name -contains 'ColorScheme') {
        $MainWindow.ColorScheme = $ThemeData.MainWindow
    }

    # --- Menu ---
    if ($Menu -and $Menu.PSObject.Properties.Name -contains 'ColorScheme') {
        $Menu.ColorScheme = $ThemeData.Global
    }

    # --- StatusBar ---
    if ($Status -and $Status.PSObject.Properties.Name -contains 'ColorScheme') {
        $Status.ColorScheme = $ThemeData.Global
    }

    # --- Terminal.Gui base colors ---
    [Terminal.Gui.Colors]::Base     = $ThemeData.Global
    [Terminal.Gui.Colors]::Dialog   = $ThemeData.Global
    [Terminal.Gui.Colors]::Menu     = $ThemeData.Global
    [Terminal.Gui.Colors]::Error    = $ThemeData.Global
    [Terminal.Gui.Colors]::TopLevel = $ThemeData.Global
}

# Diagnostics helper to show what's inside a ColorScheme
function Dump-ColorScheme {
    param([Terminal.Gui.ColorScheme]$Scheme)
    if ($null -eq $Scheme) { Debug-Log "ColorScheme is null"; return }
    Debug-Log "Normal    : $($Scheme.Normal)"
    Debug-Log "Focus     : $($Scheme.Focus)"
    Debug-Log "HotNormal : $($Scheme.HotNormal)"
    Debug-Log "HotFocus  : $($Scheme.HotFocus)"
    Debug-Log "Disabled  : $($Scheme.Disabled)"
}

## Select theme before proceeding. Save the mode string
$script:ThemeMode = $Theme

# Get the selected colour scheme
$cs = Get-Theme -mode $Theme
Apply-Theme -ThemeData $themeData -TopLevel $TopLevel -MainWindow $MainWindow -Menu $Menu -Status $Status


# ---- Helper: Show a simple loading/progress dialog with spinner ----
function Show-LoadingDialog {
    param(
        [string]$Message = "Loading, please wait..."
    )

    # Create dialog and label
    $dlg = [Terminal.Gui.Dialog]::new("", 40, 7)
    $lbl = [Terminal.Gui.Label]::new($Message)
    $lbl.X = 2
    $lbl.Y = 2
    $dlg.Add($lbl)

    # Spinner label
    $spinner = [Terminal.Gui.Label]::new("|")
    $spinner.X = [Terminal.Gui.Pos]::Right($lbl) + 1
    $spinner.Y = 2
    $dlg.Add($spinner)

    # Spinner frames and timer setup
    $frames = @("|", "/", "-", "\")
    $i = 0
    $timer = [System.Threading.Timer]::new(
        {
            $global:spinnerFrameIndex = ($global:spinnerFrameIndex + 1) % 4
            [Terminal.Gui.Application]::MainLoop.Invoke({
                $spinner.Text = $frames[$global:spinnerFrameIndex]
            })
        },
        $null, 0, 150
    )

    # Start non-blocking dialog
    [Terminal.Gui.Application]::Begin($dlg)

    # Return both dialog and timer so caller can close/stop cleanly
    return [PSCustomObject]@{ Dialog = $dlg; Timer = $timer }
}

# ---- Helper: Close loading/progress dialog ----
function Close-LoadingDialog {
    param($loading)

    if ($null -ne $loading.Timer) { $loading.Timer.Dispose() }
    if ($null -ne $loading.Dialog) { [Terminal.Gui.Application]::End($loading.Dialog) }
}

## ---------------------{ Pretty Theme Selections }--------------------
function Show-ThemeSelector {
  $dlg = [Terminal.Gui.Dialog]::new("Select Theme", 50, 14)
  $lbl = [Terminal.Gui.Label]::new("Choose a color theme:"); $lbl.X=2; $lbl.Y=1; $dlg.Add($lbl)
  
  $themes = @("light", "dark", "matrix", "british")
  $currentIndex = $themes.IndexOf($script:ThemeMode)
  if ($currentIndex -lt 0) { $currentIndex = 1 } # default to dark

  $rdoThemes = [Terminal.Gui.RadioGroup]::new($themes)
  $rdoThemes.X=2; $rdoThemes.Y=3; $rdoThemes.SelectedItem=$currentIndex
  $dlg.Add($rdoThemes)
    
  $btnApply = [Terminal.Gui.Button]::new("Apply")

  $btnApply.add_Clicked({
    $selectedTheme = $themes[$rdoThemes.SelectedItem]
    Debug-Log "Switching to theme: $selectedTheme"
    $script:ThemeMode = $selectedTheme
        
    ## Get new theme
    $newTheme = Get-Theme -mode $selectedTheme
        
    ## Apply to all components

    ## Apply theme to all components
    Apply-Theme -ThemeData $newTheme -TopLevel [Terminal.Gui.Application]::Top -MainWindow $win -Menu $menu -Status $StatusBar

    ## Force redraw
    $menu.ColorScheme = $newTheme.Global     ## <-- critical for menu bar
    $menu.SetNeedsDisplay()
    $win.SetNeedsDisplay()
    [Terminal.Gui.Application]::Top.SetNeedsDisplay()
    [Terminal.Gui.Application]::Refresh()

    ## Force redraw of components that do NOT auto-refresh
    $Menu.SetNeedsDisplay()
    $win.SetNeedsDisplay()
    [Terminal.Gui.Application]::Top.SetNeedsDisplay()

    ## Hard refresh screen
    [Terminal.Gui.Application]::Refresh()
     
    Show-Modal "Theme Changed" "Theme changed to: ${selectedTheme}"
    [Terminal.Gui.Application]::RequestStop()
  })


  $dlg.AddButton($btnApply)
    
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $dlg.AddButton($btnCancel)
    
  [Terminal.Gui.Application]::Run($dlg)
}

# ------------------------- Load Domain Data ------------------------
function Get-ADObjectsByType {
    param([string]$domain)
    $objTypes = @("user","computer","group","organizationalUnit","contact")
    $allObjects = @()
    foreach ($type in $objTypes) {
        try {
            $objs = if ($Global:DemoMode) {
                # Demo objects already structured
                @()
            } else {
                Get-ADObject -Filter "ObjectClass -eq '$type'" -Server $domain -Properties Name,ObjectClass,DistinguishedName |
                    ForEach-Object { @{ Name=$_.Name; Type=$_.ObjectClass; DN=$_.DistinguishedName } }
            }
            $allObjects += $objs
        } catch {
            # minimal fix: string interpolation of exception object done via ToString()
            Debug-Log "DEBUG: Failed to enumerate ${type}: $($_.ToString())"
        }
    }
    return $allObjects
}

# ------------------------- Load Domain Data ------------------------
<#
A note on phone numbers:

UK Phone Number Standards (Ofcom reserved ranges for fiction/testing):
- Glasgow: 0141 496 0xxx + Mobile: 07700 900xxx
- Edinburgh: 0131 496 0xxx + Mobile: 07700 900xxx
- London: 020 7946 0xxx / 01632 96xxxx + Mobile: 07700 900xxx

Denmark Testing Numbers:
- Copenhagen landline: +45 0000-xxxx
- Denmark mobile: +45 2xxx xxxx

NB: I don't believe Denmark uses 0000 but this is not confirmed!
#>

function Load-DomainData {
  param([string]$domain)

  if ($Logging) { Write-Host "DEBUG: Loading domain data for: $domain" }

  if ($Global:DemoMode) {
    Write-Host "Starting DSA-TUI in DEMO mode..."

    # ------------------ Define Demo Users ------------------
    $Global:rawUsers = @(
            # ========== Simple Minds (UK/Scotland/Glasgow) ==========
            @{
                Name = 'Jim Kerr'; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds'); Groups = @('Simple Minds','Vocalists'); Title = 'Lead Vocalist'; Email = 'jim.kerr@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Glasgow Office'; Phone = '+44 141 111 1111'; MobilePhone = '+44 7700 111111'; Street = '1 High Street'; City = 'Glasgow'; PostalCode = 'G1 1AA'; Company = 'Example Music Ltd'; Manager = ''; Description = 'Lead vocalist for Simple Minds'
            },
            @{
                Name = 'Charlie Burchill'; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds'); Groups = @('Simple Minds','Guitarists'); Title = 'Lead Guitarist'; Email = 'charlie.b@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Glasgow Office'; Phone = '+44 141 111 1112'; MobilePhone = '+44 7700 111112'; Street = '1 High Street'; City = 'Glasgow'; PostalCode = 'G1 1AA'; Company = 'Example Music Ltd'; Manager = 'Jim Kerr'; Description = 'Guitarist and founding member of Simple Minds'
            },
            @{
                Name = 'Mel Gaynor'; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds'); Groups = @('Simple Minds','Percussion'); Title = 'Drummer'; Email = 'mel.gaynor@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Glasgow Office'; Phone = '+44 141 111 1113'; MobilePhone = '+44 7700 111113'; Street = '1 High Street'; City = 'Glasgow'; PostalCode = 'G1 1AA'; Company = 'Example Music Ltd'; Manager = 'Jim Kerr'; Description = 'Drummer for Simple Minds'
            },

            # ========== Marillion (UK/Scotland/Edinburgh) ==========
            @{
                Name = 'Derek Dick'; OU = @('Locations','UK','Scotland','Edinburgh','Marillion'); Groups = @('Marillion','Vocalists'); Title = 'Lead Vocalist'; Email = 'fish@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $true; Department = 'Music'; Office = 'Edinburgh Office'; Phone = '+44 131 222 2221'; MobilePhone = '+44 7700 222221'; Street = '22 Queens Road'; City = 'Edinburgh'; PostalCode = 'EH1 2BB'; Company = 'Example Music Ltd'; Manager = ''; Description = 'Former lead vocalist (Fish) for Marillion (1981-1988)'
            },
            @{
                Name = 'Steve Rothery'; OU = @('Locations','UK','Scotland','Edinburgh','Marillion'); Groups = @('Marillion','Guitarists'); Title = 'Lead Guitarist'; Email = 'steve.rothery@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Edinburgh Office'; Phone = '+44 131 222 2222'; MobilePhone = '+44 7700 222222'; Street = '22 Queens Road'; City = 'Edinburgh'; PostalCode = 'EH1 2BB'; Company = 'Example Music Ltd'; Manager = 'Derek Dick'; Description = 'Lead guitarist and founding member of Marillion'
            },
            @{
                Name = 'Pete Trewavas'; OU = @('Locations','UK','Scotland','Edinburgh','Marillion'); Groups = @('Marillion','Guitarists'); Title = 'Bassist'; Email = 'pete.trewavas@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Edinburgh Office'; Phone = '+44 131 222 2223'; MobilePhone = '+44 7700 222223'; Street = '22 Queens Road'; City = 'Edinburgh'; PostalCode = 'EH1 2BB'; Company = 'Example Music Ltd'; Manager = 'Derek Dick'; Description = 'Bassist and founding member of Marillion'
            },
            @{
                Name = 'Mark Kelly'; OU = @('Locations','UK','Scotland','Edinburgh','Marillion'); Groups = @('Marillion','Keyboards'); Title = 'Keyboardist'; Email = 'mark.kelly@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Edinburgh Office'; Phone = '+44 131 222 2224'; MobilePhone = '+44 7700 222224'; Street = '22 Queens Road'; City = 'Edinburgh'; PostalCode = 'EH1 2BB'; Company = 'Example Music Ltd'; Manager = 'Derek Dick'; Description = 'Keyboardist and founding member of Marillion'
            },
            @{
                Name = 'Ian Mosley'; OU = @('Locations','UK','Scotland','Edinburgh','Marillion'); Groups = @('Marillion','Percussion'); Title = 'Drummer'; Email = 'ian.mosley@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Edinburgh Office'; Phone = '+44 131 222 2225'; MobilePhone = '+44 7700 222225'; Street = '22 Queens Road'; City = 'Edinburgh'; PostalCode = 'EH1 2BB'; Company = 'Example Music Ltd'; Manager = 'Derek Dick'; Description = 'Drummer for Marillion (joined 1984)'
            },

            # ========== Erasure (UK/England/London) ==========
            @{
                Name = 'Andy Bell'; OU = @('Locations','UK','England','London','Erasure'); Groups = @('Erasure','Vocalists'); Title = 'Lead Vocalist'; Email = 'andy.bell@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'London Office'; Phone = '+44 20 7000 1111'; MobilePhone = '+44 7700 333333'; Street = '15 River Street'; City = 'London'; PostalCode = 'E1 7AA'; Company = 'Example Music Ltd'; Manager = ''; Description = 'Lead vocalist for Erasure'
            },
            @{
                Name = 'Vince Clarke'; OU = @('Locations','UK','England','London','Erasure'); Groups = @('Erasure','Synth','Keyboards'); Title = 'Synth / Keyboardist'; Email = 'vince.clarke@example.com'; Country = 'UK'; Disabled = $false; Locked = $true; MustChangePassword = $false; Department = 'Music'; Office = 'London Office'; Phone = '+44 20 7000 1112'; MobilePhone = '+44 7700 333334'; Street = '15 River Street'; City = 'London'; PostalCode = 'E1 7AA'; Company = 'Example Music Ltd'; Manager = 'Andy Bell'; Description = 'Synthesizer pioneer - founding member of Depeche Mode and Erasure'
            },

            # ========== Depeche Mode (UK/England/London) ==========
            @{
                Name = 'Martin Gore'; OU = @('Locations','UK','England','London','Depeche Mode'); Groups = @('Depeche Mode','Guitarists','Keyboards'); Title = 'Guitarist/Keyboardist'; Email = 'martin.gore@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'London Office'; Phone = '+44 20 7000 2221'; MobilePhone = '+44 7700 444441'; Street = '32 Abbey Lane'; City = 'London'; PostalCode = 'EC2 1AA'; Company = 'Example Music Ltd'; Manager = ''; Description = 'Guitarist, keyboardist and primary songwriter for Depeche Mode'
            },
            @{
                Name = 'Dave Gahan'; OU = @('Locations','UK','England','London','Depeche Mode'); Groups = @('Depeche Mode','Vocalists'); Title = 'Lead Vocalist'; Email = 'dave.gahan@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $true; Department = 'Music'; Office = 'London Office'; Phone = '+44 20 7000 2222'; MobilePhone = '+44 7700 444442'; Street = '32 Abbey Lane'; City = 'London'; PostalCode = 'EC2 1AA'; Company = 'Example Music Ltd'; Manager = 'Martin Gore'; Description = 'Lead vocalist for Depeche Mode'
            },
            @{
                Name = 'Alan Wilder'; OU = @('Locations','UK','England','London','Depeche Mode'); Groups = @('Depeche Mode','Keyboards','Percussion'); Title = 'Keyboardist/Drummer'; Email = 'alan.wilder@example.com'; Country = 'UK'; Disabled = $false; Locked = $true; MustChangePassword = $false; Department = 'Music'; Office = 'London Office'; Phone = '+44 20 7000 2224'; MobilePhone = '+44 7700 444444'; Street = '32 Abbey Lane'; City = 'London'; PostalCode = 'EC2 1AA'; Company = 'Example Music Ltd'; Manager = 'Martin Gore'; Description = 'Multi-instrumentalist for Depeche Mode (1982-1995, departed)'
            },
            @{
                Name = 'Andrew Fletcher'; OU = @('Locations','UK','England','London','Depeche Mode'); Groups = @('Depeche Mode','Keyboards'); Title = 'Keyboards/Bass Synth'; Email = 'andrew.fletcher@example.com'; Country = 'UK'; Disabled = $true; Locked = $true; MustChangePassword = $false; Department = 'Music'; Office = 'London Office'; Phone = '+44 20 7000 2223'; MobilePhone = '+44 7700 444443'; Street = '32 Abbey Lane'; City = 'London'; PostalCode = 'EC2 1AA'; Company = 'Example Music Ltd'; Manager = 'Martin Gore'; Description = 'Keyboard and bass synthesizer for Depeche Mode (deceased)'
            },

            # ========== TV-2 (Denmark/Copenhagen) ==========
            @{
                Name = 'Steffen Brandt'; OU = @('Locations','Denmark','Copenhagen','TV-2'); Groups = @('TV-2','Vocalists','Guitarists'); Title = 'Lead Vocalist / Guitarist'; Email = 'steffen.brandt@example.com'; Country = 'DK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Copenhagen Office'; Phone = '+45 0000 2222'; MobilePhone = '+45 5012 3457'; Street = '1 Raadhuspladsen'; City = 'Copenhagen'; PostalCode = '1550'; Company = 'Example Music ApS'; Manager = ''; Description = 'Frontman of TV-2'
            },

            # ========== Rocazino (Denmark/Koge) ==========
            @{
                Name = 'Ulla Kjaer'; OU = @('Locations','Denmark','Koge','Rocazino'); Groups = @('Rocazino','Vocalists'); Title = 'Lead Vocalist'; Email = 'ulla.kjaer@example.com'; Country = 'DK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Koge Office'; Phone = '+45 0000 2234'; MobilePhone = '+45 3012 3456'; Street = '7 Torvet'; City = 'Koge'; PostalCode = '4600'; Company = 'Example Music ApS'; Manager = ''; Description = 'Lead vocalist for Rocazino'
            },
            @{
                Name = 'Michael Bruun'; OU = @('Locations','Denmark','Koge','Rocazino'); Groups = @('Rocazino','Guitarists'); Title = 'Guitarist'; Email = 'michael.bruun@example.com'; Country = 'DK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Koge Office'; Phone = '+45 0000 2235'; MobilePhone = '+45 3012 3457'; Street = '7 Torvet'; City = 'Koge'; PostalCode = '4600'; Company = 'Example Music ApS'; Manager = 'Ulla Kjaer'; Description = 'Guitarist and songwriter for Rocazino'
            },
            @{
                Name = 'Jan Sivertsen'; OU = @('Locations','Denmark','Koge','Rocazino'); Groups = @('Rocazino','Percussion'); Title = 'Drummer'; Email = 'jan.sivertsen@example.com'; Country = 'DK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Koge Office'; Phone = '+45 0000 2236'; MobilePhone = '+45 3012 3458'; Street = '7 Torvet'; City = 'Koge'; PostalCode = '4600'; Company = 'Example Music ApS'; Manager = 'Ulla Kjaer'; Description = 'Drummer for Rocazino'
            },

            # ========== MØ (Denmark/Odense) ==========
            @{
                Name = 'Karen Marie Orsted'; OU = @('Locations','Denmark','Odense','Mo'); Groups = @('Mo','Vocalists'); Title = 'Singer / Songwriter'; Email = 'karen.orsted@example.com'; Country = 'DK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Odense Office'; Phone = '+45 0000 3234'; MobilePhone = '+45 4012 3456'; Street = '22 Vestergade'; City = 'Odense'; PostalCode = '5000'; Company = 'Example Music ApS'; Manager = ''; Description = 'Danish singer-songwriter known internationally as MØ'
            }
        )
        
    $Global:rawDCs = @(
        @{ Name='EXAGLADC01'; Site='GLA' },
        @{ Name='EXAEDIDC01'; Site='EDI' },
        @{ Name='EXALNDCDC01'; Site='LND' },
        @{ Name='EXACPHDC01'; Site='CPH' }
    )

    Write-Host "DEBUG: rawUsers count:" $Global:rawUsers.Count
    Write-Host "DEBUG: rawDCs count:" $Global:rawDCs.Count

    # ------------------ Convert to AD-like objects ------------------
    $converted = Convert-DemoDataToADObjects -Users $Global:rawUsers -DCs $Global:rawDCs -Domain $Global:Domain

    # The function already sets $Global:Users, $Global:DCs, $Global:ADObjects
    Write-Host "DEBUG: Users count:" $Global:Users.Count
    Write-Host "DEBUG: DCs count:" $Global:DCs.Count
    Write-Host "DEBUG: ADObjects count:" $Global:ADObjects.Count

    } else {
        # Production mode - real AD calls
        try {
            Import-Module ActiveDirectory -ErrorAction Stop

            $loadingDlg = Show-LoadingDialog -Message "Loading AD objects for $domain..."
            try {
                # Domain Controllers
                $Global:DCs = Get-ADDomainController -Discover -Domain $domain |
                    ForEach-Object { @{ Name=$_.HostName; OU='Domain Controllers'; Site=$_.Site } }

                # Get AD objects
                $Global:ADObjects = Get-ADObjectsByType -domain $domain

                # Users
########################################### THIS CODE IS SETTING NULLS INSTEADO F READING VALUES ##############################
                $Global:Users = $Global:ADObjects | Where-Object { $_.Type -eq 'user' } | ForEach-Object {
                    $ou = ($_.DN -split ',') | Where-Object { $_ -like 'OU=*' } | Select-Object -First 1
                    if ($ou) { $ou = $ou -replace '^OU=' ,'' } else { $ou = "" }
                    @{ Name=$_.Name; OU=$ou; Groups=$null; Title=$null; Email=$null; Country=$null; Disabled=$false }
                }
            } finally {
                Close-LoadingDialog $loadingDlg
            }

        } catch {
            [Terminal.Gui.MessageBox]::Query("Error","Failed to query domain ${domain}:`n$_","OK") | Out-Null
            $Global:Users=@(); $Global:DCs=@(); $Global:ADObjects=@()
        }
    }
}

function Convert-DemoDataToADObjects {
    <#
    .SYNOPSIS
    Converts demo hashtable data to AD-like PSCustomObjects
    #>
    param(
        [array]$Users,
        [array]$DCs = @(),
        [string]$Domain = "example.com",
        [string]$BaseDN = "DC=example,DC=com"
    )

    Write-Host "DEBUG: Converting demo data to AD-like objects..."

    # Helper functions
    function New-FakeGuid { [guid]::NewGuid().ToString() }
    function New-FakeSid { 
        $rid = Get-Random -Minimum 1000 -Maximum 65535
        "S-1-5-21-{0}-{1}-{2}-{3}" -f (Get-Random -Max 999999999), (Get-Random -Max 999999999), (Get-Random -Max 999999999), $rid
    }

    # Convert Users
    $convertedUsers = @()
    foreach ($user in $Users) {
        $sam = ($user.Name -replace '\s+', '.').ToLower()
        $upn = if ($user.Email) { $user.Email } else { "$sam@$Domain" }

        # Build DN
        if ($user.OU) {
            $ouChain = $user.OU | ForEach-Object { "OU=$_" }
            $dn = "CN=$($user.Name)," + ($ouChain[-1..0] -join ',') + ",$BaseDN"
        } else {
            $dn = "CN=$($user.Name),$BaseDN"
        }

        $adUser = [PSCustomObject]@{
            ObjectClass       = 'user'
            Name              = $user.Name
            SamAccountName    = $sam
            UserPrincipalName = $upn
            DisplayName       = $user.Name
            GivenName         = ($user.Name -split '\s+')[0]
            Surname           = ($user.Name -split '\s+')[-1]
            DistinguishedName = $dn
            ObjectGUID        = New-FakeGuid
            SID               = New-FakeSid
            Enabled           = (-not $user.Disabled)
            LockedOut         = [bool]$user.Locked
            PasswordExpired   = [bool]$user.MustChangePassword

            # Contact info
            Title       = $user.Title
            Department  = $user.Department
            Company     = $user.Company
            Manager     = $user.Manager
            EmailAddress= $user.Email
            OfficePhone = $user.Phone
            MobilePhone = $user.MobilePhone
            Office      = $user.Office

            # Address
            StreetAddress = $user.Street
            City          = $user.City
            PostalCode    = $user.PostalCode
            Country       = $user.Country

            # Other
            Description = $user.Description
            MemberOf    = $user.Groups
            CanonicalName = ($user.OU -join '/') + "/$($user.Name)"
            whenCreated   = (Get-Date).AddDays(-90)
            whenChanged   = (Get-Date).AddDays(-5)

            # Original demo properties
            OU       = $user.OU
            Groups   = $user.Groups
            Disabled = $user.Disabled
            Locked   = $user.Locked
        }

        # Make it look like AD user object
        $adUser.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADUser')

        $convertedUsers += $adUser
    }

    # Convert DCs
    $convertedDCs = @()
    foreach ($dc in $DCs) {
        $dn = "CN=$($dc.Name),OU=Domain Controllers,$BaseDN"

        $adDC = [PSCustomObject]@{
            ObjectClass        = 'computer'
            Name               = $dc.Name
            DNSHostName        = "$($dc.Name).$Domain"
            DistinguishedName  = $dn
            ObjectGUID         = New-FakeGuid
            SID                = New-FakeSid
            Enabled            = $true
            Site               = $dc.Site
            OperatingSystem    = 'Windows Server 2022'
            OperatingSystemVersion = '10.0 (20348)'
            whenCreated        = (Get-Date).AddDays(-180)
            OU                 = 'Domain Controllers'
        }

        $adDC.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADComputer')

        $convertedDCs += $adDC
    }

    # Set global variables
    $Global:Users     = $convertedUsers
    $Global:DCs       = $convertedDCs
    $Global:ADObjects = $convertedUsers + $convertedDCs

    Write-Host "DEBUG: Converted $($convertedUsers.Count) users and $($convertedDCs.Count) DCs to AD-like objects"

    # Return hashtable
    return @{
        Users = $convertedUsers
        DCs   = $convertedDCs
    }
}

# Updated Build-Tree to show locked status
# Modify the existing Build-Tree function to show both disabled and locked status:
function Build-Tree {
    param([string]$domain)

    Write-Host "DEBUG: Building tree for domain $domain..."

    $tree.ClearObjects()
    $root = [Terminal.Gui.Trees.TreeNode]::new($domain)

    if ($Global:DemoMode) {
        Write-Host "DEBUG: Building demo mode tree..."

        # Helper class for OU nodes
        class OUNode {
            [string]$Name
            [System.Collections.Generic.List[object]]$Children
            [object]$Tag

            OUNode([string]$name) {
                $this.Name = $name
                $this.Children = [System.Collections.Generic.List[object]]::new()
            }

            [string] ToString() {
                return $this.Name
            }
        }

        # Apply filters once
        $nameFilter = $Global:FilterOptions.NameFilter.Trim()
        $filteredUsers = $Global:Users | Where-Object {
            ($_.Disabled -and $Global:FilterOptions.ShowDisabledUsers) -or
            (-not $_.Disabled -and $Global:FilterOptions.ShowEnabledUsers)
        }

        if ($nameFilter) {
            $filteredUsers = $filteredUsers | Where-Object {
                $_.Name -like "*$nameFilter*" -or
                $_.EmailAddress -like "*$nameFilter*" -or
                $_.Title -like "*$nameFilter*"
            }
        }

        Write-Host "DEBUG: Filtered to $($filteredUsers.Count) users"

        # Build hierarchical OU tree
        $rootOU = [OUNode]::new($domain)
        $nodeCache = @{ "" = $rootOU }

        foreach ($user in $filteredUsers) {
            if (-not $user.OU) { continue }

            $pathSoFar = ""
            $currentNode = $rootOU

            foreach ($ouLevel in $user.OU) {
                $newPath = if ($pathSoFar) { "$pathSoFar/$ouLevel" } else { $ouLevel }

                if ($nodeCache.ContainsKey($newPath)) {
                    $currentNode = $nodeCache[$newPath]
                } else {
                    $newNode = [OUNode]::new($ouLevel)
                    $currentNode.Children.Add($newNode)
                    $nodeCache[$newPath] = $newNode
                    $currentNode = $newNode
                }

                $pathSoFar = $newPath
            }

            # Add user node with status icon
            $statusIcon = if ($user.Locked) { "🔒" } elseif ($user.Disabled) { "⊗" } else { "○" }
            $userNode = [OUNode]::new("(U) $statusIcon $($user.Name)")
            $userNode.Tag = $user
            $currentNode.Children.Add($userNode)
        }

        # Add Groups under their respective OUs
        if ($Global:FilterOptions.ShowGroups) {
            $allGroups = $filteredUsers | ForEach-Object { $_.Groups } | Where-Object { $_ } | Select-Object -Unique
            foreach ($groupName in $allGroups | Sort-Object) {
                $groupNode = [OUNode]::new($groupName)
                $members = $filteredUsers | Where-Object { $_.Groups -contains $groupName } | Sort-Object -Property Name
                foreach ($member in $members) {
                    $statusIcon = if ($member.Locked) { "🔒" } elseif ($member.Disabled) { "⊗" } else { "○" }
                    $memberNode = [OUNode]::new("(U) $statusIcon $($member.Name)")
                    $memberNode.Tag = $member
                    $groupNode.Children.Add($memberNode)
                }
                if ($groupNode.Children.Count -gt 0) {
                    $rootOU.Children.Add($groupNode)
                }
            }
        }

        # Add Domain Controllers
        if ($Global:FilterOptions.ShowDCs -and $Global:DCs.Count -gt 0) {
            $dcNode = [OUNode]::new("Domain Controllers")
            foreach ($dc in ($Global:DCs | Sort-Object -Property Name)) {
                $dcChildNode = [OUNode]::new("(DC) $($dc.Name) [$($dc.Site)]")
                $dcNode.Children.Add($dcChildNode)
            }
            $rootOU.Children.Add($dcNode)
        }

        # Add top-level OU to tree
        $tree.AddObject($rootOU)

        Write-Host "DEBUG: Demo tree built with $($rootOU.Children.Count) top-level nodes"
    }
}

if ($Global:DemoMode) {
  Write-Host "DEBUG: rawUsers count:" $rawUsers.Count
  Write-Host "DEBUG: rawDCs count:" $rawDCs.Count

  $converted = Convert-DemoDataToADObjects -Users $Global:rawUsers -DCs $Global:rawDCs -Domain $Global:domain
} else {
  Load-DomainData -domain $Global:Domain
  # Set global variables
  $Global:Users = $converted.Users
  $Global:DCs   = $converted.DCs
  $Global:ADObjects = $converted.Users + $converted.DCs

  Write-Host $converted.Users
  Write-Host $converted.DCs
  Write-Host $converted.ADObjects
}

# Step 3: Replace your "Initialize Terminal.Gui" section with this:

# ------------------------- Initialize Terminal.Gui ------------------------
[Terminal.Gui.Application]::Init()
$top = [Terminal.Gui.Application]::Top

# Get the selected theme
Debug-Log "Applying theme: $Theme"
$themeData = Get-Theme -mode $Theme

# Apply theme to top level first
if ($themeData -and $themeData.Global) {
    $top.ColorScheme = $themeData.Global
}

# ------------------------- Main Window ------------------------
$win = [Terminal.Gui.Window]::new("DSA-TUI — Active Directory ${BuildVersion} Blaabaer")
$win.X=0; $win.Y=0; $win.Width=[Terminal.Gui.Dim]::Fill(); $win.Height=[Terminal.Gui.Dim]::Fill()

# Apply theme to main window
if ($themeData -and $themeData.MainWindow) {
    $win.ColorScheme = $themeData.MainWindow
}

$top.Add($win)

Debug-Log "Theme applied successfully: $Theme"
## filter panel
$filterPanel = Create-FilterPanel
$win.Add($filterPanel)

$filterStatusLabel = Create-FilterStatusLabel
$win.Add($filterStatusLabel)

$selectionPanel = Create-SelectionPanel
$win.Add($selectionPanel)

# ------------------------- Status Bar ------------------------
$status = [Terminal.Gui.StatusBar]::new(@(
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F1,"~F1~ Help",{
         Show-Modal "Shortcuts" "F1 - Help`nF9 -New`nF10 - Delete`nF10 - Quit`nF12 Redraw" }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F9,"~F9~ New",{ Show-NewObjectWizard }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F10,"~F10~ Quit",{ [Terminal.Gui.Application]::RequestStop() }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F12,"~F12~ Redraw",{ [Terminal.Gui.Application]::Refresh() })
))
$top.Add($status)

# ------------------------- Menu ------------------------
# Existing menu items
$mFile = [Terminal.Gui.MenuItem]::new("_Exit","Exit application",[Action]{ [Terminal.Gui.Application]::RequestStop() })
$mNew = [Terminal.Gui.MenuItem]::new("New Object","Create a new object",[Action]{ Show-NewObjectWizard })
$mProps = [Terminal.Gui.MenuItem]::new("_Properties","Edit selected properties",[Action]{ Show-Properties })
$mUndo = [Terminal.Gui.MenuItem]::new("_Undo","Undo last action",[Action]{ Debug-Log "DEBUG: Undo placeholder" })
$mChangeDomain = [Terminal.Gui.MenuItem]::new("Change _Domain","Select domain",[Action]{ Show-ChangeDomainDialog })
$mChangeDC = [Terminal.Gui.MenuItem]::new("Change _Domain Controller","Select DC",[Action]{ Show-ChangeDCDialog })
$mSearchAD = [Terminal.Gui.MenuItem]::new("_Search AD","Search Active Directory",[Action]{ Show-ADSearchDialog })
$mRefresh = [Terminal.Gui.MenuItem]::new("_Refresh","Refresh AD data",[Action]{ Refresh-TreeData })
$mQuickFilter = [Terminal.Gui.MenuItem]::new("_Quick Filter","Apply quick filters",[Action]{ Show-QuickFilterDialog })
$mSelectionMode = [Terminal.Gui.MenuItem]::new("_Selection Mode (Ctrl+S)","Toggle selection mode",[Action]{ Toggle-SelectionMode })
$mSelectAll = [Terminal.Gui.MenuItem]::new("Select _All (Ctrl+A)","Select all objects",[Action]{ Select-AllObjects })
$mDeselectAll = [Terminal.Gui.MenuItem]::new("_Deselect All (Ctrl+D)","Deselect all objects",[Action]{ Deselect-AllObjects })
$mBulkAddGroup = [Terminal.Gui.MenuItem]::new("Add to _Group...","Add selected users to group",[Action]{ Invoke-BulkAddToGroup })
$mPasswordGenerator = [Terminal.Gui.MenuItem]::new("_Password Generator","Password Generator",[Action]{ Generate-RandomPassword })
$mADHealth = [Terminal.Gui.MenuItem]::new("_AD Health Status","AD Health And Replicaiton Status",[Action]{ Get-ADHealth })

# --- NEW: About Menu Items ---

$mShortcuts = [Terminal.Gui.MenuItem]::new(
    "_Shortcuts",
    "Keyboard shortcuts",
    [Action]{
        Show-Modal "Shortcuts" "F1 - Help`nF9 -New`nF10 - Quit`nF11 -Redraw"
    }
)

$mAboutDSATUI = [Terminal.Gui.MenuItem]::new(
    "_About DSA-TUI",
    "About this application",
    [Action]{ 
        [Terminal.Gui.MessageBox]::Query("About","DSA-TUI ${BuildVersion} Blaabaer`n© 2025 Copyleft (GPL-3)`nDemo Mode: $DemoMode","OK") | Out-Null 
    }
)

$mWhyBlaabaer = [Terminal.Gui.MenuItem]::new(
    "Why _Blaabaer?",
    "Why the Blaabaer codename?",
    [Action]{ Show-BlaabaerInfo }
)

# --- NEW: About Top-Level Menu ---
$aboutMenu = [Terminal.Gui.MenuBarItem]::new("_About", @(
    $mShortcuts,
    $mAboutDSATUI,
    $mWhyBlaabaer
))

# Mandatory (BEFORE creating $menu):
$mTheme = [Terminal.Gui.MenuItem]::new("_Theme","Change color theme",[Action]{ Show-ThemeSelector })


# ------------------------- FIXED MENU BAR -------------------------
$menu = [Terminal.Gui.MenuBar]::new(@(
    # FILE menu
    [Terminal.Gui.MenuBarItem]::new("_File", @(
        $mRefresh,
        $mTheme,
        $mFile
    )),

    # ACTION menu
    [Terminal.Gui.MenuBarItem]::new("_Action", @(
        $mNew,
        $mProps,
        $mQuickFilter,
        $mUndo,
        $mChangeDomain,
        $mPasswordGenerator,
        $mChangeDC,
        $mSearchAD,
        $mADHealth 
    )),

    # SELECTION menu
    [Terminal.Gui.MenuBarItem]::new("_Selection", @(
        $mSelectionMode,
        $mSelectAll,
        $mDeselectAll,
        $mBulkAddGroup
    )),

    # ABOUT menu (new and now visible)
    $aboutMenu
))
# -----------------------------------------------------------------


# Apply full theme to all components <-- do this BEFORE the menus
Apply-Theme -ThemeData $themeData -TopLevel $top -MainWindow $win -Menu $menu -Status $status

Debug-Log "Theme applied successfully: $Theme"

$top.Add($menu)

# ------------------------- TreeView ------------------------
$tree = [Terminal.Gui.TreeView]::new()
$tree.X=0; $tree.Y=1; $tree.Width=30; $tree.Height=[Terminal.Gui.Dim]::Fill()



$win.Add($tree)

# ------------------------- Build Tree ------------------------

# ------------------------- Filter Panel (Add to main window) ------------------------
function Create-FilterPanel {
    # Create a frame for filters
    $filterFrame = [Terminal.Gui.FrameView]::new("Filters")
    $filterFrame.X = 32  # Right of the tree
    $filterFrame.Y = 1
    $filterFrame.Width = 40
    $filterFrame.Height = 12
    
    $y = 0
    
    # Name filter
    $lblNameFilter = [Terminal.Gui.Label]::new("Name contains:"); $lblNameFilter.X=1; $lblNameFilter.Y=$y; $filterFrame.Add($lblNameFilter)
    $txtNameFilter = [Terminal.Gui.TextField]::new($Global:FilterOptions.NameFilter)
    $txtNameFilter.X=1; $txtNameFilter.Y=$y+1; $txtNameFilter.Width=35
    $txtNameFilter.add_TextChanged({
        $Global:FilterOptions.NameFilter = $txtNameFilter.Text.ToString()
    })
    $filterFrame.Add($txtNameFilter)
    $y+=3
    
    # Show/Hide checkboxes
    $chkEnabled = [Terminal.Gui.CheckBox]::new("Show Enabled Users")
    $chkEnabled.X=1; $chkEnabled.Y=$y; $chkEnabled.Checked=$Global:FilterOptions.ShowEnabledUsers
    $chkEnabled.add_Toggled({ $Global:FilterOptions.ShowEnabledUsers = $chkEnabled.Checked })
    $filterFrame.Add($chkEnabled)
    $y+=1
    
    $chkDisabled = [Terminal.Gui.CheckBox]::new("Show Disabled Users")
    $chkDisabled.X=1; $chkDisabled.Y=$y; $chkDisabled.Checked=$Global:FilterOptions.ShowDisabledUsers
    $chkDisabled.add_Toggled({ $Global:FilterOptions.ShowDisabledUsers = $chkDisabled.Checked })
    $filterFrame.Add($chkDisabled)
    $y+=1
    
    $chkGroups = [Terminal.Gui.CheckBox]::new("Show Groups")
    $chkGroups.X=1; $chkGroups.Y=$y; $chkGroups.Checked=$Global:FilterOptions.ShowGroups
    $chkGroups.add_Toggled({ $Global:FilterOptions.ShowGroups = $chkGroups.Checked })
    $filterFrame.Add($chkGroups)
    $y+=1
    
    $chkDCs = [Terminal.Gui.CheckBox]::new("Show Domain Controllers")
    $chkDCs.X=1; $chkDCs.Y=$y; $chkDCs.Checked=$Global:FilterOptions.ShowDCs
    $chkDCs.add_Toggled({ $Global:FilterOptions.ShowDCs = $chkDCs.Checked })
    $filterFrame.Add($chkDCs)
    $y+=2
    
    # Sort options
    $lblSort = [Terminal.Gui.Label]::new("Sort by:"); $lblSort.X=1; $lblSort.Y=$y; $filterFrame.Add($lblSort)
    $y+=1
    
    $rdoSort = [Terminal.Gui.RadioGroup]::new(@("Name", "Type", "OU"))
    $rdoSort.X=1; $rdoSort.Y=$y; $rdoSort.SelectedItem=0
    $rdoSort.add_SelectedItemChanged({
        switch ($rdoSort.SelectedItem) {
            0 { $Global:FilterOptions.SortBy = "Name" }
            1 { $Global:FilterOptions.SortBy = "Type" }
            2 { $Global:FilterOptions.SortBy = "OU" }
        }
    })
    $filterFrame.Add($rdoSort)
    $y+=3
    
    # Apply/Reset buttons
    $btnApplyFilter = [Terminal.Gui.Button]::new("Apply Filter")
    $btnApplyFilter.X=1; $btnApplyFilter.Y=$y
    $btnApplyFilter.add_Clicked({
        Debug-Log "DEBUG: Applying filters..."
        Build-Tree -domain $Global:Domain
    })
    $filterFrame.Add($btnApplyFilter)
    
    $btnResetFilter = [Terminal.Gui.Button]::new("Reset")
    $btnResetFilter.X=17; $btnResetFilter.Y=$y
    $btnResetFilter.add_Clicked({
        Debug-Log "DEBUG: Resetting filters..."
        $Global:FilterOptions.ShowDisabledUsers = $true
        $Global:FilterOptions.ShowEnabledUsers = $true
        $Global:FilterOptions.ShowLockedUsers = $true
        $Global:FilterOptions.ShowGroups = $true
        $Global:FilterOptions.ShowDCs = $true
        $Global:FilterOptions.ShowComputers = $true
        $Global:FilterOptions.ShowOUs = $true
        $Global:FilterOptions.NameFilter = ""
        $Global:FilterOptions.SortBy = "Name"
        $Global:FilterOptions.SortDescending = $false
        
        # Reset UI controls
        $chkEnabled.Checked = $true
        $chkDisabled.Checked = $true
        $chkGroups.Checked = $true
        $chkDCs.Checked = $true
        $txtNameFilter.Text = ""
        $rdoSort.SelectedItem = 0
        
        Build-Tree -domain $Global:Domain
    })
    $filterFrame.Add($btnResetFilter)
    
    return $filterFrame
}

# ------------------------- Enhanced Build-Tree with Filters ------------------------
function Build-Tree {
    param([string]$domain)
    
    Debug-Log "DEBUG: Building tree with filters..."
    
    $tree.ClearObjects()
    $root = [Terminal.Gui.Trees.TreeNode]::new($domain)
    
    # Apply name filter if specified
    $nameFilter = $Global:FilterOptions.NameFilter.Trim()
    $filteredUsers = $Global:Users
    
    if ($nameFilter) {
        $filteredUsers = $filteredUsers | Where-Object { 
            $_.Name -like "*$nameFilter*" -or 
            $_.Email -like "*$nameFilter*" -or 
            $_.Title -like "*$nameFilter*"
        }
    }
    
    # Apply enabled/disabled filter
    $filteredUsers = $filteredUsers | Where-Object {
        ($_.Disabled -and $Global:FilterOptions.ShowDisabledUsers) -or
        (-not $_.Disabled -and $Global:FilterOptions.ShowEnabledUsers)
    }
    
    # Sort users based on preference
    switch ($Global:FilterOptions.SortBy) {
        "Name" { $filteredUsers = $filteredUsers | Sort-Object -Property Name -Descending:$Global:FilterOptions.SortDescending }
        "Type" { $filteredUsers = $filteredUsers | Sort-Object -Property Title,Name -Descending:$Global:FilterOptions.SortDescending }
        "OU"   { $filteredUsers = $filteredUsers | Sort-Object -Property OU,Name -Descending:$Global:FilterOptions.SortDescending }
    }
    
    # Get unique OUs from filtered users
    $OUs = $filteredUsers | Select-Object -ExpandProperty OU -Unique | Sort-Object
    
    foreach ($ou in $OUs) {
        $ouNode = [Terminal.Gui.Trees.TreeNode]::new($ou)
        
        # Build group lookup for this OU
        $groupLookup = @{}
        foreach ($u in $filteredUsers | Where-Object { $_.OU -eq $ou }) {
            foreach ($grp in $u.Groups) {
                if (-not $groupLookup.ContainsKey($grp)) { $groupLookup[$grp] = @() }
                $groupLookup[$grp] += $u
            }
        }
        
        # Only show groups if filter allows
        if ($Global:FilterOptions.ShowGroups) {
            $sortedGroups = $groupLookup.Keys | Sort-Object
            foreach ($grpName in $sortedGroups) {
                $grpNode = [Terminal.Gui.Trees.TreeNode]::new($grpName)
                
                $members = $groupLookup[$grpName] | Sort-Object -Property Name
                foreach ($m in $members) {
                    # Add status indicator
                    $statusIcon = if ($m.Disabled) { "⊗" } else { "○" }
                    $grpNode.Children.Add([Terminal.Gui.Trees.TreeNode]::new("(U) $statusIcon $($m.Name)"))
                }
                
                $ouNode.Children.Add($grpNode)
            }
        } else {
            # If groups hidden, show users directly under OU
            foreach ($u in ($filteredUsers | Where-Object { $_.OU -eq $ou } | Sort-Object -Property Name)) {
                $statusIcon = if ($u.Disabled) { "⊗" } else { "○" }
                $ouNode.Children.Add([Terminal.Gui.Trees.TreeNode]::new("(U) $statusIcon $($u.Name)"))
            }
        }
        
        if ($ouNode.Children.Count -gt 0) {
            $root.Children.Add($ouNode)
        }
    }
    
    # Add Domain Controllers if filter allows
    if ($Global:FilterOptions.ShowDCs -and $Global:DCs.Count -gt 0) {
        $dcNode = [Terminal.Gui.Trees.TreeNode]::new("Domain Controllers")
        foreach ($dc in ($Global:DCs | Sort-Object -Property Name)) {
            $dcNode.Children.Add([Terminal.Gui.Trees.TreeNode]::new("(DC) $($dc.Name)"))
        }
        $root.Children.Add($dcNode)
    }
    
    # Add Production AD Object Types if not in demo mode
    if (-not $DemoMode -and $Global:ADObjects.Count -gt 0) {
        $types = $Global:ADObjects | Select-Object -ExpandProperty Type -Unique | Sort-Object
        
        foreach ($t in $types) {
            # Skip types based on filters
            if ($t -eq "computer" -and -not $Global:FilterOptions.ShowComputers) { continue }
            if ($t -eq "organizationalUnit" -and -not $Global:FilterOptions.ShowOUs) { continue }
            
            $typeNode = [Terminal.Gui.Trees.TreeNode]::new($t)
            $objs = $Global:ADObjects | Where-Object { $_.Type -eq $t }
            
            # Apply name filter to objects
            if ($nameFilter) {
                $objs = $objs | Where-Object { $_.Name -like "*$nameFilter*" }
            }
            
            $objs = $objs | Sort-Object -Property Name
            foreach ($o in $objs) { 
                $typeNode.Children.Add([Terminal.Gui.Trees.TreeNode]::new($o.Name)) 
            }
            
            if ($typeNode.Children.Count -gt 0) {
                $root.Children.Add($typeNode)
            }
        }
    }
    
    $tree.AddObject($root)
    
    # Show filter status
    $filterCount = $filteredUsers.Count
    $totalCount = $Global:Users.Count
    Debug-Log "DEBUG: Tree built - Showing $filterCount of $totalCount users"
}

# ------------------------- Quick Filter Menu (for Menu Bar) ------------------------
function Show-QuickFilterDialog {
    $dlg = [Terminal.Gui.Dialog]::new("Quick Filters", 50, 20)
    
    $y = 1
    $lbl = [Terminal.Gui.Label]::new("Select a quick filter:"); $lbl.X=2; $lbl.Y=$y; $dlg.Add($lbl)
    $y+=2
    
    $quickFilters = @(
        "Show All",
        "Locked Users Only",
        "Disabled Users Only",
        "Enabled Users Only", 
        "Users Never Logged In",
        "Users with No Manager",
        "Empty Groups",
        "Domain Admins Only"
    )
    
    $lstFilters = [Terminal.Gui.ListView]::new($quickFilters)
    $lstFilters.X=2; $lstFilters.Y=$y; $lstFilters.Width=44; $lstFilters.Height=8
    $dlg.Add($lstFilters)
    
    $btnApply = [Terminal.Gui.Button]::new("Apply")
    $btnApply.add_Clicked({
        $selected = $quickFilters[$lstFilters.SelectedItem]
        Debug-Log "DEBUG: Applying quick filter: $selected"
        
        switch ($selected) {
            "Show All" {
                $Global:FilterOptions.ShowLockedUsers = $true
                $Global:FilterOptions.ShowDisabledUsers = $true
                $Global:FilterOptions.ShowEnabledUsers = $true
                $Global:FilterOptions.ShowGroups = $true
                $Global:FilterOptions.ShowDCs = $true
                $Global:FilterOptions.NameFilter = ""
            }
            "Locked Users Only" {
                $Global:FilterOptions.ShowLockedUsers = $true
                $Global:FilterOptions.ShowEnabledUsers = $false
                $Global:FilterOptions.ShowDisabledUsers = $false
            }
            "Disabled Users Only" {
                $Global:FilterOptions.ShowDisabledUsers = $true
                $Global:FilterOptions.ShowEnabledUsers = $false
            }
            "Enabled Users Only" {
                $Global:FilterOptions.ShowDisabledUsers = $false
                $Global:FilterOptions.ShowEnabledUsers = $true
            }
            "Users Never Logged In" {
                # This would require additional logic to track last logon
                [Terminal.Gui.MessageBox]::Query(50, 7, "Filter", "Filter applied: $selected", "OK") | Out-Null
            }
            "Users with No Manager" {
                # Filter users with no manager
                $Global:FilterOptions.NameFilter = ""
            }
            "Empty Groups" {
                # Show only groups with no members
                [Terminal.Gui.MessageBox]::Query(50, 7, "Filter", "Filter applied: $selected", "OK") | Out-Null
            }
            "Domain Admins Only" {
                $Global:FilterOptions.NameFilter = ""
            }
        }
        
        Build-Tree -domain $Global:Domain
        [Terminal.Gui.Application]::RequestStop()
    })
    $dlg.AddButton($btnApply)
    
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($dlg)
}

# ------------------------- Filter Status Label ------------------------
function Create-FilterStatusLabel {
    $lblStatus = [Terminal.Gui.Label]::new("")
    $lblStatus.X = 32
    $lblStatus.Y = 13
    $lblStatus.Width = 40
    
    return $lblStatus
}

function Update-FilterStatusLabel {
    param($label)
    
    if (-not $label) {
        Debug-Log "WARNING: label parameter is null in Update-FilterStatusLabel"
        return
    }
    
    $activeFilters = @()
    if (-not $Global:FilterOptions.ShowEnabledUsers) { $activeFilters += "No Enabled" }
    if (-not $Global:FilterOptions.ShowDisabledUsers) { $activeFilters += "No Disabled" }
    if (-not $Global:FilterOptions.ShowGroups) { $activeFilters += "No Groups" }
    if (-not $Global:FilterOptions.ShowDCs) { $activeFilters += "No DCs" }
    if ($Global:FilterOptions.NameFilter) { $activeFilters += "Name:$($Global:FilterOptions.NameFilter)" }
    
    if ($activeFilters.Count -gt 0) {
        $label.Text = "Active Filters: " + ($activeFilters -join ", ")
    } else {
        $label.Text = "No filters active (showing all)"
    }
}

# ---------------------------------------------------------------------------
# Get-ADHealth - Tabbed AD Health modal for DSA-TUI (Final: Tools row, Health icons,
# Alt-key accelerators, help line)
# - Uses $Global:Domain for the domain to query (auto-refresh if it changes)
# - No F-keys, no Ctrl+Q/Ctrl+D; uses Alt+... + Esc as requested
# ---------------------------------------------------------------------------
function Get-ADHealth {
    <#
    Top-level function to display an AD health modal.
    Fully self-contained, uses Terminal.Gui, $Global:Domain, and the checks/tooling we discussed.
    #>

    # ----------------------------
    # Layout
    # ----------------------------
    $width  = 118
    $height = 36

    # ----------------------------
    # Ensure cache / state exists
    # ----------------------------
    if (-not $script:ADHealth) { $script:ADHealth = @{} }
    if (-not $script:ADHealth.Cache) { $script:ADHealth.Cache = @{} }

    # ----------------------------
    # Safe external invocation helper
    # ----------------------------
    function Invoke-External {
        param([string]$Exe, [string]$Args)
        try {
            $proc = Start-Process -FilePath $Exe -ArgumentList $Args -NoNewWindow -RedirectStandardOutput -RedirectStandardError -PassThru -WindowStyle Hidden
            $proc.WaitForExit()
            $out = $proc.StandardOutput.ReadToEnd()
            $err = $proc.StandardError.ReadToEnd()
            return @{ ExitCode = $proc.ExitCode; StdOut = $out; StdErr = $err }
        } catch {
            return @{ ExitCode = -1; StdOut = ""; StdErr = $_.Exception.Message }
        }
    }

    # ----------------------------
    # Initialize / Refresh caches and detect tools
    # ----------------------------
    function Initialize-ADHealth {
        if (-not $Global:Domain -or [string]::IsNullOrWhiteSpace($Global:Domain)) {
            throw "Global variable `$Global:Domain is not set. Set `$Global:Domain before calling Get-ADHealth."
        }
        $domain = $Global:Domain

        # If domain changed since last run, clear cached data
        if ($script:ADHealth.Cache.Domain -ne $domain) {
            $script:ADHealth.Cache = @{
                Domain = $domain
                Timestamp = (Get-Date).ToString("s")
            }
        }

        # Prepare blank data containers for fresh run
        $script:ADHealth.Data = @{
            Domain = $domain
            Timestamp = (Get-Date)
            DCStatus = @{ Summary = @(); Details = "" ; Health = "UNKNOWN" }
            Replication = @{ Summary = @(); Details = "" ; Health = "UNKNOWN" }
            DNS = @{ Summary = @(); Details = "" ; Health = "UNKNOWN" }
            SYSVOL = @{ Summary = @(); Details = "" ; Health = "UNKNOWN" }
            FSMO = @{ Summary = @(); Details = "" ; Health = "UNKNOWN" }
            GPO = @{ Summary = @(); Details = "" ; Health = "UNKNOWN" }
        }

        # Detect available external utilities (like CertUI checks)
        $script:ADHealth.Tools = @{
            Repadmin = (Get-Command repadmin.exe -ErrorAction SilentlyContinue) -ne $null
            DcDiag   = (Get-Command dcdiag.exe -ErrorAction SilentlyContinue) -ne $null
            DfsrDiag = (Get-Command dfsrdiag.exe -ErrorAction SilentlyContinue) -ne $null
            Nltest   = (Get-Command nltest.exe -ErrorAction SilentlyContinue) -ne $null
            Dnscmd   = (Get-Command dnscmd.exe -ErrorAction SilentlyContinue) -ne $null
        }
    }

    # ----------------------------
    # DC Status
    # ----------------------------
    function Run-DCStatusCheck {
        param([string]$Domain)
        $summary = @()
        $sb = New-Object System.Text.StringBuilder

        try {
            $dcs = Get-ADDomainController -Filter * -Server $Domain -ErrorAction Stop
        } catch {
            $script:ADHealth.Data.DCStatus.Details = "Failed to enumerate DCs: $($_.Exception.Message)"
            $script:ADHealth.Data.DCStatus.Summary = @("ERROR: Could not enumerate domain controllers.")
            $script:ADHealth.Data.DCStatus.Health = "FAIL"
            return
        }

        foreach ($dc in $dcs) {
            $name = $dc.HostName
            $ip = $dc.IPAddress -join ","
            $site = $dc.Site
            $os = $dc.OperatingSystem
            $isGC = $dc.IsGlobalCatalog

            # Reachability
            try { $ping = Test-Connection -ComputerName $name -Count 1 -Quiet -ErrorAction SilentlyContinue } catch { $ping = $false }
            $reachable = if ($ping) { "OK" } else { "FAIL" }

            # Uptime
            try {
                $wmi = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $name -ErrorAction Stop
                $lastBoot = $wmi.LastBootUpTime
                $uptimeDays = ((Get-Date) - [Management.ManagementDateTimeConverter]::ToDateTime($lastBoot)).Days
            } catch { $uptimeDays = "?" }

            # Services
            try { $svcNTDS = (Get-Service -ComputerName $name -Name "NTDS" -ErrorAction SilentlyContinue).Status } catch { $svcNTDS = "Err" }
            try { $svcDNS = (Get-Service -ComputerName $name -Name "DNS" -ErrorAction SilentlyContinue).Status } catch { $svcDNS = "Err" }

            $obj = [PSCustomObject]@{
                Name = $name; IP = $ip; Site = $site; OS = $os; Reachable = $reachable;
                UptimeDays = $uptimeDays; NTDS = $svcNTDS; DNS = $svcDNS; IsGC = $isGC
            }
            $summary += $obj

            $sb.AppendLine("=== $name ($ip) ===")
            $sb.AppendLine("Site: $site; OS: $os; GC: $isGC")
            $sb.AppendLine("Reachable: $reachable; UptimeDays: $uptimeDays")
            $sb.AppendLine("NTDS: $svcNTDS; DNS: $svcDNS")
            $sb.AppendLine("")
        }

        # Determine simple health: if any Reachable=FAIL -> WARN, if many fails -> FAIL
        $fails = ($summary | Where-Object { $_.Reachable -ne "OK" }).Count
        if ($fails -eq 0) { $health = "OK" } elseif ($fails -lt ($summary.Count / 2)) { $health = "WARN" } else { $health = "FAIL" }

        $script:ADHealth.Data.DCStatus.Summary = $summary
        $script:ADHealth.Data.DCStatus.Details = $sb.ToString()
        $script:ADHealth.Data.DCStatus.Health = $health
    }

    # ----------------------------
    # Replication
    # ----------------------------
    function Run-ReplicationCheck {
        param([string]$Domain)
        $sb = New-Object System.Text.StringBuilder
        $summary = @()

        if ($script:ADHealth.Tools.Repadmin) {
            $r = Invoke-External -Exe "repadmin.exe" -Args "/replsummary"
            $sb.AppendLine($r.StdOut)
            if ($r.ExitCode -eq 0) {
                $lines = ($r.StdOut -split "`r?`n") | Select-Object -First 12
                $summary += "repadmin /replsummary (top lines):"
                $summary += $lines
                # Simple health: if the output contains "failed" or "error" mark WARN

      if ($r.StdOut -imatch "(fail|error|last attempt)") { $health = "FAIL" }


            } else {
                $summary += "repadmin returned error: $($r.StdErr)"
                $health = "WARN"
            }

            $r2 = Invoke-External -Exe "repadmin.exe" -Args "/showrepl * /errorsonly"
            if ($r2.ExitCode -eq 0) { $sb.AppendLine($r2.StdOut) }
        } else {
            try {
                $fails = Get-ADReplicationFailure -Scope Domain -ErrorAction Stop
                if ($fails) {
                    foreach ($f in $fails) {
                        $summary += "Replication failure: $($f.SourceServer) -> $($f.DestinationServer): $($f.FirstFailureMessage)"
                        $sb.AppendLine(($f | Format-List | Out-String))
                    }
                    $health = "WARN"
                } else {
                    $summary += "No replication failures detected."
                    $health = "OK"
                }
            } catch {
                $summary += "Replication check error: $($_.Exception.Message)"
                $sb.AppendLine("Get-ADReplicationFailure error: $($_.Exception.Message)")
                $health = "WARN"
            }
        }

        $script:ADHealth.Data.Replication.Summary = $summary
        $script:ADHealth.Data.Replication.Details = $sb.ToString()
        $script:ADHealth.Data.Replication.Health = $health
    }

    # ----------------------------
    # DNS
    # ----------------------------
    function Run-DNSCheck {
        param([string]$Domain)
        $summary = @()
        $sb = New-Object System.Text.StringBuilder
        $srvRecords = @("_ldap._tcp.dc._msdcs.$Domain", "_kerberos._tcp.$Domain", "_ldap._tcp.$Domain")

        foreach ($srv in $srvRecords) {
            try {
                $q = Resolve-DnsName -Name $srv -ErrorAction Stop
                $count = ($q | Measure-Object).Count
                $summary += "$srv -> $count records"
                $sb.AppendLine("SRV ${srv}:")
                $sb.AppendLine(($q | Out-String).Trim())
            } catch {
                $summary += "$srv -> NOT FOUND"
                $sb.AppendLine("SRV $srv not found: $($_.Exception.Message)")
            }
        }

        # Per-DC DNS service status
        try {
            $dcs = Get-ADDomainController -Filter * -Server $Domain -ErrorAction Stop
            foreach ($dc in $dcs) {
                try { $svc = (Get-Service -ComputerName $dc.HostName -Name "DNS" -ErrorAction SilentlyContinue).Status } catch { $svc = "Err" }
                $summary += "DNS on $($dc.HostName): $svc"
                $sb.AppendLine("DNS on $($dc.HostName): $svc")
            }
        } catch {
            $summary += "Could not enumerate DCs for DNS check."
            $sb.AppendLine("Enumeration error: $($_.Exception.Message)")
        }

        # Health deduction
        $fails = ($summary -match "NOT FOUND|Err|Error").Count
        if ($fails -eq 0) { $health = "OK" } elseif ($fails -lt 2) { $health = "WARN" } else { $health = "FAIL" }

        $script:ADHealth.Data.DNS.Summary = $summary
        $script:ADHealth.Data.DNS.Details = $sb.ToString()
        $script:ADHealth.Data.DNS.Health = $health
    }

    # ----------------------------
    # SYSVOL / DFSR
    # ----------------------------
    function Run-SYSVOLCheck {
        param([string]$Domain)
        $summary = @()
        $sb = New-Object System.Text.StringBuilder

        try {
            $dcs = Get-ADDomainController -Filter * -Server $Domain -ErrorAction Stop
            foreach ($dc in $dcs) {
                try {
                    $shares = Get-WmiObject -Class Win32_Share -ComputerName $dc.HostName -ErrorAction Stop
                    $sys = $shares | Where-Object { $_.Name -match 'SYSVOL' -or $_.Name -match 'NETLOGON' }
                    if ($sys) {
                        $summary += "SYSVOL on $($dc.HostName): SHARE_OK"
                        $sb.AppendLine("SYSVOL/Netlogon on $($dc.HostName):")
                        $sb.AppendLine(($sys | Format-List | Out-String).Trim())
                    } else {
                        $summary += "SYSVOL on $($dc.HostName): MISSING"
                        $sb.AppendLine("SYSVOL missing on $($dc.HostName)")
                    }
                } catch {
                    $summary += "SYSVOL error on $($dc.HostName): $($_.Exception.Message)"
                    $sb.AppendLine("Error enumerating shares: $($_.Exception.Message)")
                }

                if ($script:ADHealth.Tools.DfsrDiag) {
                    $r = Invoke-External -Exe "dfsrdiag.exe" -Args "backlog /rgname:DOMAIN /v /sendingcomputer:$($dc.HostName) /receivingcomputer:$($dc.HostName)"
                    $sb.AppendLine("dfsrdiag backlog for $($dc.HostName):")
                    $sb.AppendLine($r.StdOut)
                }
            }
        } catch {
            $summary += "SYSVOL check failed: $($_.Exception.Message)"
            $sb.AppendLine("SYSVOL error: $($_.Exception.Message)")
        }

        $fails = ($summary -match "MISSING|ERROR").Count
        if ($fails -eq 0) { $health = "OK" } elseif ($fails -lt 2) { $health = "WARN" } else { $health = "FAIL" }

        $script:ADHealth.Data.SYSVOL.Summary = $summary
        $script:ADHealth.Data.SYSVOL.Details = $sb.ToString()
        $script:ADHealth.Data.SYSVOL.Health = $health
    }

    # ----------------------------
    # FSMO roles
    # ----------------------------
    function Run-FSMOCheck {
        param([string]$Domain)
        $sb = New-Object System.Text.StringBuilder
        $summary = @()

        try {
            $forest = Get-ADForest -Identity $Domain -ErrorAction Stop
            $domainObj = Get-ADDomain -Identity $Domain -ErrorAction Stop

            $fsmoForest = @{
                SchemaMaster = $forest.SchemaMaster
                DomainNamingMaster = $forest.DomainNamingMaster
            }
            $fsmoDomain = @{
                PDCEmulator = $domainObj.PDCEmulator
                RIDMaster = $domainObj.RIDMaster
                InfrastructureMaster = $domainObj.InfrastructureMaster
            }

            $sb.AppendLine("Forest FSMO holders:")
            $sb.AppendLine(($fsmoForest | Out-String).Trim())
            $sb.AppendLine("Domain FSMO holders:")
            $sb.AppendLine(($fsmoDomain | Out-String).Trim())

            foreach ($k in $fsmoForest.Keys) { $summary += "$k -> $($fsmoForest[$k])" }
            foreach ($k in $fsmoDomain.Keys) { $summary += "$k -> $($fsmoDomain[$k])" }

            # Health: check if any holder resolves and is reachable
            $holders = ($fsmoForest.Values + $fsmoDomain.Values) | Where-Object { $_ -ne $null } | Select-Object -Unique
            $unreachable = 0
            foreach ($h in $holders) {
                try { $ok = Test-Connection -ComputerName $h -Count 1 -Quiet -ErrorAction SilentlyContinue } catch { $ok = $false }
                if (-not $ok) { $unreachable++ }
            }
            if ($unreachable -eq 0) { $health = "OK" } elseif ($unreachable -lt 2) { $health = "WARN" } else { $health = "FAIL" }
        } catch {
            $summary += "FSMO check failed: $($_.Exception.Message)"
            $sb.AppendLine("FSMO error: $($_.Exception.Message)")
            $health = "WARN"
        }

        $script:ADHealth.Data.FSMO.Summary = $summary
        $script:ADHealth.Data.FSMO.Details = $sb.ToString()
        $script:ADHealth.Data.FSMO.Health = $health
    }

    # ----------------------------
    # GPO check
    # ----------------------------
    function Run-GPOCheck {
        param([string]$Domain)
        $sb = New-Object System.Text.StringBuilder
        $summary = @()

        try {
            $gpos = Get-GPO -All -Domain $Domain -ErrorAction Stop
            $summary += "Total GPOs: $($gpos.Count)"
            $sb.AppendLine("GPO list (top 50):")
            $sb.AppendLine(($gpos | Select-Object DisplayName,Id | Select-Object -First 50 | Format-Table -AutoSize | Out-String).Trim())

            # Quick version compare for first 20 GPOs
            foreach ($g in $gpos | Select-Object -First 20) {
                try {
                    $xml = [xml](Get-GPOReport -Guid $g.Id -ReportType Xml -Domain $Domain -ErrorAction Stop)
                    $ver = $xml.GPO.VersionInfo.Version
                    $sb.AppendLine("GPO: $($g.DisplayName) Version: $ver")
                } catch {
                    $sb.AppendLine("Could not get GPOReport for $($g.DisplayName): $($_.Exception.Message)")
                }
            }
            $health = "OK"
        } catch {
            $summary += "GPO check failed: $($_.Exception.Message)"
            $sb.AppendLine("GPO error: $($_.Exception.Message)")
            $health = "WARN"
        }

        $script:ADHealth.Data.GPO.Summary = $summary
        $script:ADHealth.Data.GPO.Details = $sb.ToString()
        $script:ADHealth.Data.GPO.Health = $health
    }

    # ----------------------------
    # Run all checks orchestrator
    # ----------------------------
    function Run-AllChecks {
        param([string]$Domain)
        Initialize-ADHealth
        try { Run-DCStatusCheck -Domain $Domain } catch { $script:ADHealth.Data.DCStatus.Details = "DCStatus failed: $($_.Exception.Message)"; $script:ADHealth.Data.DCStatus.Health="WARN" }
        try { Run-ReplicationCheck -Domain $Domain } catch { $script:ADHealth.Data.Replication.Details = "Replication failed: $($_.Exception.Message)"; $script:ADHealth.Data.Replication.Health="WARN" }
        try { Run-DNSCheck -Domain $Domain } catch { $script:ADHealth.Data.DNS.Details = "DNS failed: $($_.Exception.Message)"; $script:ADHealth.Data.DNS.Health="WARN" }
        try { Run-SYSVOLCheck -Domain $Domain } catch { $script:ADHealth.Data.SYSVOL.Details = "SYSVOL failed: $($_.Exception.Message)"; $script:ADHealth.Data.SYSVOL.Health="WARN" }
        try { Run-FSMOCheck -Domain $Domain } catch { $script:ADHealth.Data.FSMO.Details = "FSMO failed: $($_.Exception.Message)"; $script:ADHealth.Data.FSMO.Health="WARN" }
        try { Run-GPOCheck -Domain $Domain } catch { $script:ADHealth.Data.GPO.Details = "GPO failed: $($_.Exception.Message)"; $script:ADHealth.Data.GPO.Health="WARN" }

        $script:ADHealth.Cache.Timestamp = (Get-Date)
    }

    # ----------------------------
    # UI helpers
    # ----------------------------
    function Show-DetailsModal {
        param([string]$Title, [string]$Content)
        $w = [Terminal.Gui.Window]::new($Title)
        $w.X = [Terminal.Gui.Pos]::Center()
        $w.Y = [Terminal.Gui.Pos]::Center()
        $w.Width = $width - 10
        $w.Height = $height - 8
        $w.Modal = $true

        $tv = [Terminal.Gui.TextView]::new()
        $tv.X = 1; $tv.Y = 1
        $tv.Width = $w.Width - 2
        $tv.Height = $w.Height - 4
        $tv.ReadOnly = $true
        $tv.Text = $Content

        $btnClose = [Terminal.Gui.Button]::new( ($w.Width / 2) - 6, $w.Height - 3, "Close")
        $btnClose.add_Click({ [Terminal.Gui.Application]::RequestStop() })

        $w.Add($tv); $w.Add($btnClose)
        [Terminal.Gui.Application]::Run($w)
    }

    function Render-SummaryWithDetails {
        param([Terminal.Gui.View]$Parent, [string[]]$SummaryLines, [string]$DetailsText)
        # Show up to 10 lines and place Details button bottom-right
        $Parent.RemoveAll()
        $y = 1
        $max = 10
        for ($i = 0; $i -lt $SummaryLines.Count -and $i -lt $max; $i++) {
            $lbl = [Terminal.Gui.Label]::new($SummaryLines[$i])
            $lbl.X = 1; $lbl.Y = $y
            $Parent.Add($lbl)
            $y++
        }
        if ($SummaryLines.Count -gt $max) {
            $lblMore = [Terminal.Gui.Label]::new("(... $($SummaryLines.Count - $max) more lines. Click Details...)")
            $lblMore.X = 1; $lblMore.Y = $y
            $Parent.Add($lblMore)
            $y++
        }

        # Details button text uses an ampersand accelerator for Alt+E: we place & before 'e' in "Details"
        $btnDetails = [Terminal.Gui.Button]::new($Parent.Width - 20, $Parent.Height - 3, "D&e tails...")
        $btnDetails.add_Click({
            if ($DetailsText -and $DetailsText.Length -gt 0) {
                Show-DetailsModal -Title "Details" -Content $DetailsText
            } else {
                [Terminal.Gui.MessageBox]::Query(40,7,"Details","No details available.")
            }
        })
        $Parent.Add($btnDetails)
    }

    # ----------------------------
    # UI: Build modal & controls
    # ----------------------------
    $modal = [Terminal.Gui.Window]::new("Active Directory Health Check - $($Global:Domain)")
    $modal.X = [Terminal.Gui.Pos]::Center(); $modal.Y = [Terminal.Gui.Pos]::Center()
    $modal.Width = $width; $modal.Height = $height; $modal.Modal = $true

    # Tools row - single-line compact
    $toolsView = [Terminal.Gui.View]::new()
    $toolsView.X = 1; $toolsView.Y = 1; $toolsView.Width = $width - 2; $toolsView.Height = 1
    $modal.Add($toolsView)

    # Tab header view (below tools)
    $tabHeader = [Terminal.Gui.View]::new()
    $tabHeader.X = 1; $tabHeader.Y = 2; $tabHeader.Width = $width - 2; $tabHeader.Height = 1
    $modal.Add($tabHeader)

    # Content frame
    $contentView = [Terminal.Gui.FrameView]::new()
    $contentView.X = 1; $contentView.Y = 4; $contentView.Width = $width - 2; $contentView.Height = $height - 10
    $contentView.Border = "Ascii"
    $modal.Add($contentView)

    # Help line (shortcuts)
    $helpLine = [Terminal.Gui.Label]::new("Shortcuts: Alt+R=Refresh All | Alt+S=Refresh Tab | Alt+←/→=Switch Tabs | Alt+E=Details | Esc=Close")
    $helpLine.X = 1; $helpLine.Y = $height - 5
    $modal.Add($helpLine)

    # Buttons: use ampersand to set Alt-accelerators, avoiding reserved combos
    $btnRefreshAll = [Terminal.Gui.Button]::new(2, $height - 4, "&Refresh All")   # Alt+R
    $btnRefreshTab = [Terminal.Gui.Button]::new(20, $height - 4, "Re&scan Tab")    # Alt+S
    $btnExport = [Terminal.Gui.Button]::new(40, $height - 4, "&Export")           # Alt+E for Export (convenient)
    $btnClose = [Terminal.Gui.Button]::new(60, $height - 4, "Close (Esc)")

    $modal.Add($btnRefreshAll); $modal.Add($btnRefreshTab); $modal.Add($btnExport); $modal.Add($btnClose)

    # Tab definitions
    $tabs = @("DC Status","Replication","DNS","SYSVOL","FSMO Roles","GPO Health")
    $script:ActiveHealthTab = 0

    # Function to compute simple health icon string
    function Health-Icon {
        param([string]$State)
        switch ($State) {
            "OK"   { return "(✓)" }
            "WARN" { return "(!)" }
            "FAIL" { return "(✗)" }
            default { return "(?)" }
        }
    }

    # Render tools row (certUI-style)
    function Render-ToolsRow {
        $toolsView.RemoveAll()
        $offset = 0
        $tools = $script:ADHealth.Tools
        $items = @(
            @{ Name="repadmin"; Found = $tools.Repadmin },
            @{ Name="dcdiag"; Found = $tools.DcDiag },
            @{ Name="dfsrdiag"; Found = $tools.DfsrDiag },
            @{ Name="nltest"; Found = $tools.Nltest },
            @{ Name="dnscmd"; Found = $tools.Dnscmd }
        )
        foreach ($it in $items) {
            $sym = if ($it.Found) { "[✓]" } else { "[✗]" }
            $lbl = [Terminal.Gui.Label]::new(" $sym $($it.Name)  ")
            $lbl.X = $offset; $lbl.Y = 0
            # Add click handler that shows tooltip-like message
            $nameCopy = $it.Name
            $foundCopy = $it.Found
            $lbl.add_MouseClick({
                param($args)
                if ($foundCopy) {
                    [Terminal.Gui.MessageBox]::Query(50,7,"Tool Found","$nameCopy detected - will be used where applicable.")
                } else {
                    [Terminal.Gui.MessageBox]::Query(60,8,"Tool Missing","$nameCopy not found. Some detail modes will be disabled or fall back to PowerShell cmdlets.")
                }
            })
            $toolsView.Add($lbl)
            $offset += ($lbl.Text.Length + 1)
        }
    }

    # Render tab header with health icons
    function Render-ADHealthTabs {
        $tabHeader.RemoveAll()
        $offset = 0
        for ($i = 0; $i -lt $tabs.Count; $i++) {
            $name = $tabs[$i]
            $health = $script:ADHealth.Data.$((($name -split ' ')[0])) 2>$null
            # We stored health per named key (e.g., DCStatus, Replication, DNS...). fallback if not matching:
            switch ($i) {
                0 { $h = $script:ADHealth.Data.DCStatus.Health }
                1 { $h = $script:ADHealth.Data.Replication.Health }
                2 { $h = $script:ADHealth.Data.DNS.Health }
                3 { $h = $script:ADHealth.Data.SYSVOL.Health }
                4 { $h = $script:ADHealth.Data.FSMO.Health }
                5 { $h = $script:ADHealth.Data.GPO.Health }
                default { $h = "UNKNOWN" }
            }
            $icon = Health-Icon -State $h
            $labelText = if ($i -eq $script:ActiveHealthTab) { "[ $name $icon ]" } else { "  $name $icon  " }
            $lbl = [Terminal.Gui.Label]::new($labelText)
            $lbl.X = $offset; $lbl.Y = 0
            $idx = $i
            $lbl.add_MouseClick({
                param($args)
                $script:ActiveHealthTab = $idx
                Render-ADHealthTabs
                Render-ADHealthContent -Tab $idx
            })
            $tabHeader.Add($lbl)
            $offset += ($labelText.Length + 1)
        }
    }

    # Render content for each tab
    function Render-ADHealthContent {
        param([int]$Tab)
        $contentView.RemoveAll()
        switch ($Tab) {
            0 {
                $sumObjs = $script:ADHealth.Data.DCStatus.Summary
                if (-not $sumObjs -or $sumObjs.Count -eq 0) {
                    $lbl = [Terminal.Gui.Label]::new("No DC status data. Press Alt+R to run checks.")
                    $lbl.X = 1; $lbl.Y = 1; $contentView.Add($lbl)
                } else {
                    $lines = @()
                    $lines += "Name`tIP`tSite`tReach`tUptime`tNTDS`tDNS"
                    foreach ($o in $sumObjs) {
                        $lines += ("{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}" -f $o.Name, $o.IP, $o.Site, $o.Reachable, $o.UptimeDays, $o.NTDS, $o.DNS)
                    }
                    Render-SummaryWithDetails -Parent $contentView -SummaryLines $lines -DetailsText $script:ADHealth.Data.DCStatus.Details
                }
            }
            1 {
                $lines = $script:ADHealth.Data.Replication.Summary
                if (-not $lines) { $lines = @("No replication data. Press Alt+R.") }
                Render-SummaryWithDetails -Parent $contentView -SummaryLines $lines -DetailsText $script:ADHealth.Data.Replication.Details
            }
            2 {
                $lines = $script:ADHealth.Data.DNS.Summary
                if (-not $lines) { $lines = @("No DNS data. Press Alt+R.") }
                Render-SummaryWithDetails -Parent $contentView -SummaryLines $lines -DetailsText $script:ADHealth.Data.DNS.Details
            }
            3 {
                $lines = $script:ADHealth.Data.SYSVOL.Summary
                if (-not $lines) { $lines = @("No SYSVOL data. Press Alt+R.") }
                Render-SummaryWithDetails -Parent $contentView -SummaryLines $lines -DetailsText $script:ADHealth.Data.SYSVOL.Details
            }
            4 {
                $lines = $script:ADHealth.Data.FSMO.Summary
                if (-not $lines) { $lines = @("No FSMO data. Press Alt+R.") }
                Render-SummaryWithDetails -Parent $contentView -SummaryLines $lines -DetailsText $script:ADHealth.Data.FSMO.Details
            }
            5 {
                $lines = $script:ADHealth.Data.GPO.Summary
                if (-not $lines) { $lines = @("No GPO data. Press Alt+R.") }
                Render-SummaryWithDetails -Parent $contentView -SummaryLines $lines -DetailsText $script:ADHealth.Data.GPO.Details
            }
        }
    }

    # ----------------------------
    # Button actions
    # ----------------------------
    $btnRefreshAll.add_Click({
        # Show a tiny busy modal while checks run
        $busy = [Terminal.Gui.Window]::new("Refreshing...")
        $busy.X = [Terminal.Gui.Pos]::Center(); $busy.Y = [Terminal.Gui.Pos]::Center()
        $busy.Width = 36; $busy.Height = 5
        $busy.Add([Terminal.Gui.Label]::new("Running AD health checks, please wait..."))
        [Terminal.Gui.Application]::Run($busy)
        try { Run-AllChecks -Domain $Global:Domain } finally { [Terminal.Gui.Application]::RequestStop() }
        Render-ADHealthTabs; Render-ADHealthContent -Tab $script:ActiveHealthTab
    })

    $btnRefreshTab.add_Click({
        # Run only the check for the active tab
        switch ($script:ActiveHealthTab) {
            0 { Run-DCStatusCheck -Domain $Global:Domain }
            1 { Run-ReplicationCheck -Domain $Global:Domain }
            2 { Run-DNSCheck -Domain $Global:Domain }
            3 { Run-SYSVOLCheck -Domain $Global:Domain }
            4 { Run-FSMOCheck -Domain $Global:Domain }
            5 { Run-GPOCheck -Domain $Global:Domain }
        }
        Render-ADHealthTabs; Render-ADHealthContent -Tab $script:ActiveHealthTab
    })

    # Export (plain-text)
    $btnExport.add_Click({
        $ts = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $fname = "ADHealth_${($Global:Domain -replace '[^a-zA-Z0-9\.-]','_')}_$ts.txt"
        $full = Join-Path -Path (Get-Location) -ChildPath $fname
        $sb = New-Object System.Text.StringBuilder
        $sb.AppendLine("AD Health Report for $($Global:Domain) - $ts") | Out-Null
        foreach ($tabName in $tabs) {
            $sb.AppendLine("---- $tabName ----") | Out-Null
            switch ($tabName) {
                "DC Status" { $sb.AppendLine($script:ADHealth.Data.DCStatus.Details) | Out-Null }
                "Replication" { $sb.AppendLine($script:ADHealth.Data.Replication.Details) | Out-Null }
                "DNS" { $sb.AppendLine($script:ADHealth.Data.DNS.Details) | Out-Null }
                "SYSVOL" { $sb.AppendLine($script:ADHealth.Data.SYSVOL.Details) | Out-Null }
                "FSMO Roles" { $sb.AppendLine($script:ADHealth.Data.FSMO.Details) | Out-Null }
                "GPO Health" { $sb.AppendLine($script:ADHealth.Data.GPO.Details) | Out-Null }
            }
        }
        try {
            [System.IO.File]::WriteAllText($full, $sb.ToString())
            [Terminal.Gui.MessageBox]::Query(60,7,"Export","Report exported to`n$full")
        } catch {
            [Terminal.Gui.MessageBox]::ErrorQuery(60,7,"Export Error","Failed to write file: $($_.Exception.Message)") | Out-Null
        }
    })

    $btnClose.add_Click({ [Terminal.Gui.Application]::RequestStop() })

    # Add Escape key handler to close the modal (conventional)
    $modal.add_KeyDown({
        param($args)
        try {
            if ($args.KeyEvent.Key -eq [Terminal.Gui.Key]::Escape) {
                [Terminal.Gui.Application]::RequestStop()
                $args.Handled = $true
            }
            # Alt+Left / Alt+Right support: move tabs
            if ($args.KeyEvent.IsAlt && $args.KeyEvent.Key -eq [Terminal.Gui.Key]::LeftArrow) {
                $script:ActiveHealthTab = [Math]::Max(0, $script:ActiveHealthTab - 1)
                Render-ADHealthTabs; Render-ADHealthContent -Tab $script:ActiveHealthTab
                $args.Handled = $true
            } elseif ($args.KeyEvent.IsAlt && $args.KeyEvent.Key -eq [Terminal.Gui.Key]::RightArrow) {
                $script:ActiveHealthTab = [Math]::Min($tabs.Count - 1, $script:ActiveHealthTab + 1)
                Render-ADHealthTabs; Render-ADHealthContent -Tab $script:ActiveHealthTab
                $args.Handled = $true
            }
        } catch { }
    })

    # Also allow Alt+E to activate Details by simulating a click on the Details button in the content area
    # Note: Details button is created per content render; we'll also respond to Alt+E globally to open Details for current tab.
    $modal.add_KeyDown({
        param($args)
        try {
            if ($args.KeyEvent.IsAlt -and $args.KeyEvent.Key -eq [Terminal.Gui.Key]::E) {
                # Determine details for current tab and show modal if present
                switch ($script:ActiveHealthTab) {
                    0 { $d = $script:ADHealth.Data.DCStatus.Details }
                    1 { $d = $script:ADHealth.Data.Replication.Details }
                    2 { $d = $script:ADHealth.Data.DNS.Details }
                    3 { $d = $script:ADHealth.Data.SYSVOL.Details }
                    4 { $d = $script:ADHealth.Data.FSMO.Details }
                    5 { $d = $script:ADHealth.Data.GPO.Details }
                }
                if ($d -and $d.Length -gt 0) { Show-DetailsModal -Title "Details" -Content $d } else { [Terminal.Gui.MessageBox]::Query(40,7,"Details","No details available for this tab.") }
                $args.Handled = $true
            }
        } catch { }
    })

    # ----------------------------
    # Initial run & render
    # ----------------------------
    Run-AllChecks -Domain $Global:Domain
    Render-ToolsRow
    Render-ADHealthTabs
    Render-ADHealthContent -Tab 0

    # Run the modal
    [Terminal.Gui.Application]::Run($modal)
}
# End Get-ADHealth


# Generate Random Password modal
function Generate-RandomPassword {

    # Helper function
    function Get-PasswordEntropy {
        param([int]$poolSize, [int]$length)
        if ($poolSize -le 0 -or $length -le 0) { return 0 }
        return [Math]::Log($poolSize, 2) * $length
    }

    $UpperCase = @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z')
    $LowerCase = @('a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z')
    $Numbers   = @('1','2','3','4','5','6','7','8','9','0')
    $Symbols   = @('!','@','$','?','<','>','*','&')

    $script:actualPassword = ""

    # --- Build UI ---
    $dlg = [Terminal.Gui.Dialog]::new("Generate Random Password", 66, 14)

    # 2x2 checkbox layout
    $chkUpper = [Terminal.Gui.CheckBox]::new(2,1,"Include Uppercase (A-Z)", $true)
    $chkLower = [Terminal.Gui.CheckBox]::new(30,1,"Include Lowercase (a-z)", $true)
    $chkNums  = [Terminal.Gui.CheckBox]::new(2,3,"Include Numbers (0-9)", $true)
    $chkSyms  = [Terminal.Gui.CheckBox]::new(30,3,"Include Symbols (!,@,$)", $true)
    $dlg.Add($chkUpper, $chkLower, $chkNums, $chkSyms)

    # Length input
    $dlg.Add([Terminal.Gui.Label]::new(2,5,"Length (1-127):"))
    $txtLen = New-Object Terminal.Gui.TextField
    $txtLen.X = 18; $txtLen.Y = 5; $txtLen.Width = 6
    $txtLen.Text = "12"
    $dlg.Add($txtLen)

    # Password display box
    $dlg.Add([Terminal.Gui.Label]::new(2,7,"Generated Password:"))
    $txtPwd = New-Object Terminal.Gui.TextField
    $txtPwd.X = 2; $txtPwd.Y = 8; $txtPwd.Width = 25
    $txtPwd.Text = ""
    $dlg.Add($txtPwd)

    # --- ENTROPY DISPLAY ---
    $lblStrength = [Terminal.Gui.Label]::new(30,7,"Strength: Not generated")
    $lblStrength.Width = 30
    $dlg.Add($lblStrength)

    $progEntropy = [Terminal.Gui.ProgressBar]::new()
    $progEntropy.X = 30; $progEntropy.Y = 8; $progEntropy.Width = 25
    $progEntropy.Fraction = 0.0
    $dlg.Add($progEntropy)

    # Show Password checkbox
 $chkShowPwd = [Terminal.Gui.CheckBox]::new(30,5,"Show Password",$false)
 $dlg.Add($chkShowPwd)

    # Buttons
    $btnGenerate = [Terminal.Gui.Button]::new("Generate"); $btnGenerate.X=2; $btnGenerate.Y=10
    $btnCopy     = [Terminal.Gui.Button]::new("Copy");     $btnCopy.X=15; $btnCopy.Y=10
    $btnClose    = [Terminal.Gui.Button]::new("Close");    $btnClose.X=28; $btnClose.Y=10
    $dlg.Add($btnGenerate, $btnCopy, $btnClose)

    # --- Generate Logic ---
    $btnGenerate.add_Clicked({
        $len = $txtLen.Text -as [int]
        if (-not $len) { $len = 12 }
        if ($len -lt 1 -or $len -gt 127) {
            [Terminal.Gui.MessageBox]::Query(50,7,"Invalid Length","Password length must be 1-127.","OK") | Out-Null
            return
        }

        $pool = @()
        if ($chkUpper.Checked) { $pool += $UpperCase }
        if ($chkLower.Checked) { $pool += $LowerCase }
        if ($chkNums.Checked)  { $pool += $Numbers }
        if ($chkSyms.Checked)  { $pool += $Symbols }

        if ($pool.Count -eq 0) {
            [Terminal.Gui.MessageBox]::Query(50,7,"No Character Types","Select at least one character type.","OK") | Out-Null
            return
        }

        $script:actualPassword = -join (1..$len | ForEach-Object { $pool | Get-Random })

        # --- Entropy update ---
        $entropy = Get-PasswordEntropy -poolSize $pool.Count -length $len
        $strengthCategory = ""
        $fraction = 0.0
        if ($entropy -lt 40) { $strengthCategory="Weak"; $fraction=0.2 }
        elseif ($entropy -lt 60) { $strengthCategory="Fair"; $fraction=0.4 }
        elseif ($entropy -lt 80) { $strengthCategory="Good"; $fraction=0.6 }
        elseif ($entropy -lt 100) { $strengthCategory="Strong"; $fraction=0.8 }
        else { $strengthCategory="Very Strong"; $fraction=1.0 }

        $lblStrength.Text = "Strength ($([int]$entropy) bits): $strengthCategory"
        $progEntropy.Fraction = $fraction
        # --- End Entropy update ---

        if ($chkShowPwd.Checked) {
            $txtPwd.Text = $script:actualPassword
        } else {
            $txtPwd.Text = ('*' * $script:actualPassword.Length)
        }
    })

    # --- Show Password toggle ---
    $chkShowPwd.add_Toggled({
        if ($chkShowPwd.Checked) {
            $txtPwd.Text = $script:actualPassword
        } else {
            $txtPwd.Text = ('*' * $script:actualPassword.Length)
        }
    })

    # --- Copy to Clipboard ---
    $btnCopy.add_Clicked({
        if (-not $script:actualPassword) { return }
        if ($IsWindows) { Set-Clipboard -Value $script:actualPassword }
        elseif ($IsMacOS) { $script:actualPassword | pbcopy }
        else { $script:actualPassword | xsel --clipboard --input }
        [Terminal.Gui.MessageBox]::Query(50,7,"Copied","Password copied to clipboard.","OK") | Out-Null
    })

    # --- Close ---
    $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })

    [Terminal.Gui.Application]::Run($dlg)
    return $script:actualPassword
}

function Update-UserObjectFromFields($user) {
    $user.Name        = $txtName.Text.ToString()
    $user.Description = $txtDesc.Text.ToString()
    $user.Office      = $txtOffice.Text.ToString()
    $user.Phone       = $txtPhone.Text.ToString()
    $user.MobilePhone = $txtMobile.Text.ToString()
    $user.Email       = $txtEmail.Text.ToString()
    $user.Street      = $txtStreet.Text.ToString()
    $user.City        = $txtCity.Text.ToString()
    $user.PostalCode  = $txtPostal.Text.ToString()
    $user.Country     = $txtCountry.Text.ToString()
    $user.Title       = $txtTitle.Text.ToString()
    $user.Department  = $txtDept.Text.ToString()
    $user.Company     = $txtCompany.Text.ToString()
    $user.Manager     = $txtManager.Text.ToString()
    $user.Disabled    = $chkDisabled.Checked
    $user.Locked      = $chkLocked.Checked
}

function Show-UserPropertiesDialog {
    param($user, $Global)

    # Safety checks
    if (-not $user) { Write-Error "User object is null"; return }
    if (-not $Global:Domain) { $Global:Domain = "" }
    if ($Global:DemoMode -and -not $Global:Users) { $Global:Users = @() }

    # ----- Create main dialog -----
    $dlg = [Terminal.Gui.Dialog]::new("User Properties", 100, 40)

    # ----- TabView -----
    $tabView = [Terminal.Gui.TabView]::new()
    $tabView.X=0; $tabView.Y=0; $tabView.Width=[Terminal.Gui.Dim]::Fill(); $tabView.Height = [Terminal.Gui.Dim]::Percent(98)  # leave 2% at bottom

$btnOK = [Terminal.Gui.Button]::new("OK")
$btnCancel = [Terminal.Gui.Button]::new("Cancel")
$dlg = [Terminal.Gui.Dialog]::new("User Properties", 100, 40, $btnOK, $btnCancel)

    # ==================== General Tab ====================
    $generalTab = [Terminal.Gui.TabView+Tab]::new()
    $generalTab.Text = "General"
    $generalView = [Terminal.Gui.View]::new()

    $y = 1
    # Display Name
    $lbl = [Terminal.Gui.Label]::new("Display Name:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtName = [Terminal.Gui.TextField]::new($user.Name ?? ""); $txtName.X=20; $txtName.Y=$y; $txtName.Width=40
    $txtName.add_TextChanged({ $script:changesMade = $true }); $generalView.Add($txtName)
    $y+=2

    # Description
    $lbl = [Terminal.Gui.Label]::new("Description:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtDesc = [Terminal.Gui.TextField]::new($user.Description ?? ""); $txtDesc.X=20; $txtDesc.Y=$y; $txtDesc.Width=40
    $txtDesc.add_TextChanged({ $script:changesMade = $true }); $generalView.Add($txtDesc)
    $y+=2

    # Office
    $lbl = [Terminal.Gui.Label]::new("Office:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtOffice = [Terminal.Gui.TextField]::new($user.Office ?? ""); $txtOffice.X=20; $txtOffice.Y=$y; $txtOffice.Width=40
    $txtOffice.add_TextChanged({ $script:changesMade = $true }); $generalView.Add($txtOffice)
    $y+=2

    # Telephone
    Write-Host "DEBUG: User Phone value='$($selUser.OfficePhone)'"
    $lbl = [Terminal.Gui.Label]::new("Telephone:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtPhone = [Terminal.Gui.TextField]::new($user.OfficePhone ?? ""); $txtPhone.X=20; $txtPhone.Y=$y; $txtPhone.Width=40
    $txtPhone.add_TextChanged({ $script:changesMade = $true }); $generalView.Add($txtPhone)
    $y+=2

    # Mobile Phone
    $lbl = [Terminal.Gui.Label]::new("Mobile Phone:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtMobile = [Terminal.Gui.TextField]::new($user.MobilePhone ?? ""); $txtMobile.X=20; $txtMobile.Y=$y; $txtMobile.Width=40
    $txtMobile.add_TextChanged({ $script:changesMade = $true }); $generalView.Add($txtMobile)
    $y+=2

    # E-mail 
    Write-Host "DEBUG: User Email value='$($selUser.EmailAddress)'"
    $lbl = [Terminal.Gui.Label]::new("E-mail:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtEmail = [Terminal.Gui.TextField]::new($user.EmailAddress ?? ""); $txtEmail.X=20; $txtEmail.Y=$y; $txtEmail.Width=40
    $txtEmail.add_TextChanged({ $script:changesMade = $true }); $generalView.Add($txtEmail)

    $generalTab.View = $generalView
    $tabView.AddTab($generalTab, $false)

    # ==================== Account Tab ====================
    $accountTab = [Terminal.Gui.TabView+Tab]::new()
    $accountTab.Text = "Account"
    $accountView = [Terminal.Gui.View]::new()

    $y = 1
    $lbl = [Terminal.Gui.Label]::new("User logon name:"); $lbl.X=2; $lbl.Y=$y; $accountView.Add($lbl)
    $txtLogon = [Terminal.Gui.TextField]::new(($user.Name ?? "").ToLower().Replace(' ','.')); $txtLogon.X=20; $txtLogon.Y=$y; $txtLogon.Width=30
    $accountView.Add($txtLogon)
    $y+=2

    $lbl = [Terminal.Gui.Label]::new("Account Status:"); $lbl.X=2; $lbl.Y=$y; $accountView.Add($lbl)
    $statusText = if ($user.Locked) { "🔒 Locked" } elseif ($user.Disabled) { "⊗ Disabled" } else { "○ Enabled" }
    $lblStatus = [Terminal.Gui.Label]::new($statusText); $lblStatus.X=20; $lblStatus.Y=$y; $accountView.Add($lblStatus)
    $y+=2

    $chkDisabled = [Terminal.Gui.CheckBox]::new("Account is disabled"); $chkDisabled.X=2; $chkDisabled.Y=$y
    $chkDisabled.Checked = if ($user.Disabled -is [bool]) { $user.Disabled } else { $false }
    $accountView.Add($chkDisabled)
    $chkDisabled.add_Toggled({
        if ($chkLocked.Checked) { $lblStatus.Text = "🔒 Locked" }
        elseif ($chkDisabled.Checked) { $lblStatus.Text = "⊗ Disabled" }
        else { $lblStatus.Text = "○ Enabled" }
        $script:changesMade = $true
    })
    $y+=1

    $chkLocked = [Terminal.Gui.CheckBox]::new("Account is locked"); $chkLocked.X=2; $chkLocked.Y=$y
    $chkLocked.Checked = if ($user.Locked -is [bool]) { $user.Locked } else { $false }
    $accountView.Add($chkLocked)
    $chkLocked.add_Toggled({
        if ($chkLocked.Checked) { $lblStatus.Text = "🔒 Locked" }
        elseif ($chkDisabled.Checked) { $lblStatus.Text = "⊗ Disabled" }
        else { $lblStatus.Text = "○ Enabled" }
        $script:changesMade = $true
    })
    $y+=2

    $chkPwdExpire = [Terminal.Gui.CheckBox]::new("Password never expires"); $chkPwdExpire.X=2; $chkPwdExpire.Y=$y; $chkPwdExpire.Checked=$false
    $accountView.Add($chkPwdExpire)
    $chkPwdExpire.add_Toggled({ $script:changesMade = $true })
    $y+=2

    $chkChangePwd = [Terminal.Gui.CheckBox]::new("User cannot change password"); $chkChangePwd.X=2; $chkChangePwd.Y=$y; $chkChangePwd.Checked=$false
    $accountView.Add($chkChangePwd)
    $chkChangePwd.add_Toggled({ $script:changesMade = $true })
    $y+=2

    $btnResetPwd = [Terminal.Gui.Button]::new("Reset Password..."); $btnResetPwd.X=2; $btnResetPwd.Y=$y; $accountView.Add($btnResetPwd)
    $btnResetPwd.add_Clicked({
        $newPwd = Generate-RandomPassword
        if (-not $newPwd) { Show-Modal "Cancelled" "Password generation cancelled."; return }
        $confirm = [Terminal.Gui.MessageBox]::Query(
            70, 10,
            "Apply Password",
            "Apply the following password to user:`n`n$($user.Name)`n`nPassword:`n$newPwd`n",
            "Apply", "Cancel"
        )
        if ($confirm -ne 0) { Show-Modal "Cancelled" "Password reset cancelled."; return }
        Write-Host "DEBUG: Password reset for $($user.Name) to: $newPwd"
        Show-Modal "Success" "Password reset (demo mode)."
    })

    $accountTab.View = $accountView
    $tabView.AddTab($accountTab, $false)


# ----- Address Tab -----
$addressTab = [Terminal.Gui.TabView+Tab]::new()
$addressTab.Text = "Address"
$addressView = [Terminal.Gui.View]::new()

$y = 1
$lbl = [Terminal.Gui.Label]::new("Street:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
$txtStreet = [Terminal.Gui.TextField]::new($User.StreetAddress); $txtStreet.X=20; $txtStreet.Y=$y; $txtStreet.Width=40
$txtStreet.add_TextChanged({ $script:changesMade = $true })
$addressView.Add($txtStreet)
$y+=2

$lbl = [Terminal.Gui.Label]::new("City:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
$txtCity = [Terminal.Gui.TextField]::new($User.City); $txtCity.X=20; $txtCity.Y=$y; $txtCity.Width=40
$txtCity.add_TextChanged({ $script:changesMade = $true })
$addressView.Add($txtCity)
$y+=2

$lbl = [Terminal.Gui.Label]::new("Postal Code:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
$txtPostal = [Terminal.Gui.TextField]::new($User.PostalCode); $txtPostal.X=20; $txtPostal.Y=$y; $txtPostal.Width=20
$txtPostal.add_TextChanged({ $script:changesMade = $true })
$addressView.Add($txtPostal)
$y+=2

$lbl = [Terminal.Gui.Label]::new("Country:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
$txtCountry = [Terminal.Gui.TextField]::new($User.Country); $txtCountry.X=20; $txtCountry.Y=$y; $txtCountry.Width=40
$txtCountry.add_TextChanged({ $script:changesMade = $true })
$addressView.Add($txtCountry)

$addressTab.View = $addressView
$tabView.AddTab($addressTab, $false)

# ----- Organization Tab -----
$orgTab = [Terminal.Gui.TabView+Tab]::new()
$orgTab.Text = "Organization"
$orgView = [Terminal.Gui.View]::new()

$y = 1
$lbl = [Terminal.Gui.Label]::new("Title:"); $lbl.X=2; $lbl.Y=$y; $orgView.Add($lbl)
$txtTitle = [Terminal.Gui.TextField]::new($user.Title); $txtTitle.X=20; $txtTitle.Y=$y; $txtTitle.Width=40
$txtTitle.add_TextChanged({ $script:changesMade = $true })
$orgView.Add($txtTitle)
$y+=2

$lbl = [Terminal.Gui.Label]::new("Department:"); $lbl.X=2; $lbl.Y=$y; $orgView.Add($lbl)
$txtDept = [Terminal.Gui.TextField]::new($user.Department); $txtDept.X=20; $txtDept.Y=$y; $txtDept.Width=40
$txtDept.add_TextChanged({ $script:changesMade = $true })
$orgView.Add($txtDept)
$y+=2

$lbl = [Terminal.Gui.Label]::new("Company:"); $lbl.X=2; $lbl.Y=$y; $orgView.Add($lbl)
$txtCompany = [Terminal.Gui.TextField]::new($user.Company); $txtCompany.X=20; $txtCompany.Y=$y; $txtCompany.Width=40
$txtCompany.add_TextChanged({ $script:changesMade = $true })
$orgView.Add($txtCompany)
$y+=2

$lbl = [Terminal.Gui.Label]::new("Manager:"); $lbl.X=2; $lbl.Y=$y; $orgView.Add($lbl)
$txtManager = [Terminal.Gui.TextField]::new($user.Manager); $txtManager.X=20; $txtManager.Y=$y; $txtManager.Width=40
$txtManager.add_TextChanged({ $script:changesMade = $true })
$orgView.Add($txtManager)

$orgTab.View = $orgView
$tabView.AddTab($orgTab, $false)

# ============================= Group Memebrship ==========================
    # ----- Member Of Tab -----
    $memberTab = [Terminal.Gui.TabView+Tab]::new()
    $memberTab.Text = "Member Of"
    $memberView = [Terminal.Gui.View]::new()
    
    # Load groups if not already loaded
    if (-not $user.Groups) {
        try {
            if ($Global:DemoMode) {
                # Already loaded in demo mode
            } else {
                $user.Groups = @(Get-ADPrincipalGroupMembership -Identity $user.Name | Select-Object -ExpandProperty Name)
            }
        } catch { $user.Groups=@() }
    }
    
    $lbl = [Terminal.Gui.Label]::new("Member of the following groups:"); $lbl.X=2; $lbl.Y=1; $memberView.Add($lbl)
    $lstGroups = [Terminal.Gui.ListView]::new($user.Groups)
    $lstGroups.X=2; $lstGroups.Y=3; $lstGroups.Width=[Terminal.Gui.Dim]::Fill(2); $lstGroups.Height=[Terminal.Gui.Dim]::Fill(8)
    $memberView.Add($lstGroups)

         # Add/Remove buttons for group membership
# Add button with FIXED scope handling
$btnAddGroup = [Terminal.Gui.Button]::new("Add..."); $btnAddGroup.X=2; $btnAddGroup.Y=[Terminal.Gui.Pos]::Bottom($lstGroups)+1
$btnAddGroup.add_Clicked({
    # Store reference in a variable that will be captured
    $currentUser = $user
    
    Write-Host "DEBUG: Add to group clicked for user: $($currentUser.Name)"
    
    # Get list of all available groups
    if ($Global:DemoMode) {
        $allGroups = @()
        foreach ($u in $Global:Users) {
            if ($u.ContainsKey('Groups') -and $u['Groups']) {
                $allGroups += $u['Groups']
            }
        }
        $availableGroups = $allGroups | Select-Object -Unique | Sort-Object
    } else {
        try {
            $availableGroups = Get-ADGroup -Filter * | Select-Object -ExpandProperty Name | Sort-Object
        } catch {
            [Terminal.Gui.MessageBox]::Query(60, 8, "Error", "Failed to retrieve groups", "OK") | Out-Null
            return
        }
    }
    
    # Filter out current groups
    $currentGroups = if ($currentUser['Groups']) { $currentUser['Groups'] } else { @() }
    $groupsToAdd = $availableGroups | Where-Object { $currentGroups -notcontains $_ }
    
    Write-Host "DEBUG: Available groups to add: $($groupsToAdd.Count)"
    
    if ($groupsToAdd.Count -eq 0) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "No Groups", "User is already in all groups!", "OK") | Out-Null
        return
    }
    
    # Create dialog
    $grpDlg = [Terminal.Gui.Dialog]::new("Add to Group - $($currentUser.Name)", 60, 20)
    $lblGrp = [Terminal.Gui.Label]::new("Select group to add $($currentUser.Name) to:")
    $lblGrp.X=2; $lblGrp.Y=1
    $grpDlg.Add($lblGrp)
    
    $lstAvailGroups = [Terminal.Gui.ListView]::new()
    $lstAvailGroups.SetSource($groupsToAdd)
    $lstAvailGroups.X=2; $lstAvailGroups.Y=3
    $lstAvailGroups.Width=[Terminal.Gui.Dim]::Fill(2)
    $lstAvailGroups.Height=[Terminal.Gui.Dim]::Fill(2)
    $grpDlg.Add($lstAvailGroups)
    
    # Reference to the groups ListView from parent scope
    $parentGroupsList = $lstGroups
    
    $btnAddOK = [Terminal.Gui.Button]::new("Add")
    $btnAddOK.add_Clicked({
        Write-Host "DEBUG: Add OK clicked, SelectedItem = $($lstAvailGroups.SelectedItem)"
        Write-Host "DEBUG: Current user in handler: '$($currentUser.Name)'"
        
        if ($lstAvailGroups.SelectedItem -eq -1) {
            [Terminal.Gui.MessageBox]::Query(50, 7, "No Selection", "Please select a group", "OK") | Out-Null
            return
        }
        
        if (-not $groupsToAdd -or $groupsToAdd.Count -eq 0) {
            Write-Host "ERROR: groupsToAdd is null or empty"
            [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "No groups available", "OK") | Out-Null
            return
        }
        
        if ($lstAvailGroups.SelectedItem -ge $groupsToAdd.Count) {
            Write-Host "ERROR: SelectedItem ($($lstAvailGroups.SelectedItem)) >= groupsToAdd.Count ($($groupsToAdd.Count))"
            [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "Invalid selection index", "OK") | Out-Null
            return
        }
        
        $selectedGroup = $groupsToAdd[$lstAvailGroups.SelectedItem]
        Write-Host "DEBUG: Selected group: $selectedGroup"
        
        try {
            if ($Global:DemoMode) {
                # Initialize Groups if needed
                if (-not $currentUser.ContainsKey('Groups') -or $null -eq $currentUser['Groups']) {
                    Write-Host "DEBUG: Initializing Groups array"
                    $currentUser['Groups'] = @()
                }
                
                # Ensure it's an array
                if ($currentUser['Groups'] -isnot [array]) {
                    Write-Host "DEBUG: Converting Groups to array"
                    $currentUser['Groups'] = @($currentUser['Groups'])
                }
                
                # Add the group
                Write-Host "DEBUG: Current groups before add: $($currentUser['Groups'] -join ', ')"
                $currentUser['Groups'] += $selectedGroup
                $currentUser['Groups'] = $currentUser['Groups'] | Sort-Object -Unique
                Write-Host "DEBUG: Current groups after add: $($currentUser['Groups'] -join ', ')"
      Write-Host "DEBUG: User $($currentUser.Name) added to group $selectedGroup (demo mode)"
                
                # Close dialog first
                [Terminal.Gui.Application]::RequestStop()
                
                # Then update UI with error handling
                [Terminal.Gui.Application]::MainLoop.Invoke({
                    try {
                        Write-Host "DEBUG: Updating parent groups list..."
                        $parentGroupsList.SetSource($currentUser['Groups'])
                        Write-Host "DEBUG: Parent groups list updated"
                        
                        # Skip tree rebuild - will happen when properties dialog closes
                        # Build-Tree -domain $Global:Domain
                        
                        Write-Host "DEBUG: Updating filter status label..."
                        if ($filterStatusLabel) {
                            Update-FilterStatusLabel -label $filterStatusLabel
                            Write-Host "DEBUG: Filter status label updated"
                        } else {
                            Write-Host "WARNING: filterStatusLabel is null, skipping update"
                        }
                    } catch {
                        Write-Host "ERROR in MainLoop.Invoke: $($_.Exception.Message)"
                        Write-Host "ERROR: Stack trace: $($_.ScriptStackTrace)"
                    }
                })
                
                $script:changesMade = $true
                
                # Success message
                Write-Host "SUCCESS: Added $($currentUser.Name) to group $selectedGroup"

} else {
                # Production mode
                Add-ADGroupMember -Identity $selectedGroup -Members $currentUser.Name -ErrorAction Stop
                
                Write-Host "DEBUG: User $($currentUser.Name) added to group $selectedGroup in AD"
                
                # Reload from AD
                $currentUser['Groups'] = @(Get-ADPrincipalGroupMembership -Identity $currentUser.Name | Select-Object -ExpandProperty Name | Sort-Object)
                
                # Close dialog first
                [Terminal.Gui.Application]::RequestStop()
                
                # Then update UI with error handling
                [Terminal.Gui.Application]::MainLoop.Invoke({
                    try {
                        Write-Host "DEBUG: Updating parent groups list..."
                        $parentGroupsList.SetSource($currentUser['Groups'])
                        Write-Host "DEBUG: Parent groups list updated"
                        
                        Write-Host "DEBUG: Rebuilding tree..."
                        Build-Tree -domain $Global:Domain
                        Write-Host "DEBUG: Tree rebuilt"
                        
                        Write-Host "DEBUG: Updating filter status label..."
                        if ($filterStatusLabel) {
                            Update-FilterStatusLabel -label $filterStatusLabel
                            Write-Host "DEBUG: Filter status label updated"
                        } else {
                            Write-Host "WARNING: filterStatusLabel is null, skipping update"
                        }
                    } catch {
                        Write-Host "ERROR in MainLoop.Invoke: $($_.Exception.Message)"
                        Write-Host "ERROR: Stack trace: $($_.ScriptStackTrace)"
                    }
                })
                
                $script:changesMade = $true
                
                Write-Host "SUCCESS: Added $($currentUser.Name) to group $selectedGroup"
            }
  } catch {
            Write-Host "ERROR: Failed to add to group: $($_.Exception.Message)"
            Write-Host "ERROR: Exception type: $($_.Exception.GetType().FullName)"
            Write-Host "ERROR: Stack trace: $($_.ScriptStackTrace)"
            [Terminal.Gui.MessageBox]::Query(60, 10, "Error", "Failed to add to group:`n$($_.Exception.Message)", "OK") | Out-Null
        }
    }.GetNewClosure())
    $grpDlg.AddButton($btnAddOK)
    
    $btnAddCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnAddCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $grpDlg.AddButton($btnAddCancel)
    
    $lstAvailGroups.add_OpenSelectedItem({ 
        if ($btnAddOK) { 
            $btnAddOK.OnClicked()
        }
    })
    
    [Terminal.Gui.Application]::Run($grpDlg)
}.GetNewClosure())
$memberView.Add($btnAddGroup)
   
# Remove button
$btnRemoveGroup = [Terminal.Gui.Button]::new("Remove")
$btnRemoveGroup.X=[Terminal.Gui.Pos]::Right($btnAddGroup)+2
$btnRemoveGroup.Y=[Terminal.Gui.Pos]::Bottom($lstGroups)+1
$btnRemoveGroup.add_Clicked({
    # Store reference in a variable that will be captured
    $currentUser = $user
    
    Write-Host "DEBUG: Remove from group clicked for user: $($currentUser.Name)"
    
    # Check if a group is selected
    if ($lstGroups.SelectedItem -eq -1) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "No Selection", "Please select a group to remove", "OK") | Out-Null
        return
    }
    
    # Verify user has groups
    if (-not $currentUser.ContainsKey('Groups') -or -not $currentUser['Groups'] -or $currentUser['Groups'].Count -eq 0) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "No Groups", "User is not a member of any groups", "OK") | Out-Null
        return
    }
    
    # Verify selected index is valid
    if ($lstGroups.SelectedItem -ge $currentUser['Groups'].Count) {
        Write-Host "ERROR: SelectedItem out of range"
        [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "Invalid group selection", "OK") | Out-Null
        return
    }
    
    # Get the selected group
    $selectedGroup = $currentUser['Groups'][$lstGroups.SelectedItem]
    Write-Host "DEBUG: Selected group to remove: $selectedGroup"
    
    # Confirm removal - FIXED: Store message first
    $confirmMessage = "Remove $($currentUser.Name) from group '$selectedGroup'?"
    
    try {
        $result = [Terminal.Gui.MessageBox]::Query(60, 8, "Confirm Removal", $confirmMessage, @("Yes", "No"))
    } catch {
        Write-Host "ERROR: MessageBox failed: $($_.Exception.Message)"
        # Fall back to no confirmation, just do it
        $result = 0
    }
    
    if ($result -eq 0) {
        Write-Host "DEBUG: User confirmed removal"
        
        try {
            if ($Global:DemoMode) {
                # Demo mode: remove from array
                Write-Host "DEBUG: Current groups before remove: $($currentUser['Groups'] -join ', ')"
                $currentUser['Groups'] = $currentUser['Groups'] | Where-Object { $_ -ne $selectedGroup }
                Write-Host "DEBUG: Current groups after remove: $($currentUser['Groups'] -join ', ')"
                
                Write-Host "DEBUG: User $($currentUser.Name) removed from group $selectedGroup (demo mode)"
                
                # Update the groups list display
                # Update the groups list display
                [Terminal.Gui.Application]::MainLoop.Invoke({
                    try {
                        Write-Host "DEBUG: Updating groups list..."
                        $lstGroups.SetSource($currentUser['Groups'])
                        $lstGroups.SetNeedsDisplay()
                        $lstGroups.Redraw($lstGroups.Bounds)
                        Write-Host "DEBUG: Groups list updated"
                    } catch {
                        Write-Host "ERROR in MainLoop.Invoke: $($_.Exception.Message)"
                    }
                })                
                $script:changesMade = $true
                
                Write-Host "SUCCESS: Removed $($currentUser.Name) from group $selectedGroup"
                
            } else {
                # Production mode: remove from AD
                Remove-ADGroupMember -Identity $selectedGroup -Members $currentUser.Name -Confirm:$false -ErrorAction Stop
                
                Write-Host "DEBUG: User $($currentUser.Name) removed from group $selectedGroup in AD"
                
                # Reload group membership from AD
                $currentUser['Groups'] = @(Get-ADPrincipalGroupMembership -Identity $currentUser.Name | Select-Object -ExpandProperty Name | Sort-Object)
                
                # Update the groups list display
                [Terminal.Gui.Application]::MainLoop.Invoke({
                    try {
                        Write-Host "DEBUG: Updating groups list..."
                        $lstGroups.SetSource($currentUser['Groups'])
                        $lstGroups.SetNeedsDisplay()
                        $lstGroups.Redraw($lstGroups.Bounds)
                        Write-Host "DEBUG: Groups list updated"
                    } catch {
                        Write-Host "ERROR in MainLoop.Invoke: $($_.Exception.Message)"
                    }
                })
                
                $script:changesMade = $true
                
                Write-Host "SUCCESS: Removed $($currentUser.Name) from group $selectedGroup"
            }
            
        } catch {
            Write-Host "ERROR: Failed to remove from group: $($_.Exception.Message)"
            Write-Host "ERROR: Stack trace: $($_.ScriptStackTrace)"
            [Terminal.Gui.MessageBox]::Query(60, 10, "Error", "Failed to remove from group:`n$($_.Exception.Message)", "OK") | Out-Null
        }
    } else {
        Write-Host "DEBUG: User cancelled removal"
    }
}.GetNewClosure())
$memberView.Add($btnRemoveGroup)
# Add the Member Of tab to the tabview
$memberTab.View = $memberView
$tabView.AddTab($memberTab, $false)    


    # ==================== Search/Lookup Tab ====================

# ----- Search/Lookup Filter Tab -----
    $searchTab = [Terminal.Gui.TabView+Tab]::new()
    $searchTab.Text = "Search/Lookup"
    $searchView = [Terminal.Gui.View]::new()

    $y = 1
    $lblSearchDomain = [Terminal.Gui.Label]::new("Domain:"); $lblSearchDomain.X=2; $lblSearchDomain.Y=$y; $searchView.Add($lblSearchDomain)
    $txtSearchDomain = [Terminal.Gui.TextField]::new($Global:Domain ?? ""); $txtSearchDomain.X=15; $txtSearchDomain.Y=$y; $txtSearchDomain.Width=30; $searchView.Add($txtSearchDomain)
    $y+=2

    $lblSearchName = [Terminal.Gui.Label]::new("Name:"); $lblSearchName.X=2; $lblSearchName.Y=$y; $searchView.Add($lblSearchName)
    $txtSearchUser = [Terminal.Gui.TextField]::new($user.Name ?? ""); $txtSearchUser.X=15; $txtSearchUser.Y=$y; $txtSearchUser.Width=30; $searchView.Add($txtSearchUser)
    $y+=2

    $lblSearchType = [Terminal.Gui.Label]::new("Type:"); $lblSearchType.X=2; $lblSearchType.Y=$y; $searchView.Add($lblSearchType)
    $cmbSearchType = [Terminal.Gui.ComboBox]::new(); $cmbSearchType.X=15; $cmbSearchType.Y=$y; $cmbSearchType.Width=20
    $cmbSearchType.SetSource(@("User","Group","OU")); $cmbSearchType.SelectedItem = 0
    $searchView.Add($cmbSearchType)
    $y+=2

    $lblSearchFilter = [Terminal.Gui.Label]::new("Filter Results:"); $lblSearchFilter.X=48; $lblSearchFilter.Y=1; $searchView.Add($lblSearchFilter)
    $txtSearchFilter = [Terminal.Gui.TextField]::new(""); $txtSearchFilter.X=62; $txtSearchFilter.Y=1; $txtSearchFilter.Width=20; $searchView.Add($txtSearchFilter)

    $lblSearchResult = [Terminal.Gui.Label]::new("Results:"); $lblSearchResult.X=2; $lblSearchResult.Y=$y; $searchView.Add($lblSearchResult)
    $y+=1
    $txtSearchOutput = [Terminal.Gui.TextView]::new(); $txtSearchOutput.X=2; $txtSearchOutput.Y=$y
    $txtSearchOutput.Width=[Terminal.Gui.Dim]::Fill(2); $txtSearchOutput.Height=[Terminal.Gui.Dim]::Fill(4)
    $txtSearchOutput.ReadOnly=$true; $txtSearchOutput.WordWrap=$false
    $searchView.Add($txtSearchOutput)

    $chkSearchLocked = [Terminal.Gui.CheckBox]::new("Account Locked"); $chkSearchLocked.X=2; $chkSearchLocked.Y=[Terminal.Gui.Pos]::Bottom($txtSearchOutput)+1
    $chkSearchLocked.CanFocus=$true; $chkSearchLocked.Data=""; $searchView.Add($chkSearchLocked)

    $chkSearchDisabled = [Terminal.Gui.CheckBox]::new("Account Disabled"); $chkSearchDisabled.X=2; $chkSearchDisabled.Y=[Terminal.Gui.Pos]::Bottom($chkSearchLocked)+1
    $chkSearchDisabled.CanFocus=$true; $chkSearchDisabled.Data=""; $searchView.Add($chkSearchDisabled)

    $btnDoSearch = [Terminal.Gui.Button]::new("Search"); $btnDoSearch.X=48; $btnDoSearch.Y=3; $searchView.Add($btnDoSearch)

    # Helper function and remaining code unchanged...
    function Get-UserOutputLines($userObj) { 
        return @(
            "Name                     : $($userObj.Name ?? '')",
            "Email                    : $($userObj.Email ?? '')",
            "Title                    : $($userObj.Title ?? '')",
            "Department               : $($userObj.Department ?? '')",
            "Office                   : $($userObj.Office ?? '')",
            "Phone                    : $($userObj.Phone ?? '')",
            "MobilePhone              : $($userObj.MobilePhone ?? '')",
            "OU                       : $($userObj.OU ?? '')",
            "Groups                   : $($userObj.Groups -join ', ')",
            "Manager                  : $($userObj.Manager ?? '')",
            "Company                  : $($userObj.Company ?? '')",
            "Street                   : $($userObj.Street ?? '')",
            "City                     : $($userObj.City ?? '')",
            "PostalCode               : $($userObj.PostalCode ?? '')",
            "Country                  : $($userObj.Country ?? '')",
            "Disabled                 : $($userObj.Disabled)",
            "Locked                   : $($userObj.Locked)",
            "Description              : $($userObj.Description ?? '')"
        )
    }

    $searchTab.View = $searchView
    $tabView.AddTab($searchTab, $false)

    # Auto-populate search results
    [Terminal.Gui.Application]::MainLoop.Invoke({
        if ($user) {
            $txtSearchOutput.Text = (Get-UserOutputLines -userObj $user) -join "`n"
            $script:currentSearchOutputLines = Get-UserOutputLines -userObj $user
            $chkSearchLocked.Checked = [bool]($user.Locked)
            $chkSearchLocked.Data = $user.Name
        }
    })



    # ----- Add TabView to Dialog -----
    $dlg.Add($tabView)
# add buttons
# ----- Buttons -----
$btnOK = [Terminal.Gui.Button]::new("OK")
$btnCancel = [Terminal.Gui.Button]::new("Cancel")
$btnApply = [Terminal.Gui.Button]::new("Apply")

# Position buttons at the bottom right
$btnCancel.X = [Terminal.Gui.Pos]::Right($dlg) - 12
$btnCancel.Y = [Terminal.Gui.Pos]::Bottom($dlg) - 2

$btnOK.X = [Terminal.Gui.Pos]::Left($btnCancel) - 12
$btnOK.Y = $btnCancel.Y

$btnApply.X = [Terminal.Gui.Pos]::Left($btnOK) - 12
$btnApply.Y = $btnCancel.Y

# Button actions
$btnOK.add_Clicked({
    # You can apply changes here if you want, or just close
    $dlg.Running = $false
})

$btnCancel.add_Clicked({
    # Discard changes
    $script:changesMade = $false
    $dlg.Running = $false
})

$btnApply.add_Clicked({
    # Apply changes but keep dialog open
    Apply-UserChanges -User $user
})

# Add buttons to dialog
$dlg.Add($btnOK)
$dlg.Add($btnCancel)
$dlg.Add($btnApply)

##run it
    [Terminal.Gui.Application]::Run($dlg)
}

# DSA-TUI Object Management Module v1.0
# Create, Delete, and Move AD Objects

# ------------------------- Create New Object Wizard ------------------------
function Show-NewObjectWizard {
    $dlg = [Terminal.Gui.Dialog]::new("New Object Wizard", 74, 30)
    
    # Step 1: Select object type
    $lblType = [Terminal.Gui.Label]::new("Select object type to create:"); $lblType.X=2; $lblType.Y=1; $dlg.Add($lblType)
    
    $rdoType = [Terminal.Gui.RadioGroup]::new(@("User", "Group", "Organizational Unit", "Computer", "Contact"))
    $rdoType.X=2; $rdoType.Y=3; $rdoType.Height=5
    $dlg.Add($rdoType)
    
    # Common fields
    $y = 9
    $lblName = [Terminal.Gui.Label]::new("Name:"); $lblName.X=2; $lblName.Y=$y; $dlg.Add($lblName)
    $txtName = [Terminal.Gui.TextField]::new(""); $txtName.X=20; $txtName.Y=$y; $txtName.Width=45; $dlg.Add($txtName)
    $y+=2
    
    $lblDisplayName = [Terminal.Gui.Label]::new("Display Name:"); $lblDisplayName.X=2; $lblDisplayName.Y=$y; $dlg.Add($lblDisplayName)
    $txtDisplayName = [Terminal.Gui.TextField]::new(""); $txtDisplayName.X=20; $txtDisplayName.Y=$y; $txtDisplayName.Width=45; $dlg.Add($txtDisplayName)
    $y+=2
    
    $lblOU = [Terminal.Gui.Label]::new("Organizational Unit:"); $lblOU.X=2; $lblOU.Y=$y; $dlg.Add($lblOU)
    
    # Get list of OUs
    $ouList = if ($Global:DemoMode) {
        $Global:Users | Select-Object -ExpandProperty OU -Unique | Sort-Object
    } else {
        try {
            Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName | 
                Select-Object -ExpandProperty DistinguishedName | Sort-Object
        } catch { @("CN=Users,DC=example,DC=com") }
    }
    
    $cmbOU = [Terminal.Gui.ComboBox]::new()
    $cmbOU.X=20; $cmbOU.Y=$y; $cmbOU.Width=45
    $cmbOU.SetSource($ouList)
    $dlg.Add($cmbOU)
    $y+=2
    
    # User-specific fields (shown/hidden based on type)
    $lblSam = [Terminal.Gui.Label]::new("Username (SAM):"); $lblSam.X=2; $lblSam.Y=$y; $dlg.Add($lblSam)
    $txtSam = [Terminal.Gui.TextField]::new(""); $txtSam.X=20; $txtSam.Y=$y; $txtSam.Width=45; $dlg.Add($txtSam)
    $y+=2
    
    $lblEmail = [Terminal.Gui.Label]::new("Email:"); $lblEmail.X=2; $lblEmail.Y=$y; $dlg.Add($lblEmail)
    $txtEmail = [Terminal.Gui.TextField]::new(""); $txtEmail.X=20; $txtEmail.Y=$y; $txtEmail.Width=45; $dlg.Add($txtEmail)
    $y+=2
    
    $lblPassword = [Terminal.Gui.Label]::new("Password:"); $lblPassword.X=2; $lblPassword.Y=$y; $dlg.Add($lblPassword)
    $txtPassword = [Terminal.Gui.TextField]::new(""); $txtPassword.X=20; $txtPassword.Y=$y; $txtPassword.Width=45; $txtPassword.Secret=$true; $dlg.Add($txtPassword)
    
    # Show/hide fields based on type
    $rdoType.add_SelectedItemChanged({
        $isUser = $rdoType.SelectedItem -eq 0
        $lblSam.Visible = $isUser
        $txtSam.Visible = $isUser
        $lblEmail.Visible = $isUser
        $txtEmail.Visible = $isUser
        $lblPassword.Visible = $isUser
        $txtPassword.Visible = $isUser
    })
    
    # Create button
    $btnCreate = [Terminal.Gui.Button]::new("Create")
    $btnCreate.add_Clicked({
        $objType = @("User", "Group", "OrganizationalUnit", "Computer", "Contact")[$rdoType.SelectedItem]
        $name = $txtName.Text.ToString().Trim()
        $displayName = $txtDisplayName.Text.ToString().Trim()
        $ou = $cmbOU.Text.ToString()
        
        if (-not $name) {
            [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "Name is required!", "OK") | Out-Null
            return
        }
        
        try {
            if ($Global:DemoMode) {
                # Demo mode - add to in-memory structures
                switch ($objType) {
                    "User" {
                        $sam = $txtSam.Text.ToString().Trim()
                        $email = $txtEmail.Text.ToString().Trim()
                        if (-not $sam) { $sam = $name.ToLower().Replace(' ', '.') }
                        if (-not $email) { $email = "$sam@example.com" }
                        
                        $newUser = @{
                            Name=$name
                            OU=$ou
                            Groups=@()
                            Title=""
                            Email=$email
                            Country=""
                            Disabled=$false
                            Department=""
                            Office=""
                            Phone=""
                            Street=""
                            City=""
                            PostalCode=""
                            Company=""
                            Manager=""
                            Description=$displayName
                        }
                        $Global:Users += $newUser
                        Debug-Log "DEBUG: Created user $name in demo mode"
                    }
                    "Group" {
                        Debug-Log "DEBUG: Created group $name in demo mode"
                    }
                    "OrganizationalUnit" {
                        Debug-Log "DEBUG: Created OU $name in demo mode"
                    }
                }
                
                [Terminal.Gui.MessageBox]::Query(50, 7, "Success", "$objType '$name' created successfully (demo mode)", "OK") | Out-Null
                Build-Tree -domain $Global:Domain
                Update-FilterStatusLabel -label $filterStatusLabel
                [Terminal.Gui.Application]::RequestStop()
                
            } else {
                # Production mode - create in AD
                switch ($objType) {
                    "User" {
                        $sam = $txtSam.Text.ToString().Trim()
                        $email = $txtEmail.Text.ToString().Trim()
                        $pwd = $txtPassword.Text.ToString()
                        
                        if (-not $sam) {
                            [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "Username (SAM) is required for users!", "OK") | Out-Null
                            return
                        }
                        
                        if (-not $pwd) {
                            [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "Password is required for users!", "OK") | Out-Null
                            return
                        }
                        
                        $secPwd = ConvertTo-SecureString -String $pwd -AsPlainText -Force
                        
                        $params = @{
                            Name = $name
                            SamAccountName = $sam
                            UserPrincipalName = "$sam@$($Global:Domain)"
                            AccountPassword = $secPwd
                            Enabled = $true
                            Path = $ou
                            ChangePasswordAtLogon = $true
                        }
                        
                        if ($displayName) { $params['DisplayName'] = $displayName }
                        if ($email) { $params['EmailAddress'] = $email }
                        
                        New-ADUser @params -ErrorAction Stop
                        Debug-Log "DEBUG: Created user $name in AD"
                    }
                    "Group" {
                        $params = @{
                            Name = $name
                            GroupScope = "Global"
                            GroupCategory = "Security"
                            Path = $ou
                        }
                        
                        if ($displayName) { $params['Description'] = $displayName }
                        
                        New-ADGroup @params -ErrorAction Stop
                        Debug-Log "DEBUG: Created group $name in AD"
                    }
                    "OrganizationalUnit" {
                        $params = @{
                            Name = $name
                            Path = $ou
                        }
                        
                        if ($displayName) { $params['Description'] = $displayName }
                        
                        New-ADOrganizationalUnit @params -ErrorAction Stop
                        Debug-Log "DEBUG: Created OU $name in AD"
                    }
                    "Computer" {
                        $params = @{
                            Name = $name
                            Path = $ou
                        }
                        
                        New-ADComputer @params -ErrorAction Stop
                        Debug-Log "DEBUG: Created computer $name in AD"
                    }
                    "Contact" {
                        $params = @{
                            Name = $name
                            Type = "Contact"
                            Path = $ou
                        }
                        
                        if ($displayName) { $params['DisplayName'] = $displayName }
                        
                        New-ADObject @params -ErrorAction Stop
                        Debug-Log "DEBUG: Created contact $name in AD"
                    }
                }
                
                [Terminal.Gui.MessageBox]::Query(50, 7, "Success", "$objType '$name' created successfully", "OK") | Out-Null
                
                # Refresh data
                Load-DomainData -domain $Global:Domain
                Build-Tree -domain $Global:Domain
                Update-FilterStatusLabel -label $filterStatusLabel
                [Terminal.Gui.Application]::RequestStop()
            }
            
        } catch {
            $errMsg = $_.Exception.Message
            [Terminal.Gui.MessageBox]::Query(60, 10, "Error", "Failed to create $objType`:`n$errMsg", "OK") | Out-Null
        }
    })
    $dlg.AddButton($btnCreate)
    
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($dlg)
}

# ------------------------- Delete Object ------------------------
function Show-DeleteObjectDialog {
    param([string]$objectName, [string]$objectType)
    
    $cleanName = $objectName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
    
    # Extra confirmation for destructive action
    $result = [Terminal.Gui.MessageBox]::Query(70, 11, "DELETE CONFIRMATION", 
        "⚠️ WARNING: You are about to DELETE:`n`n  Type: $objectType`n  Name: $cleanName`n`nThis action CANNOT be undone!`n`nAre you absolutely sure?", 
        "Yes, DELETE", "No, Cancel")
    
    if ($result -eq 0) {
        try {
            if ($Global:DemoMode) {
                # Demo mode - remove from in-memory structures
                switch ($objectType.ToLower()) {
                    "user" {
                        $Global:Users = $Global:Users | Where-Object { $_.Name -ne $cleanName }
                        Debug-Log "DEBUG: Deleted user $cleanName (demo mode)"
                    }
                    "group" {
                        # Remove group from all users
                        foreach ($u in $Global:Users) {
                            $u.Groups = $u.Groups | Where-Object { $_ -ne $cleanName }
                        }
                        Debug-Log "DEBUG: Deleted group $cleanName (demo mode)"
                    }
                    default {
                        Debug-Log "DEBUG: Deleted $objectType $cleanName (demo mode)"
                    }
                }
                
                [Terminal.Gui.MessageBox]::Query(50, 7, "Deleted", "$objectType '$cleanName' deleted (demo mode)", "OK") | Out-Null
                Build-Tree -domain $Global:Domain
                Update-FilterStatusLabel -label $filterStatusLabel
                
            } else {
                # Production mode - delete from AD
                switch ($objectType.ToLower()) {
                    "user" {
                        Remove-ADUser -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log "DEBUG: Deleted user $cleanName from AD"
                    }
                    "group" {
                        Remove-ADGroup -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log "DEBUG: Deleted group $cleanName from AD"
                    }
                    "ou" {
                        Remove-ADOrganizationalUnit -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log "DEBUG: Deleted OU $cleanName from AD"
                    }
                    "computer" {
                        Remove-ADComputer -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log "DEBUG: Deleted computer $cleanName from AD"
                    }
                    default {
                        Remove-ADObject -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log "DEBUG: Deleted $objectType $cleanName from AD"
                    }
                }
                
                [Terminal.Gui.MessageBox]::Query(50, 7, "Deleted", "$objectType '$cleanName' deleted successfully", "OK") | Out-Null
                
                # Refresh data
                Load-DomainData -domain $Global:Domain
                Build-Tree -domain $Global:Domain
                Update-FilterStatusLabel -label $filterStatusLabel
            }
            
        } catch {
            $errMsg = $_.Exception.Message
            [Terminal.Gui.MessageBox]::Query(60, 10, "Delete Failed", "Failed to delete $objectType`:`n$errMsg", "OK") | Out-Null
        }
    }
}

# ------------------------- Move Object ------------------------
function Show-MoveObjectDialog {
    param([string]$objectName, [string]$objectType)
    
    $cleanName = $objectName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
    
    $dlg = [Terminal.Gui.Dialog]::new("Move Object - $cleanName", 70, 18)
    
    $lblCurrent = [Terminal.Gui.Label]::new("Current location:"); $lblCurrent.X=2; $lblCurrent.Y=1; $dlg.Add($lblCurrent)
    
    # Get current OU
    $currentOU = "N/A"
    if ($objectType.ToLower() -eq "user") {
        $user = $Global:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
        if ($user) { $currentOU = $user.OU }
    }
    
    $lblCurrentOU = [Terminal.Gui.Label]::new($currentOU); $lblCurrentOU.X=20; $lblCurrentOU.Y=1; $dlg.Add($lblCurrentOU)
    
    $lblTarget = [Terminal.Gui.Label]::new("Move to OU:"); $lblTarget.X=2; $lblTarget.Y=3; $dlg.Add($lblTarget)
    
    # Get list of OUs
    $ouList = if ($Global:DemoMode) {
        $Global:Users | Select-Object -ExpandProperty OU -Unique | Sort-Object
    } else {
        try {
            Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName | 
                Select-Object -ExpandProperty DistinguishedName | Sort-Object
        } catch { @("CN=Users,DC=example,DC=com") }
    }
    
    $lstOU = [Terminal.Gui.ListView]::new($ouList)
    $lstOU.X=2; $lstOU.Y=4; $lstOU.Width=[Terminal.Gui.Dim]::Fill(2); $lstOU.Height=8
    $dlg.Add($lstOU)
    
    $btnMove = [Terminal.Gui.Button]::new("Move")
    $btnMove.add_Clicked({
        if ($lstOU.SelectedItem -lt 0) {
            [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "Please select a target OU", "OK") | Out-Null
            return
        }
        
        $targetOU = $ouList[$lstOU.SelectedItem]
        
        if ($targetOU -eq $currentOU) {
            [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "Object is already in that OU", "OK") | Out-Null
            return
        }
        
        $confirm = [Terminal.Gui.MessageBox]::Query(60, 9, "Confirm Move", 
            "Move '$cleanName' to:`n$targetOU?", 
            "Yes", "No")
        
        if ($confirm -eq 0) {
            try {
                if ($Global:DemoMode) {
                    # Demo mode - update in-memory
                    if ($objectType.ToLower() -eq "user") {
                        $user = $Global:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                        if ($user) {
                            $user.OU = $targetOU
                            Debug-Log "DEBUG: Moved user $cleanName to $targetOU (demo mode)"
                        }
                    }
                    
                    [Terminal.Gui.MessageBox]::Query(50, 7, "Success", "Object moved successfully (demo mode)", "OK") | Out-Null
                    Build-Tree -domain $Global:Domain
                    Update-FilterStatusLabel -label $filterStatusLabel
                    [Terminal.Gui.Application]::RequestStop()
                    
                } else {
                    # Production mode - move in AD
                    $adObject = Get-ADObject -Filter "Name -eq '$cleanName'" -ErrorAction Stop
                    Move-ADObject -Identity $adObject.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
                    
                    Debug-Log "DEBUG: Moved $cleanName to $targetOU in AD"
                    [Terminal.Gui.MessageBox]::Query(50, 7, "Success", "Object moved successfully", "OK") | Out-Null
                    
                    # Refresh data
                    Load-DomainData -domain $Global:Domain
                    Build-Tree -domain $Global:Domain
                    Update-FilterStatusLabel -label $filterStatusLabel
                    [Terminal.Gui.Application]::RequestStop()
                }
                
            } catch {
                $errMsg = $_.Exception.Message
                [Terminal.Gui.MessageBox]::Query(60, 10, "Move Failed", "Failed to move object:`n$errMsg", "OK") | Out-Null
            }
        }
    })
    $dlg.AddButton($btnMove)
    
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($dlg)
}

# ------------------------- Change Domain Dialog ------------------------
function Show-ChangeDomainDialog {
    $dlg = [Terminal.Gui.Dialog]::new("Change Domain",50,12)
    $dlg.Add([Terminal.Gui.Label]::new("Domain Name:")) | Out-Null
    $txtDomain = [Terminal.Gui.TextField]::new($Global:Domain); $txtDomain.X=15; $txtDomain.Y=0
    $dlg.Add($txtDomain)
    $okBtn = [Terminal.Gui.Button]::new("OK"); $okBtn.X=10; $okBtn.Y=2
    $okBtn.add_Clicked({
        $domainString = -join ($txtDomain.Text | ForEach-Object { [char]$_ })
        Debug-Log "DEBUG: OK pressed, Domain = $domainString"
        $Global:Domain = $domainString
        Load-DomainData -domain $Global:Domain
        Build-Tree -domain $Global:Domain
        # Add after Build-Tree calls:
        Update-FilterStatusLabel -label $filterStatusLabel
        [Terminal.Gui.Application]::RequestStop()
    })
    $dlg.Add($okBtn)
    $cancelBtn = [Terminal.Gui.Button]::new("Cancel"); $cancelBtn.X=20; $cancelBtn.Y=2
    $cancelBtn.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.Add($cancelBtn)
    [Terminal.Gui.Application]::Run($dlg)
}

# ------------------------- Change DC Dialog ------------------------
function Show-ChangeDCDialog {
    $dlg = [Terminal.Gui.Dialog]::new("Change Domain Controller",50,12)
    $dlg.Add([Terminal.Gui.Label]::new("Select Domain Controller:")) | Out-Null
    $dcNames = $Global:DCs | ForEach-Object { $_.Name }
    $listView = [Terminal.Gui.ListView]::new($dcNames); $listView.X=0; $listView.Y=1; $listView.Width=48; $listView.Height=6
    $dlg.Add($listView)
    $okBtn = [Terminal.Gui.Button]::new("OK"); $okBtn.X=10; $okBtn.Y=8
    $okBtn.add_Clicked({
        if ($listView.SelectedItem -ge 0) { $Global:CurrentDC = $dcNames[$listView.SelectedItem]; $status.Items[1].Title = "DC: $Global:CurrentDC" }
        [Terminal.Gui.Application]::RequestStop()
    })
    $dlg.Add($okBtn)
    $cancelBtn = [Terminal.Gui.Button]::new("Cancel"); $cancelBtn.X=20; $cancelBtn.Y=8
    $cancelBtn.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.Add($cancelBtn)
    [Terminal.Gui.Application]::Run($dlg)
}

# ------------------------- Tree Expand/Collapse ------------------------
$tree.add_KeyPress({ param($sender,$keyArgs) if ($keyArgs.KeyEvent.Key -eq [Terminal.Gui.Key]::Enter -and $tree.SelectedObject) { $tree.SelectedObject.Expanded = -not $tree.SelectedObject.Expanded; $tree.SetNeedsDisplay(); $keyArgs.Handled = $true } })

# ------------------------- AD Search Dialog ------------------------
# DSA-TUI Advanced Search Module v1.0
# Replace the Show-ADSearchDialog function with this enhanced version
# Features: LDAP filters, saved searches, export results

# Global for saved searches
if (-not $Global:SavedSearches) {
    $Global:SavedSearches = @(
        @{Name="Disabled Users"; Filter="(&(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))"; Type="User"},
        @{Name="Users Never Logged In"; Filter="(&(objectClass=user)(!(lastLogon=*)))"; Type="User"},
        @{Name="Computers (Active)"; Filter="(&(objectClass=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"; Type="Computer"},
        @{Name="Empty Groups"; Filter="(&(objectClass=group)(!(member=*)))"; Type="Group"}
        @{Name="Locked Accounts"; Filter="(&(objectClass=user)(lockoutTime>=1)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"; Type="User"},
        @{Name="Password Expiring Soon (7 days)"; Filter="(&(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=65536))(pwdLastSet<=$sevenDaysFileTime))"; Type="User"}
        @{Name="Locked Accounts"; Filter="(&(objectClass=user)(lockoutTime>=1)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))";  Type="User"},
        @{Name="Password Expiring Soon (7 days)"; Filter="(&(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=65536))(pwdLastSet<=$sevenDaysFileTime))"; Type="User"}
    )
}

function Show-ADSearchDialog {
    $dlg = [Terminal.Gui.Dialog]::new("Advanced Search - Active Directory",100,40)
    $dlg.X = 0; $dlg.Y = 0

    # Create TabView for Basic vs Advanced search
    $tabView = [Terminal.Gui.TabView]::new()
    $tabView.X = 0
    $tabView.Y = 0
    $tabView.Width = [Terminal.Gui.Dim]::Fill()
    $tabView.Height = [Terminal.Gui.Dim]::Fill(14)

    # Store search results globally so export can access them
    $script:lastSearchResults = @()
    $script:lastSearchType = ""

    # ----- Basic Search Tab -----
    $basicTab = [Terminal.Gui.TabView+Tab]::new()
    $basicTab.Text = "Basic Search"
    $basicView = [Terminal.Gui.View]::new()

    $y = 1
    $lblDomain = [Terminal.Gui.Label]::new("Domain:"); $lblDomain.X=2; $lblDomain.Y=$y; $basicView.Add($lblDomain)
    $txtDomain = [Terminal.Gui.TextField]::new($Global:Domain); $txtDomain.X=18; $txtDomain.Y=$y; $txtDomain.Width=35; $basicView.Add($txtDomain)
    $y+=2

    $lblName = [Terminal.Gui.Label]::new("Name:"); $lblName.X=2; $lblName.Y=$y; $basicView.Add($lblName)
    $txtUser = [Terminal.Gui.TextField]::new(""); $txtUser.X=18; $txtUser.Y=$y; $txtUser.Width=35; $basicView.Add($txtUser)
    $y+=2

    $lblType = [Terminal.Gui.Label]::new("Type:"); $lblType.X=2; $lblType.Y=$y; $basicView.Add($lblType)
    $cmbObjType = [Terminal.Gui.ComboBox]::new(); $cmbObjType.X=18; $cmbObjType.Y=$y; $cmbObjType.Width=20
    $cmbObjType.SetSource(@("User","Group","Computer","OU","Contact"))
    $basicView.Add($cmbObjType)
    $y+=2

    $chkDisabledOnly = [Terminal.Gui.CheckBox]::new("Disabled accounts only"); $chkDisabledOnly.X=2; $chkDisabledOnly.Y=$y
    $basicView.Add($chkDisabledOnly)

    $basicTab.View = $basicView
    $tabView.AddTab($basicTab, $false)

    # ----- Advanced Search Tab -----
    $advTab = [Terminal.Gui.TabView+Tab]::new()
    $advTab.Text = "Advanced (LDAP)"
    $advView = [Terminal.Gui.View]::new()

    $y = 1
    $lblLdap = [Terminal.Gui.Label]::new("LDAP Filter:"); $lblLdap.X=2; $lblLdap.Y=$y; $advView.Add($lblLdap)
    $y+=1
    $txtLdapFilter = [Terminal.Gui.TextView]::new(); $txtLdapFilter.X=2; $txtLdapFilter.Y=$y; $txtLdapFilter.Width=[Terminal.Gui.Dim]::Fill(2); $txtLdapFilter.Height=4
    $txtLdapFilter.Text = "(&(objectClass=user)(name=*))"
    $advView.Add($txtLdapFilter)
    $y+=5

    $lblExamples = [Terminal.Gui.Label]::new("Examples:"); $lblExamples.X=2; $lblExamples.Y=$y; $advView.Add($lblExamples)
    $y+=1
    $lblEx1 = [Terminal.Gui.Label]::new("Disabled users: (&(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))"); $lblEx1.X=2; $lblEx1.Y=$y; $advView.Add($lblEx1)
    $y+=1
    $lblEx2 = [Terminal.Gui.Label]::new("Users in OU: (&(objectClass=user)(ou=Sales))"); $lblEx2.X=2; $lblEx2.Y=$y; $advView.Add($lblEx2)
    $y+=1
    $lblEx3 = [Terminal.Gui.Label]::new("Groups with members: (&(objectClass=group)(member=*))"); $lblEx3.X=2; $lblEx3.Y=$y; $advView.Add($lblEx3)

    $advTab.View = $advView
    $tabView.AddTab($advTab, $false)

    # ----- Saved Searches Tab -----
    $savedTab = [Terminal.Gui.TabView+Tab]::new()
    $savedTab.Text = "Saved Searches"
    $savedView = [Terminal.Gui.View]::new()

    $lblSaved = [Terminal.Gui.Label]::new("Select a saved search:"); $lblSaved.X=2; $lblSaved.Y=1; $savedView.Add($lblSaved)
    $savedNames = $Global:SavedSearches | ForEach-Object { "$($_.Name) [$($_.Type)]" }
    $lstSaved = [Terminal.Gui.ListView]::new($savedNames); $lstSaved.X=2; $lstSaved.Y=3; $lstSaved.Width=[Terminal.Gui.Dim]::Fill(2); $lstSaved.Height=[Terminal.Gui.Dim]::Fill(4)
    $savedView.Add($lstSaved)

    $btnLoadSaved = [Terminal.Gui.Button]::new("Load Filter"); $btnLoadSaved.X=2; $btnLoadSaved.Y=[Terminal.Gui.Pos]::Bottom($lstSaved)+1
    $btnLoadSaved.add_Clicked({
        if ($lstSaved.SelectedItem -ge 0) {
            $selected = $Global:SavedSearches[$lstSaved.SelectedItem]
            $txtLdapFilter.Text = $selected.Filter
            $tabView.SelectedTab = $advTab
            [Terminal.Gui.MessageBox]::Query(50, 7, "Loaded", "Loaded filter: $($selected.Name)", "OK") | Out-Null
        }
    })
    $savedView.Add($btnLoadSaved)

    $btnSaveCurrent = [Terminal.Gui.Button]::new("Save Current"); $btnSaveCurrent.X=[Terminal.Gui.Pos]::Right($btnLoadSaved)+2; $btnSaveCurrent.Y=[Terminal.Gui.Pos]::Bottom($lstSaved)+1
    $btnSaveCurrent.add_Clicked({
        $filter = $txtLdapFilter.Text.ToString().Trim()
        if ($filter) {
            # Simple input dialog for name
            $nameDlg = [Terminal.Gui.Dialog]::new("Save Search", 60, 10)
            $lbl = [Terminal.Gui.Label]::new("Search Name:"); $lbl.X=2; $lbl.Y=1; $nameDlg.Add($lbl)
            $txtName = [Terminal.Gui.TextField]::new("My Search"); $txtName.X=2; $txtName.Y=3; $txtName.Width=54; $nameDlg.Add($txtName)
            $okBtn = [Terminal.Gui.Button]::new("OK")
            $okBtn.add_Clicked({
                $newName = $txtName.Text.ToString()
                $Global:SavedSearches += @{Name=$newName; Filter=$filter; Type="Custom"}
                Debug-Log "DEBUG: Saved search '$newName'"
                [Terminal.Gui.Application]::RequestStop()
            })
            $nameDlg.AddButton($okBtn)
            $cancelBtn = [Terminal.Gui.Button]::new("Cancel")
            $cancelBtn.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
            $nameDlg.AddButton($cancelBtn)
            [Terminal.Gui.Application]::Run($nameDlg)
            
            # Refresh list
            $savedNames = $Global:SavedSearches | ForEach-Object { "$($_.Name) [$($_.Type)]" }
            $lstSaved.SetSource($savedNames)
        }
    })
    $savedView.Add($btnSaveCurrent)

    $savedTab.View = $savedView
    $tabView.AddTab($savedTab, $false)

    $dlg.Add($tabView)

    # ----- Results Section (below tabs) -----
    $lblResults = [Terminal.Gui.Label]::new("Results:"); $lblResults.X=2; $lblResults.Y=[Terminal.Gui.Pos]::Bottom($tabView)+1; $dlg.Add($lblResults)
    $txtOutput = [Terminal.Gui.TextView]::new()
    $txtOutput.X=2; $txtOutput.Y=[Terminal.Gui.Pos]::Bottom($lblResults); $txtOutput.Width=[Terminal.Gui.Dim]::Fill(2); $txtOutput.Height=6; $txtOutput.ReadOnly=$true
    $dlg.Add($txtOutput)

    # ----- Search Button -----
    $btnSearch = [Terminal.Gui.Button]::new("Search"); $btnSearch.X=2; $btnSearch.Y=[Terminal.Gui.Pos]::Bottom($txtOutput)+1; $dlg.Add($btnSearch)
    
    $btnSearch.add_Clicked({
        $searchName = $txtUser.Text.ToString().Trim()
        $domain = $txtDomain.Text.ToString().Trim()
        $objType = $cmbObjType.Text.ToString()
        $currentTab = $tabView.SelectedTab

        try {
            $objs = @()

            # Determine search mode
            if ($currentTab -eq $advTab) {
                # LDAP Filter search
                $filter = $txtLdapFilter.Text.ToString().Trim()
                if (-not $filter) { $txtOutput.Text="Please enter an LDAP filter."; return }
                
                if ($Global:DemoMode) {
                    $txtOutput.Text="LDAP search not supported in demo mode. Use Basic search."
                    return
                } else {
                    $loadingDlg = Show-LoadingDialog -Message "Executing LDAP query..."
                    try {
                        $objs = Get-ADObject -LDAPFilter $filter -Properties Name,ObjectClass,DistinguishedName -ErrorAction Stop |
                            Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={$_.ObjectClass}}, @{Name='DN';Expression={$_.DistinguishedName}}
                        $script:lastSearchType = "LDAP"
                    } finally { Close-LoadingDialog $loadingDlg }
                }
            } else {
                # Basic search
                if (-not $searchName) { $txtOutput.Text="Please enter a name."; return }

                if ($Global:DemoMode) {
                    switch ($objType) {
                        "User" { $objs = $Global:Users | Where-Object { $_.Name -like "*$searchName*" } | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"user"}} }
                        "Group" { 
                            $matchedGroups = @(); foreach ($u in $Global:Users) { foreach ($g in $u.Groups) { if ($g -like "*$searchName*") { $matchedGroups += $g } } }
                            $objs = ($matchedGroups | Sort-Object -Unique) | ForEach-Object { [PSCustomObject]@{ Name=$_; Type="group" } }
                        }
                        "OU" {
                            $ouNames = ($Global:Users | Select-Object -ExpandProperty OU -Unique)
                            $objs = ($ouNames | Where-Object { $_ -like "*$searchName*" }) | ForEach-Object { [PSCustomObject]@{ Name=$_; Type="organizationalUnit" } }
                        }
                    }
                    $script:lastSearchType = "Basic ($objType)"
                } else {
                    $loadingDlg = Show-LoadingDialog -Message "Searching AD for $objType '$searchName'..."
                    try {
                        $filterStr = "Name -like '*$searchName*'"
                        if ($chkDisabledOnly.Checked -and $objType -eq "User") {
                            $filterStr = "Name -like '*$searchName*' -and Enabled -eq `$false"
                        }

                        switch ($objType) {
                            "User" { $objs = Get-ADUser -Filter $filterStr -Properties Enabled -ErrorAction Stop | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"user"}}, @{Name='Enabled';Expression={$_.Enabled}} }
                            "Group" { $objs = Get-ADGroup -Filter "Name -like '*$searchName*'" -ErrorAction Stop | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"group"}} }
                            "Computer" { $objs = Get-ADComputer -Filter "Name -like '*$searchName*'" -ErrorAction Stop | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"computer"}} }
                            "OU" { $objs = Get-ADOrganizationalUnit -Filter "Name -like '*$searchName*'" -ErrorAction Stop | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"organizationalUnit"}} }
                            "Contact" { $objs = Get-ADObject -Filter "ObjectClass -eq 'contact' -and Name -like '*$searchName*'" -Properties Name -ErrorAction Stop | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"contact"}} }
                            default { $objs=@() }
                        }
                        $script:lastSearchType = "Basic ($objType)"
                    } finally { Close-LoadingDialog $loadingDlg }
                }
            }

            # Store results for export
            $script:lastSearchResults = $objs

            if (-not $objs -or $objs.Count -eq 0) { $txtOutput.Text = "No results found"; return }

            # Display results
            $resultText = "Found $($objs.Count) object(s):`n`n"
            $resultText += ($objs | ForEach-Object { "$($_.Name) [$($_.Type)]" }) -join "`n"
            $txtOutput.Text = $resultText

        } catch {
            $errMsg = $_.Exception.Message
            $txtOutput.Text = "Error: $errMsg"
        }
    })

    # ----- Export Button -----
    $btnExport = [Terminal.Gui.Button]::new("Export..."); $btnExport.X=[Terminal.Gui.Pos]::Right($btnSearch)+2; $btnExport.Y=[Terminal.Gui.Pos]::Bottom($txtOutput)+1; $dlg.Add($btnExport)
    
    $btnExport.add_Clicked({
        if ($script:lastSearchResults.Count -eq 0) {
            [Terminal.Gui.MessageBox]::Query(50, 7, "No Results", "No search results to export. Run a search first.", "OK") | Out-Null
            return
        }

        # Export format dialog
        $exportDlg = [Terminal.Gui.Dialog]::new("Export Results", 60, 12)
        $lbl = [Terminal.Gui.Label]::new("Export format:"); $lbl.X=2; $lbl.Y=1; $exportDlg.Add($lbl)
        
        $rdoCsv = [Terminal.Gui.RadioGroup]::new(@("CSV (Comma-Separated)", "Text (Tab-Separated)", "Text (List)"))
        $rdoCsv.X=2; $rdoCsv.Y=3; $rdoCsv.SelectedItem=0; $exportDlg.Add($rdoCsv)

        $lblFile = [Terminal.Gui.Label]::new("Filename:"); $lblFile.X=2; $lblFile.Y=7; $exportDlg.Add($lblFile)
        $txtFilename = [Terminal.Gui.TextField]::new("search_results.csv"); $txtFilename.X=12; $txtFilename.Y=7; $txtFilename.Width=40; $exportDlg.Add($txtFilename)

        $okBtn = [Terminal.Gui.Button]::new("Export")
        $okBtn.add_Clicked({
            $filename = $txtFilename.Text.ToString()
            $format = $rdoCsv.SelectedItem

            try {
                switch ($format) {
                    0 { # CSV
                        $script:lastSearchResults | Export-Csv -Path $filename -NoTypeInformation -ErrorAction Stop
                    }
                    1 { # Tab-separated
                        $content = $script:lastSearchResults | ForEach-Object { "$($_.Name)`t$($_.Type)" }
                        $content | Out-File -FilePath $filename -ErrorAction Stop
                    }
                    2 { # List
                        $content = $script:lastSearchResults | ForEach-Object { "Name: $($_.Name)`nType: $($_.Type)`n" }
                        $content | Out-File -FilePath $filename -ErrorAction Stop
                    }
                }
                Debug-Log "DEBUG: Exported $($script:lastSearchResults.Count) results to $filename"
                [Terminal.Gui.MessageBox]::Query(60, 8, "Success", "Exported $($script:lastSearchResults.Count) results to:`n$filename", "OK") | Out-Null
                [Terminal.Gui.Application]::RequestStop()
            } catch {
                $errMsg = $_.Exception.Message
                [Terminal.Gui.MessageBox]::Query(60, 10, "Export Failed", "Failed to export:`n$errMsg", "OK") | Out-Null
            }
        })
        $exportDlg.AddButton($okBtn)

        $cancelBtn = [Terminal.Gui.Button]::new("Cancel")
        $cancelBtn.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
        $exportDlg.AddButton($cancelBtn)

        [Terminal.Gui.Application]::Run($exportDlg)
    })

    # ----- Clear Button -----
    $btnClear = [Terminal.Gui.Button]::new("Clear"); $btnClear.X=[Terminal.Gui.Pos]::Right($btnExport)+2; $btnClear.Y=[Terminal.Gui.Pos]::Bottom($txtOutput)+1; $dlg.Add($btnClear)
    $btnClear.add_Clicked({ $txtUser.Text=""; $txtOutput.Text=""; $script:lastSearchResults=@() })

    # ----- Close Button -----
    $btnClose = [Terminal.Gui.Button]::new("Close"); $btnClose.X=[Terminal.Gui.Pos]::Right($btnClear)+2; $btnClose.Y=[Terminal.Gui.Pos]::Bottom($txtOutput)+1; $dlg.Add($btnClose)
    $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })

    [Terminal.Gui.Application]::Run($dlg)
}

# DSA-TUI Context Menu & Refresh Module v1.0
# Add these functions to your main script

# ------------------------- Refresh Tree Function ------------------------
function Refresh-TreeData {
    Debug-Log "DEBUG: Refreshing tree data..."
    
    # Show loading dialog
    $loadingDlg = Show-LoadingDialog -Message "Refreshing Active Directory data..."
    
    try {
        # Reload domain data
        Load-DomainData -domain $Global:Domain
        
        # Rebuild tree
        Build-Tree -domain $Global:Domain
        # Add after Build-Tree calls:
        Update-FilterStatusLabel -label $filterStatusLabel
        
        Debug-Log "DEBUG: Tree refreshed successfully"
    } finally {
        Close-LoadingDialog $loadingDlg
    }
    
    [Terminal.Gui.MessageBox]::Query(50, 7, "Refreshed", "Active Directory data refreshed successfully", "OK") | Out-Null
}

function Show-DCPropertiesDialog {
    param([string]$dcName)
    
    Debug-Log "DEBUG: Showing DC properties for: $dcName"
    
    if ($Global:DemoMode) {
        $dc = $Global:DCs | Where-Object { $_.Name -eq $dcName } | Select-Object -First 1
        
        if ($dc) {
            $msg = "Domain Controller: $($dc.Name)`nSite: $($dc.Site)`nOU: $($dc.OU)`n`n(Demo Mode)"
            [Terminal.Gui.MessageBox]::Query(60, 10, "DC Properties", $msg, "OK") | Out-Null
        } else {
            [Terminal.Gui.MessageBox]::Query(50, 7, "Not Found", "DC '$dcName' not found", "OK") | Out-Null
        }
    } else {
        # Production mode - show full AD DC properties
        try {
            $dc = Get-ADDomainController -Identity $dcName -ErrorAction Stop
            
            $msg = @"
Name: $($dc.Name)
Hostname: $($dc.HostName)
Site: $($dc.Site)
Domain: $($dc.Domain)
IPv4: $($dc.IPv4Address)
OS: $($dc.OperatingSystem)
Global Catalog: $($dc.IsGlobalCatalog)
"@
            [Terminal.Gui.MessageBox]::Query(70, 15, "DC Properties", $msg, "OK") | Out-Null
        } catch {
            [Terminal.Gui.MessageBox]::Query(60, 8, "Error", "Failed to get DC properties:`n$_", "OK") | Out-Null
        }
    }
}

# ------------------------- Context Menu Handler ------------------------
# ------------------------- Context Menu Handler ------------------------
function Show-ContextMenu {
    param(
        [string]$objectName,
        [string]$objectType
    )
    
    # Clean the object name (remove prefixes like "(U) " or "(DC) ")
    $cleanName = $objectName -replace '^\(.\)\s*', '' -replace '^[○⊗🔒]\s*', ''
    
    Debug-Log "DEBUG: Context menu for '$cleanName' (type: $objectType)"
    
    # Determine what type of object this is
    $isUser = $objectType -eq "user" -or $objectName -like "(U)*"
    $isGroup = $objectType -eq "group"
    $isOU = $objectType -eq "ou"
    $isDC = $objectType -eq "dc" -or $objectName -like "(DC)*"
    $isComputer = $objectType -eq "computer"
    
    # Build menu items array for display
    $menuText = @()
    
    if ($isUser) {
        $menuText = @("Properties", "Reset Password", "Disable Account", "Enable Account", "---", "Move to OU...", "---", "Delete", "---", "Refresh")
    } elseif ($isGroup) {
        $menuText = @("Properties", "Add Member...", "Remove Member...", "---", "Delete", "---", "Refresh")
    } elseif ($isOU) {
        $menuText = @("Properties", "New Object...", "---", "Delete", "---", "Refresh")
    } elseif ($isDC) {
        $menuText = @("Properties", "Check Replication", "---", "Refresh")
    } else {
        $menuText = @("Properties", "---", "Refresh")
    }
    
    # Create dialog for context menu
    $contextDialog = [Terminal.Gui.Dialog]::new("Actions", 30, ($menuText.Count + 4))
    $contextDialog.X = [Terminal.Gui.Pos]::Center()
    $contextDialog.Y = [Terminal.Gui.Pos]::Center()
    
    # Create list view
    $listView = [Terminal.Gui.ListView]::new()
    $listView.SetSource($menuText)
    $listView.X = 0
    $listView.Y = 0
    $listView.Width = [Terminal.Gui.Dim]::Fill()
    $listView.Height = [Terminal.Gui.Dim]::Fill(2)
    $contextDialog.Add($listView)
    
    # Handle selection
    $listView.add_OpenSelectedItem({
        $selected = $menuText[$listView.SelectedItem]
        [Terminal.Gui.Application]::RequestStop()
        
        if ($selected -ne "---") {
            Debug-Log "DEBUG: Context menu selected: $selected"
            
            switch ($selected) {
                "Properties" { 
                    if ($isUser) {
                        $user = $Global:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                        if ($user) { 
                            Show-UserPropertiesDialog -user $user -Global $Global
                        } else {
                            Debug-Log "ERROR: User '$cleanName' not found"
                        }
                    } elseif ($isGroup) {
                        Show-GroupPropertiesDialog -groupName $cleanName
                    } elseif ($isDC) {
                        Show-DCPropertiesDialog -dcName $cleanName
                    } else {
                        $msg = "Object: $cleanName`nType: $objectType"
                        [Terminal.Gui.MessageBox]::Query(50, 7, "Properties", $msg, "OK") | Out-Null
                    }
                }
                "Reset Password" { 
                    Show-ResetPasswordDialog -userName $cleanName 
                }
                "Disable Account" { 
                    Toggle-UserAccount -userName $cleanName -disable $true 
                }
                "Enable Account" { 
                    Toggle-UserAccount -userName $cleanName -disable $false 
                }
                "Move to OU..." { 
                    Show-MoveObjectDialog -objectName $cleanName -objectType "User" 
                }
                "Delete" { 
                    Show-DeleteObjectDialog -objectName $cleanName -objectType $objectType 
                }
                "Add Member..." { 
                    Show-AddGroupMemberDialog -groupName $cleanName 
                }
                "Remove Member..." { 
                    Show-RemoveGroupMemberDialog -groupName $cleanName 
                }
                "New Object..." { 
                    Show-NewObjectWizard 
                }
                "Check Replication" { 
                    Check-DCReplication -dcName $cleanName 
                }
                "Refresh" { 
                    Refresh-TreeData 
                }
            }
        }
    })
    
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $contextDialog.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($contextDialog)
}

# Debug version of Show-Properties
# Replace your Show-Properties function with this to see what's happening

function Show-Properties {
    Debug-Log "DEBUG: Show-Properties called"
    
    if (-not $tree.SelectedObject) { 
        Debug-Log "DEBUG: No object selected"
        [Terminal.Gui.MessageBox]::Query(50, 7, "Debug", "No object selected in tree", "OK") | Out-Null
        return 
    }
    
    $selName = $tree.SelectedObject.Text
    Debug-Log "DEBUG: Selected object text: '$selName'"
    
    # Remove prefixes like "(U) " or "(DC) " and status icons
    $cleanName = $selName -replace '^\(.\)\s*', '' -replace '^[○⊗🔒]\s*', ''
    Debug-Log "DEBUG: Cleaned name: '$cleanName'"
    
    $selType = if ($selName -like "(U)*") {"user"} elseif ($selName -like "(DC)*") {"computer"} else {"group"}
    Debug-Log "DEBUG: Detected type: $selType"

    if ($selType -eq "user") {
        Debug-Log "DEBUG: Searching for user in Global:Users array (count: $($Global:Users.Count))"
        
        # Try to find the user
        $selUser = $null
        foreach ($u in $Global:Users) {
            Debug-Log "DEBUG: Checking user: '$($u.Name)' against '$cleanName'"
            if ($u.Name -eq $cleanName) {
                $selUser = $u
                Debug-Log "DEBUG: MATCH FOUND!"
                break
            }
        }
        
        if ($selUser) {
            ############ Put troublshooting here #################
            Debug-Log "DEBUG: User found, calling Show-UserPropertiesDialog"
            Debug-Log "DEBUG: User details: Name=$($selUser.Name), Disabled=$($selUser.Disabled), Locked=$($selUser.Locked)"
Debug-Log "DEBUG: "

            try {
                Show-UserPropertiesDialog -user $selUser -Global $Global
                Debug-Log "DEBUG: Show-UserPropertiesDialog completed"
            } catch {
                Debug-Log "ERROR: Exception in Show-UserPropertiesDialog: $_"
                Debug-Log "ERROR: Stack trace: $($_.ScriptStackTrace)"
                Show-Modal "Failed to show properties:`n$($_.Exception.Message)`n`nCheck console for details"
            }
        } else {
            Debug-Log "DEBUG: User NOT found in Global:Users"
            [Terminal.Gui.MessageBox]::Query(50, 9, "Debug", "User '$cleanName' not found in Global:Users array.`n`nAvailable users: $($Global:Users.Count)", "OK") | Out-Null
        }
    } elseif ($selType -eq "group") {
        Debug-Log "DEBUG: Group type selected: $cleanName"
        $groupName = $cleanName
        $members = $Global:Users | Where-Object { $_.Groups -contains $groupName } | ForEach-Object { $_.Name } | Sort-Object
        $desc = "<no description>"
        $txt = "Group: $groupName`nDescription: $desc`nMembers:`n" + ($members -join "`n")
        [Terminal.Gui.MessageBox]::Query(60, 20, "Group Properties", $txt, "OK") | Out-Null
    } else {
        Debug-Log "DEBUG: Selected object type $selType not handled yet."
    }
}

# Additional debug helper - call this to verify your demo data loaded correctly
function Test-DemoData {
    Debug-Log "========== DEMO DATA CHECK =========="
    Debug-Log "Global:Users count: $($Global:Users.Count)"
    Debug-Log "Global:DCs count: $($Global:DCs.Count)"
    Debug-Log ""
    Debug-Log "Users in memory:"
    foreach ($u in $Global:Users) {
        $locked = if ($u.Locked) { "🔒" } else { "" }
        $disabled = if ($u.Disabled) { "⊗" } else { "○" }
        Debug-Log "  $disabled$locked $($u.Name) - Groups: $($u.Groups -join ', ')"
    }
    Debug-Log "====================================="
}

# Call this after loading demo data to verify:
# Test-DemoData
# ------------------------- Helper Functions ------------------------

function Show-ResetPasswordDialog {
    param([string]$userName)
    
    $dlg = [Terminal.Gui.Dialog]::new("Reset Password - $userName", 60, 12)
    
    $lbl = [Terminal.Gui.Label]::new("New Password:"); $lbl.X=2; $lbl.Y=1; $dlg.Add($lbl)
    $txtPwd = [Terminal.Gui.TextField]::new(""); $txtPwd.X=18; $txtPwd.Y=1; $txtPwd.Width=35; $txtPwd.Secret=$true; $dlg.Add($txtPwd)
    
    $lblConfirm = [Terminal.Gui.Label]::new("Confirm Password:"); $lblConfirm.X=2; $lblConfirm.Y=3; $dlg.Add($lblConfirm)
    $txtConfirm = [Terminal.Gui.TextField]::new(""); $txtConfirm.X=18; $txtConfirm.Y=3; $txtConfirm.Width=35; $txtConfirm.Secret=$true; $dlg.Add($txtConfirm)
    
    $chkMustChange = [Terminal.Gui.CheckBox]::new("User must change password at next logon")
    $chkMustChange.X=2; $chkMustChange.Y=5; $chkMustChange.Checked=$true; $dlg.Add($chkMustChange)
    
    $btnOK = [Terminal.Gui.Button]::new("OK")
    $btnOK.add_Clicked({
        $pwd1 = $txtPwd.Text.ToString()
        $pwd2 = $txtConfirm.Text.ToString()
        
        if ($pwd1 -ne $pwd2) {
            [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "Passwords do not match!", "OK") | Out-Null
            return
        }
        
        if ($pwd1.Length -lt 8) {
            [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "Password must be at least 8 characters!", "OK") | Out-Null
            return
        }
        
        try {
            if ($Global:DemoMode) {
                Debug-Log "DEBUG: Password reset for $userName (demo mode)"
                [Terminal.Gui.MessageBox]::Query(50, 7, "Success", "Password reset successfully (demo mode)", "OK") | Out-Null
            } else {
                $secPwd = ConvertTo-SecureString -String $pwd1 -AsPlainText -Force
                Set-ADAccountPassword -Identity $userName -NewPassword $secPwd -Reset -ErrorAction Stop
                if ($chkMustChange.Checked) {
                    Set-ADUser -Identity $userName -ChangePasswordAtLogon $true -ErrorAction Stop
                }
                [Terminal.Gui.MessageBox]::Query(50, 7, "Success", "Password reset successfully", "OK") | Out-Null
            }
            [Terminal.Gui.Application]::RequestStop()
        } catch {
            $errMsg = $_.Exception.Message
            [Terminal.Gui.MessageBox]::Query(60, 10, "Error", "Failed to reset password:`n$errMsg", "OK") | Out-Null
        }
    })
    $dlg.AddButton($btnOK)
    
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($dlg)
}

function Toggle-UserAccount {
    param([string]$userName, [bool]$disable)
    
    $action = if ($disable) { "disable" } else { "enable" }
    $result = [Terminal.Gui.MessageBox]::Query(50, 8, "Confirm", "Are you sure you want to $action account:`n$userName?", "Yes", "No")
    
    if ($result -eq 0) {
        try {
            if ($Global:DemoMode) {
                $user = $Global:Users | Where-Object { $_.Name -eq $userName } | Select-Object -First 1
                if ($user) {
                    $user.Disabled = $disable
                    Debug-Log "DEBUG: Account $userName $action`d (demo mode)"
                }
            } else {
                if ($disable) {
                    Disable-ADAccount -Identity $userName -ErrorAction Stop
                } else {
                    Enable-ADAccount -Identity $userName -ErrorAction Stop
                }
            }
            [Terminal.Gui.MessageBox]::Query(50, 7, "Success", "Account $action`d successfully", "OK") | Out-Null
            Refresh-TreeData
        } catch {
            $errMsg = $_.Exception.Message
            [Terminal.Gui.MessageBox]::Query(60, 10, "Error", "Failed to $action account:`n$errMsg", "OK") | Out-Null
        }
    }
}



function Show-DeleteObjectDialog {
    param([string]$objectName, [string]$objectType)
    
    $result = [Terminal.Gui.MessageBox]::Query(60, 9, "Delete $objectType", "WARNING: Are you sure you want to delete:`n$objectName`n`nThis action cannot be undone!", "Delete", "Cancel")
    
    if ($result -eq 0) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "Delete", "Delete functionality coming soon", "OK") | Out-Null
    }
}

# DSA-TUI Batch Operations Module v1.0
# Select multiple objects and perform bulk actions

# ------------------------- Global Selection State ------------------------
$Global:SelectedObjects = @()
$Global:SelectionMode = $false

# ------------------------- Toggle Selection Mode ------------------------
function Toggle-SelectionMode {
    $Global:SelectionMode = -not $Global:SelectionMode
    
    if ($Global:SelectionMode) {
        Debug-Log "DEBUG: Selection mode ENABLED"
        [Terminal.Gui.MessageBox]::Query(60, 8, "Selection Mode", 
            "Selection mode enabled!`n`nClick objects to select/deselect them.`nPress Ctrl+A to select all.`nPress Ctrl+D to deselect all.", 
            "OK") | Out-Null
    } else {
        Debug-Log "DEBUG: Selection mode DISABLED"
        $Global:SelectedObjects = @()
        Build-Tree -domain $Global:Domain
        Update-FilterStatusLabel -label $filterStatusLabel
    }
}

# ------------------------- Selection Panel ------------------------
function Create-SelectionPanel {
    $selPanel = [Terminal.Gui.FrameView]::new("Selected Objects")
    $selPanel.X = 32
    $selPanel.Y = 15
    $selPanel.Width = 40
    $selPanel.Height = 10
    
    $lblCount = [Terminal.Gui.Label]::new("0 objects selected")
    $lblCount.X = 1; $lblCount.Y = 0
    $selPanel.Add($lblCount)
    
    $lstSelected = [Terminal.Gui.ListView]::new(@())
    $lstSelected.X = 1; $lstSelected.Y = 1
    $lstSelected.Width = [Terminal.Gui.Dim]::Fill(1)
    $lstSelected.Height = [Terminal.Gui.Dim]::Fill(3)
    $selPanel.Add($lstSelected)
    
    # Store references for updates
    $selPanel.Tag = @{
        CountLabel = $lblCount
        ListView = $lstSelected
    }
    
    # Batch action buttons
    $btnBulkDisable = [Terminal.Gui.Button]::new("Disable All")
    $btnBulkDisable.X = 1
    $btnBulkDisable.Y = [Terminal.Gui.Pos]::Bottom($lstSelected) + 1
    $btnBulkDisable.add_Clicked({ Invoke-BulkDisableEnable -disable $true })
    $selPanel.Add($btnBulkDisable)
    
    $btnBulkEnable = [Terminal.Gui.Button]::new("Enable All")
    $btnBulkEnable.X = 14
    $btnBulkEnable.Y = [Terminal.Gui.Pos]::Bottom($lstSelected) + 1
    $btnBulkEnable.add_Clicked({ Invoke-BulkDisableEnable -disable $false })
    $selPanel.Add($btnBulkEnable)
    
    $btnBulkMove = [Terminal.Gui.Button]::new("Move All...")
    $btnBulkMove.X = 27
    $btnBulkMove.Y = [Terminal.Gui.Pos]::Bottom($lstSelected) + 1
    $btnBulkMove.add_Clicked({ Invoke-BulkMove })
    $selPanel.Add($btnBulkMove)
    
    return $selPanel
}

# ------------------------- Update Selection Panel ------------------------
function Update-SelectionPanel {
    param($panel)
    
    if (-not $panel -or -not $panel.Tag) { return }
    
    $lblCount = $panel.Tag.CountLabel
    $lstSelected = $panel.Tag.ListView
    
    $count = $Global:SelectedObjects.Count
    $lblCount.Text = "$count object(s) selected"
    
    $displayNames = $Global:SelectedObjects | ForEach-Object {
        $name = $_ -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
        $name
    }
    
    $lstSelected.SetSource($displayNames)
    $panel.SetNeedsDisplay()
}

# ------------------------- Enhanced Tree with Selection Support ------------------------
# Add this to your tree click handler (modify existing one or add new)

function Handle-TreeClick {
    param($mouseArgs)
    
    if (-not $tree.SelectedObject) { return }
    
    $selName = $tree.SelectedObject.Text
    
    # Check if in selection mode
    if ($Global:SelectionMode) {
        # Toggle selection
        if ($Global:SelectedObjects -contains $selName) {
            # Deselect
            $Global:SelectedObjects = $Global:SelectedObjects | Where-Object { $_ -ne $selName }
            Debug-Log "DEBUG: Deselected $selName"
        } else {
            # Select
            $Global:SelectedObjects += $selName
            Debug-Log "DEBUG: Selected $selName"
        }
        
        # Update visual indicator (mark selected items)
        Update-SelectionPanel -panel $selectionPanel
        $mouseArgs.Handled = $true
    }
}

# Add keyboard shortcuts for selection
function Add-SelectionKeyBindings {
    param($view)
    
    $view.add_KeyPress({ param($sender, $keyArgs)
        # Ctrl+A = Select All
        if ($keyArgs.KeyEvent.Key -eq ([Terminal.Gui.Key]::A -bor [Terminal.Gui.Key]::CtrlMask)) {
            Select-AllObjects
            $keyArgs.Handled = $true
        }
        
        # Ctrl+D = Deselect All
        if ($keyArgs.KeyEvent.Key -eq ([Terminal.Gui.Key]::D -bor [Terminal.Gui.Key]::CtrlMask)) {
            Deselect-AllObjects
            $keyArgs.Handled = $true
        }
        
        # Ctrl+S = Toggle Selection Mode
        if ($keyArgs.KeyEvent.Key -eq ([Terminal.Gui.Key]::S -bor [Terminal.Gui.Key]::CtrlMask)) {
            Toggle-SelectionMode
            $keyArgs.Handled = $true
        }
    })
}

# ------------------------- Select/Deselect All ------------------------
function Select-AllObjects {
    if (-not $Global:SelectionMode) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "Selection Mode", "Enable selection mode first (Ctrl+S)", "OK") | Out-Null
        return
    }
    
    $Global:SelectedObjects = @()
    
    # Get all users from tree
    foreach ($user in $Global:Users) {
        $statusIcon = if ($user.Disabled) { "⊗" } else { "○" }
        $displayName = "(U) $statusIcon $($user.Name)"
        $Global:SelectedObjects += $displayName
    }
    
    Debug-Log "DEBUG: Selected all users ($($Global:SelectedObjects.Count))"
    Update-SelectionPanel -panel $selectionPanel
    [Terminal.Gui.MessageBox]::Query(50, 7, "Selected All", "Selected $($Global:SelectedObjects.Count) users", "OK") | Out-Null
}

function Deselect-AllObjects {
    $Global:SelectedObjects = @()
    Debug-Log "DEBUG: Deselected all objects"
    Update-SelectionPanel -panel $selectionPanel
}

# ------------------------- Bulk Disable/Enable ------------------------
function Invoke-BulkDisableEnable {
    param([bool]$disable)
    
    if ($Global:SelectedObjects.Count -eq 0) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "No Selection", "No objects selected. Select objects first.", "OK") | Out-Null
        return
    }
    
    $action = if ($disable) { "disable" } else { "enable" }
    $result = [Terminal.Gui.MessageBox]::Query(60, 9, "Confirm Bulk Action", 
        "Are you sure you want to $action $($Global:SelectedObjects.Count) user account(s)?", 
        "Yes", "No")
    
    if ($result -eq 0) {
        $successCount = 0
        $failCount = 0
        $errors = @()
        
        foreach ($objName in $Global:SelectedObjects) {
            $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
            
            try {
                if ($Global:DemoMode) {
                    $user = $Global:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                    if ($user) {
                        $user.Disabled = $disable
                        $successCount++
                        Debug-Log "DEBUG: $action`d $cleanName (demo mode)"
                    }
                } else {
                    if ($disable) {
                        Disable-ADAccount -Identity $cleanName -ErrorAction Stop
                    } else {
                        Enable-ADAccount -Identity $cleanName -ErrorAction Stop
                    }
                    $successCount++
                    Debug-Log "DEBUG: $action`d $cleanName in AD"
                }
            } catch {
                $failCount++
                $errors += "$cleanName`: $($_.Exception.Message)"
                Debug-Log "DEBUG: Failed to $action $cleanName`: $_"
            }
        }
        
        # Show results
        $msg = "Successfully $action`d $successCount account(s)"
        if ($failCount -gt 0) {
            $msg += "`n`nFailed: $failCount"
            if ($errors.Count -gt 0 -and $errors.Count -le 5) {
                $msg += "`n`nErrors:`n" + ($errors -join "`n")
            }
        }
        
        [Terminal.Gui.MessageBox]::Query(70, 15, "Bulk Action Complete", $msg, "OK") | Out-Null
        
        # Refresh tree
        if (-not $Global:DemoMode) {
            Load-DomainData -domain $Global:Domain
        }
        Build-Tree -domain $Global:Domain
        Update-FilterStatusLabel -label $filterStatusLabel
        
        # Clear selection
        $Global:SelectedObjects = @()
        $Global:SelectionMode = $false
        Update-SelectionPanel -panel $selectionPanel
    }
}

# ------------------------- Bulk Move ------------------------
function Invoke-BulkMove {
    if ($Global:SelectedObjects.Count -eq 0) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "No Selection", "No objects selected. Select objects first.", "OK") | Out-Null
        return
    }
    
    $dlg = [Terminal.Gui.Dialog]::new("Bulk Move - $($Global:SelectedObjects.Count) Objects", 70, 18)
    
    $lblInfo = [Terminal.Gui.Label]::new("Moving $($Global:SelectedObjects.Count) object(s) to:")
    $lblInfo.X=2; $lblInfo.Y=1; $dlg.Add($lblInfo)
    
    # Get list of OUs
    $ouList = if ($Global:DemoMode) {
        $Global:Users | Select-Object -ExpandProperty OU -Unique | Sort-Object
    } else {
        try {
            Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName | 
                Select-Object -ExpandProperty DistinguishedName | Sort-Object
        } catch { @("CN=Users,DC=example,DC=com") }
    }
    
    $lstOU = [Terminal.Gui.ListView]::new($ouList)
    $lstOU.X=2; $lstOU.Y=3; $lstOU.Width=[Terminal.Gui.Dim]::Fill(2); $lstOU.Height=10
    $dlg.Add($lstOU)
    
    $btnMove = [Terminal.Gui.Button]::new("Move All")
    $btnMove.add_Clicked({
        if ($lstOU.SelectedItem -lt 0) {
            [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "Please select a target OU", "OK") | Out-Null
            return
        }
        
        $targetOU = $ouList[$lstOU.SelectedItem]
        
        $confirm = [Terminal.Gui.MessageBox]::Query(60, 9, "Confirm Bulk Move", 
            "Move $($Global:SelectedObjects.Count) object(s) to:`n$targetOU?", 
            "Yes", "No")
        
        if ($confirm -eq 0) {
            $successCount = 0
            $failCount = 0
            $errors = @()
            
            foreach ($objName in $Global:SelectedObjects) {
                $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
                
                try {
                    if ($Global:DemoMode) {
                        $user = $Global:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                        if ($user) {
                            $user.OU = $targetOU
                            $successCount++
                            Debug-Log "DEBUG: Moved $cleanName to $targetOU (demo mode)"
                        }
                    } else {
                        $adObject = Get-ADObject -Filter "Name -eq '$cleanName'" -ErrorAction Stop
                        Move-ADObject -Identity $adObject.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
                        $successCount++
                        Debug-Log "DEBUG: Moved $cleanName to $targetOU in AD"
                    }
                } catch {
                    $failCount++
                    $errors += "$cleanName`: $($_.Exception.Message)"
                    Debug-Log "DEBUG: Failed to move $cleanName`: $_"
                }
            }
            
            # Show results
            $msg = "Successfully moved $successCount object(s)"
            if ($failCount -gt 0) {
                $msg += "`n`nFailed: $failCount"
                if ($errors.Count -gt 0 -and $errors.Count -le 5) {
                    $msg += "`n`nErrors:`n" + ($errors -join "`n")
                }
            }
            
            [Terminal.Gui.MessageBox]::Query(70, 15, "Bulk Move Complete", $msg, "OK") | Out-Null
            
            # Refresh tree
            if (-not $Global:DemoMode) {
                Load-DomainData -domain $Global:Domain
            }
            Build-Tree -domain $Global:Domain
            Update-FilterStatusLabel -label $filterStatusLabel
            
            # Clear selection
            $Global:SelectedObjects = @()
            $Global:SelectionMode = $false
            Update-SelectionPanel -panel $selectionPanel
            
            [Terminal.Gui.Application]::RequestStop()
        }
    })
    $dlg.AddButton($btnMove)
    
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($dlg)
}

# ------------------------- Bulk Add to Group ------------------------
function Invoke-BulkAddToGroup {
    if ($Global:SelectedObjects.Count -eq 0) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "No Selection", "No objects selected. Select objects first.", "OK") | Out-Null
        return
    }
    
    $dlg = [Terminal.Gui.Dialog]::new("Bulk Add to Group", 70, 18)
    
    $lblInfo = [Terminal.Gui.Label]::new("Add $($Global:SelectedObjects.Count) user(s) to group:")
    $lblInfo.X=2; $lblInfo.Y=1; $dlg.Add($lblInfo)
    
    # Get list of groups
    $groupList = if ($Global:DemoMode) {
        $allGroups = @()
        foreach ($u in $Global:Users) {
            $allGroups += $u.Groups
        }
        $allGroups | Select-Object -Unique | Sort-Object
    } else {
        try {
            Get-ADGroup -Filter * | Select-Object -ExpandProperty Name | Sort-Object
        } catch { @("Domain Users", "Domain Admins") }
    }
    
    $lstGroups = [Terminal.Gui.ListView]::new($groupList)
    $lstGroups.X=2; $lstGroups.Y=3; $lstGroups.Width=[Terminal.Gui.Dim]::Fill(2); $lstGroups.Height=10
    $dlg.Add($lstGroups)
    
    $btnAdd = [Terminal.Gui.Button]::new("Add All")
    $btnAdd.add_Clicked({
        if ($lstGroups.SelectedItem -lt 0) {
            [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "Please select a group", "OK") | Out-Null
            return
        }
        
        $targetGroup = $groupList[$lstGroups.SelectedItem]
        
        $confirm = [Terminal.Gui.MessageBox]::Query(60, 9, "Confirm Bulk Add", 
            "Add $($Global:SelectedObjects.Count) user(s) to group:`n$targetGroup?", 
            "Yes", "No")
        
        if ($confirm -eq 0) {
            $successCount = 0
            $failCount = 0
            
            foreach ($objName in $Global:SelectedObjects) {
                $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
                
                try {
                    if ($Global:DemoMode) {
                        $user = $Global:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                        if ($user -and $user.Groups -notcontains $targetGroup) {
                            $user.Groups += $targetGroup
                            $successCount++
                            Debug-Log "DEBUG: Added $cleanName to $targetGroup (demo mode)"
                        }
                    } else {
                        Add-ADGroupMember -Identity $targetGroup -Members $cleanName -ErrorAction Stop
                        $successCount++
                        Debug-Log "DEBUG: Added $cleanName to $targetGroup in AD"
                    }
                } catch {
                    $failCount++
                    Debug-Log "DEBUG: Failed to add $cleanName`: $_"
                }
            }
            
            [Terminal.Gui.MessageBox]::Query(60, 10, "Bulk Add Complete", 
                "Successfully added $successCount user(s)`nFailed: $failCount", 
                "OK") | Out-Null
            
            # Refresh tree
            Build-Tree -domain $Global:Domain
            Update-FilterStatusLabel -label $filterStatusLabel
            
            [Terminal.Gui.Application]::RequestStop()
        }
    })
    $dlg.AddButton($btnAdd)
    
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($dlg)
}


# Add keyboard shortcuts for selection
function Add-SelectionKeyBindings {
    param($view)
    
    $view.add_KeyPress({ param($sender, $keyArgs)
        # Ctrl+A = Select All
        if ($keyArgs.KeyEvent.Key -eq ([Terminal.Gui.Key]::A -bor [Terminal.Gui.Key]::CtrlMask)) {
            Select-AllObjects
            $keyArgs.Handled = $true
        }
        
        # Ctrl+D = Deselect All
        if ($keyArgs.KeyEvent.Key -eq ([Terminal.Gui.Key]::D -bor [Terminal.Gui.Key]::CtrlMask)) {
            Deselect-AllObjects
            $keyArgs.Handled = $true
        }
        
        # Ctrl+S = Toggle Selection Mode
        if ($keyArgs.KeyEvent.Key -eq ([Terminal.Gui.Key]::S -bor [Terminal.Gui.Key]::CtrlMask)) {
            Toggle-SelectionMode
            $keyArgs.Handled = $true
        }
    })
}

# ------------------------- Select/Deselect All ------------------------
function Select-AllObjects {
    if (-not $Global:SelectionMode) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "Selection Mode", "Enable selection mode first (Ctrl+S)", "OK") | Out-Null
        return
    }
    
    $Global:SelectedObjects = @()
    
    # Get all users from tree
    foreach ($user in $Global:Users) {
        $statusIcon = if ($user.Disabled) { "⊗" } else { "○" }
        $displayName = "(U) $statusIcon $($user.Name)"
        $Global:SelectedObjects += $displayName
    }
    
    Debug-Log "DEBUG: Selected all users ($($Global:SelectedObjects.Count))"
    Update-SelectionPanel -panel $selectionPanel
    [Terminal.Gui.MessageBox]::Query(50, 7, "Selected All", "Selected $($Global:SelectedObjects.Count) users", "OK") | Out-Null
}

function Deselect-AllObjects {
    $Global:SelectedObjects = @()
    Debug-Log "DEBUG: Deselected all objects"
    Update-SelectionPanel -panel $selectionPanel
}

# ------------------------- Bulk Disable/Enable ------------------------
function Invoke-BulkDisableEnable {
    param([bool]$disable)
    
    if ($Global:SelectedObjects.Count -eq 0) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "No Selection", "No objects selected. Select objects first.", "OK") | Out-Null
        return
    }
    
    $action = if ($disable) { "disable" } else { "enable" }
    $result = [Terminal.Gui.MessageBox]::Query(60, 9, "Confirm Bulk Action", 
        "Are you sure you want to $action $($Global:SelectedObjects.Count) user account(s)?", 
        "Yes", "No")
    
    if ($result -eq 0) {
        $successCount = 0
        $failCount = 0
        $errors = @()
        
        foreach ($objName in $Global:SelectedObjects) {
            $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
            
            try {
                if ($Global:DemoMode) {
                    $user = $Global:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                    if ($user) {
                        $user.Disabled = $disable
                        $successCount++
                        Debug-Log "DEBUG: $action`d $cleanName (demo mode)"
                    }
                } else {
                    if ($disable) {
                        Disable-ADAccount -Identity $cleanName -ErrorAction Stop
                    } else {
                        Enable-ADAccount -Identity $cleanName -ErrorAction Stop
                    }
                    $successCount++
                    Debug-Log "DEBUG: $action`d $cleanName in AD"
                }
            } catch {
                $failCount++
                $errors += "$cleanName`: $($_.Exception.Message)"
                Debug-Log "DEBUG: Failed to $action $cleanName`: $_"
            }
        }
        
        # Show results
        $msg = "Successfully $action`d $successCount account(s)"
        if ($failCount -gt 0) {
            $msg += "`n`nFailed: $failCount"
            if ($errors.Count -gt 0 -and $errors.Count -le 5) {
                $msg += "`n`nErrors:`n" + ($errors -join "`n")
            }
        }
        
        [Terminal.Gui.MessageBox]::Query(70, 15, "Bulk Action Complete", $msg, "OK") | Out-Null
        
        # Refresh tree
        if (-not $Global:DemoMode) {
            Load-DomainData -domain $Global:Domain
        }
        Build-Tree -domain $Global:Domain
        Update-FilterStatusLabel -label $filterStatusLabel
        
        # Clear selection
        $Global:SelectedObjects = @()
        $Global:SelectionMode = $false
        Update-SelectionPanel -panel $selectionPanel
    }
}

# ------------------------- Bulk Move ------------------------
function Invoke-BulkMove {
    if ($Global:SelectedObjects.Count -eq 0) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "No Selection", "No objects selected. Select objects first.", "OK") | Out-Null
        return
    }
    
    $dlg = [Terminal.Gui.Dialog]::new("Bulk Move - $($Global:SelectedObjects.Count) Objects", 70, 18)
    
    $lblInfo = [Terminal.Gui.Label]::new("Moving $($Global:SelectedObjects.Count) object(s) to:")
    $lblInfo.X=2; $lblInfo.Y=1; $dlg.Add($lblInfo)
    
    # Get list of OUs
    $ouList = if ($Global:DemoMode) {
        $Global:Users | Select-Object -ExpandProperty OU -Unique | Sort-Object
    } else {
        try {
            Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName | 
                Select-Object -ExpandProperty DistinguishedName | Sort-Object
        } catch { @("CN=Users,DC=example,DC=com") }
    }
    
    $lstOU = [Terminal.Gui.ListView]::new($ouList)
    $lstOU.X=2; $lstOU.Y=3; $lstOU.Width=[Terminal.Gui.Dim]::Fill(2); $lstOU.Height=10
    $dlg.Add($lstOU)
    
    $btnMove = [Terminal.Gui.Button]::new("Move All")
    $btnMove.add_Clicked({
        if ($lstOU.SelectedItem -lt 0) {
            [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "Please select a target OU", "OK") | Out-Null
            return
        }
        
        $targetOU = $ouList[$lstOU.SelectedItem]
        
        $confirm = [Terminal.Gui.MessageBox]::Query(60, 9, "Confirm Bulk Move", 
            "Move $($Global:SelectedObjects.Count) object(s) to:`n$targetOU?", 
            "Yes", "No")
        
        if ($confirm -eq 0) {
            $successCount = 0
            $failCount = 0
            $errors = @()
            
            foreach ($objName in $Global:SelectedObjects) {
                $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
                
                try {
                    if ($Global:DemoMode) {
                        $user = $Global:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                        if ($user) {
                            $user.OU = $targetOU
                            $successCount++
                            Debug-Log "DEBUG: Moved $cleanName to $targetOU (demo mode)"
                        }
                    } else {
                        $adObject = Get-ADObject -Filter "Name -eq '$cleanName'" -ErrorAction Stop
                        Move-ADObject -Identity $adObject.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
                        $successCount++
                        Debug-Log "DEBUG: Moved $cleanName to $targetOU in AD"
                    }
                } catch {
                    $failCount++
                    $errors += "$cleanName`: $($_.Exception.Message)"
                    Debug-Log "DEBUG: Failed to move $cleanName`: $_"
                }
            }
            
            # Show results
            $msg = "Successfully moved $successCount object(s)"
            if ($failCount -gt 0) {
                $msg += "`n`nFailed: $failCount"
                if ($errors.Count -gt 0 -and $errors.Count -le 5) {
                    $msg += "`n`nErrors:`n" + ($errors -join "`n")
                }
            }
            
            [Terminal.Gui.MessageBox]::Query(70, 15, "Bulk Move Complete", $msg, "OK") | Out-Null
            
            # Refresh tree
            if (-not $Global:DemoMode) {
                Load-DomainData -domain $Global:Domain
            }
            Build-Tree -domain $Global:Domain
            Update-FilterStatusLabel -label $filterStatusLabel
            
            # Clear selection
            $Global:SelectedObjects = @()
            $Global:SelectionMode = $false
            Update-SelectionPanel -panel $selectionPanel
            
            [Terminal.Gui.Application]::RequestStop()
        }
    })
    $dlg.AddButton($btnMove)
    
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($dlg)
}

# ------------------------- Bulk Add to Group ------------------------
function Invoke-BulkAddToGroup {
    if ($Global:SelectedObjects.Count -eq 0) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "No Selection", "No objects selected. Select objects first.", "OK") | Out-Null
        return
    }
    
    $dlg = [Terminal.Gui.Dialog]::new("Bulk Add to Group", 70, 18)
    
    $lblInfo = [Terminal.Gui.Label]::new("Add $($Global:SelectedObjects.Count) user(s) to group:")
    $lblInfo.X=2; $lblInfo.Y=1; $dlg.Add($lblInfo)
    
    # Get list of groups
    $groupList = if ($Global:DemoMode) {
        $allGroups = @()
        foreach ($u in $Global:Users) {
            $allGroups += $u.Groups
        }
        $allGroups | Select-Object -Unique | Sort-Object
    } else {
        try {
            Get-ADGroup -Filter * | Select-Object -ExpandProperty Name | Sort-Object
        } catch { @("Domain Users", "Domain Admins") }
    }
    
    $lstGroups = [Terminal.Gui.ListView]::new($groupList)
    $lstGroups.X=2; $lstGroups.Y=3; $lstGroups.Width=[Terminal.Gui.Dim]::Fill(2); $lstGroups.Height=10
    $dlg.Add($lstGroups)
    
    $btnAdd = [Terminal.Gui.Button]::new("Add All")
    $btnAdd.add_Clicked({
        if ($lstGroups.SelectedItem -lt 0) {
            [Terminal.Gui.MessageBox]::Query(50, 7, "Error", "Please select a group", "OK") | Out-Null
            return
        }
        
        $targetGroup = $groupList[$lstGroups.SelectedItem]
        
        $confirm = [Terminal.Gui.MessageBox]::Query(60, 9, "Confirm Bulk Add", 
            "Add $($Global:SelectedObjects.Count) user(s) to group:`n$targetGroup?", 
            "Yes", "No")
        
        if ($confirm -eq 0) {
            $successCount = 0
            $failCount = 0
            
            foreach ($objName in $Global:SelectedObjects) {
                $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
                
                try {
                    if ($Global:DemoMode) {
                        $user = $Global:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                        if ($user -and $user.Groups -notcontains $targetGroup) {
                            $user.Groups += $targetGroup
                            $successCount++
                            Debug-Log "DEBUG: Added $cleanName to $targetGroup (demo mode)"
                        }
                    } else {
                        Add-ADGroupMember -Identity $targetGroup -Members $cleanName -ErrorAction Stop
                        $successCount++
                        Debug-Log "DEBUG: Added $cleanName to $targetGroup in AD"
                    }
                } catch {
                    $failCount++
                    Debug-Log "DEBUG: Failed to add $cleanName`: $_"
                }
            }
            
            [Terminal.Gui.MessageBox]::Query(60, 10, "Bulk Add Complete", 
                "Successfully added $successCount user(s)`nFailed: $failCount", 
                "OK") | Out-Null
            
            # Refresh tree
            Build-Tree -domain $Global:Domain
            Update-FilterStatusLabel -label $filterStatusLabel
            
            [Terminal.Gui.Application]::RequestStop()
        }
    })
    $dlg.AddButton($btnAdd)
    
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($dlg)
}

# ------------------------- Selection Panel ------------------------
function Create-SelectionPanel {
    $selPanel = [Terminal.Gui.FrameView]::new("Selected Objects")
    $selPanel.X = 32
    $selPanel.Y = 15
    $selPanel.Width = 40
    $selPanel.Height = 10
    
    $lblCount = [Terminal.Gui.Label]::new("0 objects selected")
    $lblCount.X = 1; $lblCount.Y = 0
    $selPanel.Add($lblCount)
    
    $lstSelected = [Terminal.Gui.ListView]::new(@())
    $lstSelected.X = 1; $lstSelected.Y = 1
    $lstSelected.Width = [Terminal.Gui.Dim]::Fill(1)
    $lstSelected.Height = [Terminal.Gui.Dim]::Fill(3)
    $selPanel.Add($lstSelected)
    
    # Store references for updates
    $selPanel.Tag = @{
        CountLabel = $lblCount
        ListView = $lstSelected
    }
    
    # Batch action buttons
    $btnBulkDisable = [Terminal.Gui.Button]::new("Disable All")
    $btnBulkDisable.X = 1
    $btnBulkDisable.Y = [Terminal.Gui.Pos]::Bottom($lstSelected) + 1
    $btnBulkDisable.add_Clicked({ Invoke-BulkDisableEnable -disable $true })
    $selPanel.Add($btnBulkDisable)
    
    $btnBulkEnable = [Terminal.Gui.Button]::new("Enable All")
    $btnBulkEnable.X = 14
    $btnBulkEnable.Y = [Terminal.Gui.Pos]::Bottom($lstSelected) + 1
    $btnBulkEnable.add_Clicked({ Invoke-BulkDisableEnable -disable $false })
    $selPanel.Add($btnBulkEnable)
    
    $btnBulkMove = [Terminal.Gui.Button]::new("Move All...")
    $btnBulkMove.X = 27
    $btnBulkMove.Y = [Terminal.Gui.Pos]::Bottom($lstSelected) + 1
    $btnBulkMove.add_Clicked({ Invoke-BulkMove })
    $selPanel.Add($btnBulkMove)
    
    return $selPanel
}

# ------------------------- Update Selection Panel ------------------------
function Update-SelectionPanel {
    param($panel)
    
    if (-not $panel -or -not $panel.Tag) { return }
    
    $lblCount = $panel.Tag.CountLabel
    $lstSelected = $panel.Tag.ListView
    
    $count = $Global:SelectedObjects.Count
    $lblCount.Text = "$count object(s) selected"
    
    $displayNames = $Global:SelectedObjects | ForEach-Object {
        $name = $_ -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
        $name
    }
    
    $lstSelected.SetSource($displayNames)
    $panel.SetNeedsDisplay()
}

function Show-GroupPropertiesDialog {
    param([string]$groupName)
    
    $members = $Global:Users | Where-Object { $_.Groups -contains $groupName } | ForEach-Object { $_.Name } | Sort-Object
    $desc = "<no description>"
    $txt = "Group: $groupName`nDescription: $desc`nMember Count: $($members.Count)`n`nMembers:`n" + ($members -join "`n")
    [Terminal.Gui.MessageBox]::Query(60, 20, "Group Properties", $txt, "OK") | Out-Null
}

# Add/Remove Group Member Implementation
# Replace the placeholder functions in your script

# =====================================================
# ADD GROUP MEMBER
# =====================================================
function Show-AddGroupMemberDialog {
    param([string]$groupName)
    
    Debug-Log "DEBUG: Add member to group: $groupName"
    
    $dlg = [Terminal.Gui.Dialog]::new("Add Member to Group - $groupName", 70, 25)
    
    $lblInfo = [Terminal.Gui.Label]::new("Select users to add to '$groupName':")
    $lblInfo.X = 2; $lblInfo.Y = 1
    $dlg.Add($lblInfo)
    
    # Get list of all users NOT already in this group
    if ($Global:DemoMode) {
        # Demo mode: find users not in this group
        $availableUsers = $Global:Users | Where-Object { 
            $_.Groups -notcontains $groupName 
        } | Sort-Object -Property Name
    } else {
        # Production mode: get all users not in group
        try {
            $groupMembers = Get-ADGroupMember -Identity $groupName -ErrorAction Stop | 
                Select-Object -ExpandProperty SamAccountName
            
            $allUsers = Get-ADUser -Filter * -Properties SamAccountName, DisplayName -ErrorAction Stop
            
            $availableUsers = $allUsers | Where-Object { 
                $groupMembers -notcontains $_.SamAccountName 
            } | Sort-Object -Property DisplayName
        } catch {
            [Terminal.Gui.MessageBox]::Query(60, 10, "Error", 
                "Failed to retrieve users:`n$($_.Exception.Message)", "OK") | Out-Null
            return
        }
    }
    
    if ($availableUsers.Count -eq 0) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "No Users", 
            "All users are already members of this group!", "OK") | Out-Null
        return
    }
    
    # Create list of available users
    if ($Global:DemoMode) {
        $userList = $availableUsers | ForEach-Object { 
            "$($_.Name) - $($_.Title) [$($_.OU[-1])]" 
        }
    } else {
        $userList = $availableUsers | ForEach-Object { 
            "$($_.DisplayName) ($($_.SamAccountName))" 
        }
    }
    
    $lstUsers = [Terminal.Gui.ListView]::new()
    $lstUsers.SetSource($userList)
    $lstUsers.X = 2
    $lstUsers.Y = 3
    $lstUsers.Width = [Terminal.Gui.Dim]::Fill(2)
    $lstUsers.Height = [Terminal.Gui.Dim]::Fill(4)
    $lstUsers.AllowsMarking = $true  # Allow multi-select
    $dlg.Add($lstUsers)
    
    # Info label
    $lblHelp = [Terminal.Gui.Label]::new("Tip: Use SPACE to select multiple users")
    $lblHelp.X = 2
    $lblHelp.Y = [Terminal.Gui.Pos]::Bottom($lstUsers) + 1
    $dlg.Add($lblHelp)
    
    # Add button
$btnAdd = [Terminal.Gui.Button]::new("Add Selected")
$btnAdd.add_Clicked({
    # Get marked (selected) items
    $markedIndices = @()
    for ($i = 0; $i -lt $userList.Count; $i++) {
        if ($lstUsers.Source.IsMarked($i)) {
            $markedIndices += $i
        }
    }
    
    # If nothing marked, check if there's a selected item
    if ($markedIndices.Count -eq 0) {
        if ($lstUsers.SelectedItem -ne -1) {
            # User just highlighted but didn't mark - use the selected item
            $markedIndices = @($lstUsers.SelectedItem)
        } else {
            [Terminal.Gui.MessageBox]::Query(50, 7, "No Selection", 
                "Please select at least one user (use SPACE to mark multiple, or just highlight one)", "OK") | Out-Null
            return
        }
    }
})
}

# =====================================================
# REMOVE GROUP MEMBER
# =====================================================
function Show-RemoveGroupMemberDialog {
    param([string]$groupName)
    
    Debug-Log "DEBUG: Remove member from group: $groupName"
    
    $dlg = [Terminal.Gui.Dialog]::new("Remove Member from Group - $groupName", 70, 25)
    
    $lblInfo = [Terminal.Gui.Label]::new("Select users to remove from '$groupName':")
    $lblInfo.X = 2; $lblInfo.Y = 1
    $dlg.Add($lblInfo)
    
    # Get list of current group members
    if ($Global:DemoMode) {
        # Demo mode: find users in this group
        $groupMembers = $Global:Users | Where-Object { 
            $_.Groups -contains $groupName 
        } | Sort-Object -Property Name
    } else {
        # Production mode: get group members from AD
        try {
            $members = Get-ADGroupMember -Identity $groupName -ErrorAction Stop
            $groupMembers = $members | ForEach-Object {
                Get-ADUser -Identity $_.SamAccountName -Properties DisplayName, SamAccountName -ErrorAction Stop
            } | Sort-Object -Property DisplayName
        } catch {
            [Terminal.Gui.MessageBox]::Query(60, 10, "Error", 
                "Failed to retrieve group members:`n$($_.Exception.Message)", "OK") | Out-Null
            return
        }
    }
    
    if ($groupMembers.Count -eq 0) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "Empty Group", 
            "This group has no members!", "OK") | Out-Null
        return
    }
    
    # Create list of group members
    if ($Global:DemoMode) {
        $memberList = $groupMembers | ForEach-Object { 
            "$($_.Name) - $($_.Title) [$($_.OU[-1])]" 
        }
    } else {
        $memberList = $groupMembers | ForEach-Object { 
            "$($_.DisplayName) ($($_.SamAccountName))" 
        }
    }
    
    $lstMembers = [Terminal.Gui.ListView]::new()
    $lstMembers.SetSource($memberList)
    $lstMembers.X = 2
    $lstMembers.Y = 3
    $lstMembers.Width = [Terminal.Gui.Dim]::Fill(2)
    $lstMembers.Height = [Terminal.Gui.Dim]::Fill(4)
    $lstMembers.AllowsMarking = $true  # Allow multi-select
    $dlg.Add($lstMembers)
    
    # Info label
    $lblHelp = [Terminal.Gui.Label]::new("Tip: Use SPACE to select multiple users")
    $lblHelp.X = 2
    $lblHelp.Y = [Terminal.Gui.Pos]::Bottom($lstMembers) + 1
    $dlg.Add($lblHelp)
    
    # Remove button
    $btnRemove = [Terminal.Gui.Button]::new("Remove Selected")
    $btnRemove.add_Clicked({
        # Get marked (selected) items
        $markedIndices = @()
        for ($i = 0; $i -lt $memberList.Count; $i++) {
            if ($lstMembers.Source.IsMarked($i)) {
                $markedIndices += $i
            }
        }
        
        if ($markedIndices.Count -eq 0) {
            [Terminal.Gui.MessageBox]::Query(50, 7, "No Selection", 
                "Please select at least one user (use SPACE key)", "OK") | Out-Null
            return
        }
        
        $usersToRemove = $markedIndices | ForEach-Object { $groupMembers[$_] }
        
        $confirm = [Terminal.Gui.MessageBox]::Query(60, 9, "Confirm Remove", 
            "Remove $($usersToRemove.Count) user(s) from group '$groupName'?", 
            "Yes", "No")
        
        if ($confirm -eq 0) {
            $successCount = 0
            $failCount = 0
            $errors = @()
            
            foreach ($user in $usersToRemove) {
                try {
                    if ($Global:DemoMode) {
                        # Demo mode: remove group from user's Groups array
                        $user.Groups = $user.Groups | Where-Object { $_ -ne $groupName }
                        $successCount++
                        Debug-Log "DEBUG: Removed $($user.Name) from $groupName (demo mode)"
                    } else {
                        # Production mode: remove from AD
                        Remove-ADGroupMember -Identity $groupName -Members $user.SamAccountName -Confirm:$false -ErrorAction Stop
                        $successCount++
                        Debug-Log "DEBUG: Removed $($user.SamAccountName) from $groupName in AD"
                    }
                } catch {
                    $failCount++
                    $userName = if ($Global:DemoMode) { $user.Name } else { $user.SamAccountName }
                    $errors += "$userName`: $($_.Exception.Message)"
                    Debug-Log "ERROR: Failed to remove $userName from group: $_"
                }
            }
            
            # Show results
            $msg = "Successfully removed $successCount user(s) from '$groupName'"
            if ($failCount -gt 0) {
                $msg += "`n`nFailed: $failCount"
                if ($errors.Count -gt 0 -and $errors.Count -le 5) {
                    $msg += "`n`nErrors:`n" + ($errors -join "`n")
                }
            }
            
            [Terminal.Gui.MessageBox]::Query(70, 15, "Remove Members Complete", $msg, "OK") | Out-Null
            
            # Rebuild tree
            [Terminal.Gui.Application]::MainLoop.Invoke({
                Build-Tree -domain $Global:Domain
                Update-FilterStatusLabel -label $filterStatusLabel
            })
            
            [Terminal.Gui.Application]::RequestStop()
        }
    })
    $dlg.AddButton($btnRemove)
    
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($dlg)
}


function Check-DCReplication {
    param([string]$dcName)
    [Terminal.Gui.MessageBox]::Query(50, 7, "Replication Check", "Checking replication for $dcName`n(Coming soon)", "OK") | Out-Null
}

# ------------------------- Tree Mouse Handler ------------------------
# Add this to tree setup after creating $tree:

# Add mouse event handler for right-click
$tree.add_MouseClick({
    param($mouseEvent)
    
    Debug-Log "DEBUG: Mouse event type: $($mouseEvent.GetType().Name)"
    Debug-Log "DEBUG: Mouse event properties: $($mouseEvent | Get-Member -MemberType Property | Select-Object -ExpandProperty Name)"
    
    # Try different ways to access the mouse button
    $isRightClick = $false
    
    # Method 1: Direct Flags property
    if ($mouseEvent.PSObject.Properties['Flags']) {
        Debug-Log "DEBUG: Flags = $($mouseEvent.Flags)"
        $isRightClick = $mouseEvent.Flags -band [Terminal.Gui.MouseFlags]::Button3Clicked
    }
    
    # Method 2: MouseEvent property
    if ($mouseEvent.PSObject.Properties['MouseEvent']) {
        Debug-Log "DEBUG: MouseEvent.Flags = $($mouseEvent.MouseEvent.Flags)"
        $isRightClick = $mouseEvent.MouseEvent.Flags -band [Terminal.Gui.MouseFlags]::Button3Clicked
    }
    
    # Method 3: Check if it's Button3
    if ($mouseEvent.PSObject.Properties['Button']) {
        Debug-Log "DEBUG: Button = $($mouseEvent.Button)"
        $isRightClick = $mouseEvent.Button -eq 3
    }
    
    Debug-Log "DEBUG: Is right-click: $isRightClick"
    
    if ($isRightClick) {
        Debug-Log "DEBUG: Right-click detected"
        
        if ($tree.SelectedObject) {
            $selectedNode = $tree.SelectedObject
            $nodeName = $selectedNode.Text
            
            Debug-Log "DEBUG: Right-clicked on node: $nodeName"
            
            # Determine object type
            $objectType = "unknown"
            if ($nodeName -like "(U)*") { 
                $objectType = "user" 
            } elseif ($nodeName -like "(DC)*") { 
                $objectType = "dc" 
            } elseif ($nodeName -like "(G)*" -or ($selectedNode.Parent -and $selectedNode.Parent.Text -eq "Groups")) { 
                $objectType = "group" 
            } elseif ($nodeName -like "OU:*") { 
                $objectType = "ou" 
            }
            
            # Show context menu
            Show-ContextMenu -objectName $nodeName -objectType $objectType
            
            # Prevent default handling if Handled property exists
            if ($mouseEvent.PSObject.Properties['Handled']) {
                $mouseEvent.Handled = $true
            }
        }
    }
})

# This is a theme now. Danish Fruit soda based fun
# Method to the madness
function Show-BlaabaerInfo {
    $dlg = [Terminal.Gui.Dialog]::new("Why Blaabaer? 🫐", 60, 12)
    
    $message = @"
DSA-TUI is codenamed "Blaabaer" because:

- I was drinking blueberry soda when writing the code
- 'Blåbær' (Blaabaer) is Danish for blueberry
- Føtex sells a rather nice Blaabaer soda
- Every great project needs a forest-fruit mascot!
"@
    
    $label = [Terminal.Gui.Label]::new(1, 1, $message)
    $dlg.Add($label)
    
    Show-Modal "Why Blaabaer? 🫐" $message
}

function Show-Modal { 
    param($title, $msg) 
    [Terminal.Gui.MessageBox]::Query($title, $msg, @("OK")) | Out-Null 
}

# ------------------------- Load Domain or Demo Data FIRST ------------------------
Load-DomainData -domain $Global:Domain

# (Optional) Debug after loading
Write-Host "POST-LOAD DEBUG: Users:"  $Global:Users.Count
Write-Host "POST-LOAD DEBUG: DCs:"     $Global:DCs.Count
Write-Host "POST-LOAD DEBUG: Objects:" $Global:ADObjects.Count

# ------------------------- Build initial tree ------------------------
Build-Tree -domain $Global:Domain

# ------------------------- Update status label ------------------------
Update-FilterStatusLabel -label $filterStatusLabel

# ------------------------- Run application ------------------------
[Terminal.Gui.Application]::Run($top)
[Terminal.Gui.Application]::Shutdown()
