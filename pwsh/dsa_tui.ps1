<#

DSA-TUI Text Mode version of dsa.msc for powershell
Locked-in baseline: dynamic resize, menu, demo data mirrors prod format, Change Domain fixed, fixed DC selection, full production AD object detection, properties modal, AD search popup

https://nyheder.tv2.dk/lokalt/2021-10-19-er-det-her-postbuddets-vaerste-skraek-hvem-pokker-har-fundet-paa-det-her

===========================================================================================
 DSA-TUI Blåbær — Active Directory TUI Tool
 Historical Build Notes and Change Log
===========================================================================================

 1.0.0.0  (Initial Experimental)
 - First internal test build. Basic TUI scaffolding only.
 - Bare Window + Menu + Exit only. No AD integration.
 - Non-functional placeholder TreeView.

 1.1.0.0  (Initial AD Integration)
 - Added basic Domain Bind and LDAP query functions.
 - Added Build-Tree function (initial non-recursive prototype).
 - Added minimal Properties popup (placeholder).

 1.1.0.1  (Bugfix)
 - Fixed null-domain crash.
 - Fixed title bar misalignment on Linux/macOS terminals.

 1.2.0.0  (Tree + Navigation)
 - Introduced TreeView AD structure display.
 - Added OU expansion, user nodes, group nodes.
 - Implemented Refresh (F5) bound to Build-Tree.
 - Added basic status bar with simple messages only.

 1.2.0.1  (Bugfix)
 - Fixed node-expansion crash when encountering empty OUs.
 - Fixed cosmetic padding/spacing inconsistencies.

 1.3.0.0  (Selection, Node Info)
 - Added node selection handling.
 - Added Show-Properties modal (initial version).
 - Added object type detection for icons (U/G/OU/DC).

 1.3.0.1  (Bugfix)
 - Fixed Properties dialog not clearing previous content.
 - Fixed MessageBox misalignment under Terminal.Gui 1.16.

 1.4.0.0  (Filter System v1)
 - Added Filter Panel (right side) with toggles.
 - Added Global:FilterOptions hashtable.
 - Added Update-FilterStatusLabel function.
 - Added name-filter support (“search by name”).

 1.4.0.1  (Bugfix)
 - Fixed filter panel overlapping TreeView.
 - Fixed name-filter not persisting during refresh.
 - Fixed missing redrawing after filter changes.

 1.5.0.0  (Full Refresh Engine + Searchable Properties Rewrite)
 - Major rewrite of Refresh / Build-Tree pipeline.
 - Added "Searchable Attributes" handling:
       * name
       * displayName
       * sAMAccountName
       * userPrincipalName
       * givenName / sn
 - Optimized LDAP lookups to only fetch required fields.
 - Added caching to reduce domain traffic.

 1.5.0.1  (Bugfix)
 - Fixed several typos in attribute lookup keys (displayName vs displayname).
 - Fixed "OU:" prefixes duplicating on some nodes.
 - Fixed sorting of users/groups inside OUs.

 1.5.0.2  (Bugfix)
 - Fixed rare crash when node had malformed DN.
 - Corrected spacing in status label (“No filters active” line).
 - Fixed cosmetic typo: “Serach” → “Search”.

 1.5.0.3  (Bugfix)
 - Fixed name filter not updating until second refresh.
 - Fixed stale nodes remaining after filter changes.
 - Added missing "show groups" toggle check.

 1.6.0.0  (Major UI Improvements)
 - Introduced fully functional modal system (non-blocking).
 - Replaced Read-Host prompts with TUI modals.
 - Added Create-FilterPanel (initial modern version).
 - Added Show-QuickFilterDialog function.
 - Added TreeView bounds fixes + visibility fixes.

 1.6.0.1  (Bugfix)
 - Fixed filter panel incorrectly covering entire window.
 - Fixed TreeView being hidden beneath filter layer.
 - Fixed Update-FilterStatusLabel rendering directly on main window.

 1.6.0.2  (Bugfix)
 - Fixed MessageBox defaultButton index errors.
 - Fixed PasswordGenerator dialog always showing success regardless of click.
 - Corrected missing `.Visible = $false` on filter panel startup.

 1.6.3.0  (New Feature: Password Generator)
 - Added Generate-RandomPassword function.
 - Added menu entry “_Password Generator”.
 - Added modal with secret password textbox.
 - Added copy-to-clipboard support.
 - Added character-set toggles: upper/lower/numbers/symbols.

 1.6.4.0  (Bugfix)
 - Fixed "password copied" message showing when NOT copying.
 - Fixed textbox not rendering due to incorrect X/Y offsets.
 - Fixed modal stacking order (TreeView was drawing under modals).

 1.6.5.0  (Bugfix)
 - Fixed MessageBox.Query always returning OK due to wrong button array.
 - Corrected typos: “Copie” → “Copied”, “Genertor” → “Generator”.
 - Fixed UI padding around password modal.

 1.6.6.0  (UI & Layout Fixes)
 - Fixed TreeView not anchored properly when window resized.
 - Fixed filter panel stealing focus on startup.
 - Fixed modal shadows not redrawing.

 1.6.7.0  (Bugfix)
 - Corrected menu hotkeys.
 - Fixed password modal height too small on macOS Terminal.
 - Fixed Build-Tree not auto-refreshing after filter changes.

 1.6.8.0  (Stability & AD Query Fixes)
 - Fixed recursive OU building missing final child nodes.
 - Fixed groups sometimes displayed as users due to schema mismatch.
 - Fixed refresh loop running twice on some domains.
 - Added safer DN parsing with fallback.

 1.6.9.0  (Today’s Fixes — Window/Layout Rebuild)
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

 1.7.0.0  (Demo data and menu fixes)
 - Rework demo data to be more AD like cleaning up redundancy
 - Create dedicated about menu
 - Add shortcuts modal

 1.7.0.1  (Demo data expansion - New cities and band)
 - Rendr demo data in a better fashion
 - Add user devices and office printers
 - Add Rocazino form the Koge office
 - Have other locations in Denmark along with different user properites e.g. Alan Wilder

 1.7.1.1  (Fixes for regressions in 1.7.0.1)
 - Clean up entropy box
 - Clean up Demo data tree building
 - Add and Remove button code fixes

1.8.0.0  (Domian Controller Information)
 - Domain controller informaiton (WIP)
 - Domain replication and syncing modal (WIP)
 - Fix tree view right click context menu so it shows now

1.8.1.0  (Verbosity)
 - Add Debug-Log and clean up Debug-Log calls
 - Convert Demo data to use same forma as produciton data which cuts code down significantly
 - clean up outut via Debug-Log and removing dead or unused  dmeo code stanzas

1.8.2.0  (Lumberjack mode)
 - Rework Build-Tree from https://jdhitsolutions.com/blog/active-directory/8173/climbing-trees-in-powershell/
 - Remove duplicatedd Build-Tree code
 - Add updated Theme-Selector modal
 - Additional Debug-Log code printed only when Verbose is true
 - Explain Show-Properties and why that is called instead of e.g. Show-UserProperties

1.8.3.0.0  (More cowbell)
  - Add misisng and former group memebrs of bands
  - Add Get-CleanObjectInfo to reduce code re-use
  - Fix buttons on user properties modal
  - Bring forward $Global code for a more streamlined approach

1.8.4.0  (Refactoring)
  - Refactor code, clean up spacing, unify comments, clean up stanza spacing, etc.
  - Add OU, DC and group propery and editing support where possible
  - Add more Debug code where needed in debug mode
  - Swap out many dialog windows for the pre-exisintg Show-Modal code for to reduce bloat
  - Add certain check modes for Demo data which is still required, but in less areas of code

1.8.6.0 (Colour my life)
  - New themes DSB Danish State Railways and Pan Am Airlines

2.0.0.0  (Multi Domain support with skittles mode)
  - Mutliple domains support
  - Fix Refresh-Data tree crash in menu - it's tied to not chekcing Demo mode
  - Further Code alignment, cleanup and redundancy clear out
  - Add gemstones scotrail and class91 (BR Intercity Swallow livery) themes
  - Rework theme selection to have two columns

2.0.0.1  (Bug fix release)
 - Fix window fill so it's leaving a single line for the status bar
 - Debug log now also writes to a file if asked
 - document- thogugh not yet fix the right click context menu issues (WIP)
 - Modals such as displayname text boxes now 60 chars because Danes and Germans have long names
 - Search user properties now dynamically searches again with addition of textChainged
 - Cosmetic fixes for theme selection to be more intuitive
 - Reflect if account is locked and/or disabled in the search modal too

2.1.5.2  (Computers and refresh debug)
  - Add computers to the mix
  - Allow properties of computers to display
  - Code to give more debug output to track down the refresh bug
  - Close-LoadingDialog $loadingDlg seems to be the source of this as it's not a funciton
  - Search AD properties modal now more refined and supports multiple domains via combobox
  - if the pswritecoloour module is installed fancy up debug output
  - print instrucitons on how to install missing modules
  - statusbar shortcuts are visible and a new "refreshing..." counter status on the right
  - use jukebox.example TLD which is RFC 2606 and RFC 6761 compliant for documentation and testing.
  - Add The Police in Newcastle and Echo and The Bunnymen in Liverpool
  - RFC 2606 — Reserved Second-Level Domains: example.com example.net example.org
  - RFC 6761 — Special-Use Domain Names: example (TLD) invalid localhost test
  - Domain data loading and refresh loads computers
  - New and improved status bar
  - More natural keyboard shortcuts
  
TODO:
 - Refresh Domian data kills the script <-- edging closer to a fix TODO: Test
 - misisng AD module is non fatal BUT if it's not installed, a global needs to not let users do stupid stuff
 - the terminal icons module, likewise if it's there great, if not fall back
  
===========================================================================================
#>

param(
  [switch]$DemoMode,
  [switch]$Logging,
  [string]$LogFile,
  [string]$Domain,  ## User can specify domain
  [ValidateSet("light","dark","matrix","british", "panam", "dsb", "gemstones", "class91", "scotrail" )]
  [string]$Theme
)

## Define the build version, project and code names once only
$Global:ProjectName = "DSA-TUI pwsh dsa.msc TUI"
$Global:FruitName = "Blåbær"
$Global:BuildVersion = "2.1.5.2"

## Debugging
# DEBUG: Check what we received
Write-Host "DemoMode: $DemoMode" -ForegroundColor Yellow
Write-Host "Logging: $Logging" -ForegroundColor Yellow
Write-Host "LogFile: $LogFile" -ForegroundColor Yellow

# Initialize logging
if ($Logging -or $LogFile) {
    Write-Host "Logging condition TRUE" -ForegroundColor Green
    
    if (-not $LogFile) {
        $LogFile = Join-Path $PSScriptRoot "dsa_tui_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        Write-Host "Auto-generated LogFile: $LogFile" -ForegroundColor Green
    }

    # Get full path - use current directory if not absolute path
    if (-not [System.IO.Path]::IsPathRooted($LogFile)) {
        $LogFile = Join-Path (Get-Location).Path $LogFile
    }
    
    ## Logging
    $Global:Logging = $true
    $Global:LogFile = $LogFile
    
    ## Initialize spinner system
    $Global:statusSpinner = @('|', '/', '-', '\')
    $Global:statusSpinnerIndex = 0
    
    # Create a simple dummy object with a Title property
    $Global:StatusItem = [PSCustomObject]@{
    Title = "Initializing..."
}


    Write-Host "Attempting to create log at: $LogFile" -ForegroundColor Green
    
    try {
        $Global:LogStream = [System.IO.StreamWriter]::new($LogFile, $false)
        $Global:LogStream.AutoFlush = $true
        Write-Host "SUCCESS: Log file created at $LogFile" -ForegroundColor Green
        
        # Write test line
        $Global:LogStream.WriteLine("=== Log Started $(Get-Date) ===")

       ## FORCE immediate write to create the file
       $Global:LogStream.WriteLine("=== DSA-TUI Log Started $(Get-Date) ===")
       $Global:LogStream.WriteLine("DemoMode: $DemoMode")
       $Global:LogStream.WriteLine("Theme: $Theme")
       $Global:LogStream.Flush()  # Force flush to disk NOW

    } catch {
        Write-Host "FAILED to create log: $_" -ForegroundColor Red
        $Global:Logging = $false
    }
} else {
    Write-Host "Logging condition FALSE - no logging enabled" -ForegroundColor Red
}

## Get ready for the launch
Write-Host "Starting $($Global:ProjectName) Codename: $($Global:FruitName) v$($Global:BuildVersion) in $(if($DemoMode){'DEMO'}else{'PRODUCTION'}) mode with $Theme theme..."

## For passwords expiring soon
$sevenDaysFileTime = (Get-Date).AddDays(-7).ToFileTime()

## TODO: Add the colourful module and check if it's loaded too, if not, default to "boridng" debug-log
## There exists a powershell icons library and font. Put osmething in, if it isn't thee use old method
## yes, that iwll also owrk wonders for PSMC
function Test-RequiredModule {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [string]$MinimumVersion = $null
    )

    $module = Get-Module -ListAvailable -Name $Name -ErrorAction SilentlyContinue

    if ($module) {
        if ($MinimumVersion -and ($module.Version -lt [version]$MinimumVersion)) {
            Write-Host "⚠️  Module '$Name' found but version is too old. Need $MinimumVersion or later."
            return $false
        }

        # If Write-Colour is available, use it
        if (Get-Module -ListAvailable -Name PSWriteColor) {
            Write-Colour -Text "✔️  Module '$Name' is installed." -Color DarkGreen
            Import-Module $Name -ErrorAction SilentlyContinue
            $Global:HasPSWriteColor = $true

        } else {
            Write-Host "✔️  Module '$Name' is installed."
        }

        return $true
    }
    else {
        Write-Host "❌ Module '$Name' is NOT installed. Please run:"
        Write-Host "    Install-Module -Name $Name"
        return $false
    }
}

## --------------------------{ Module Checks }-------------------------
$ok = $true

## They aren't strictly needed, but we'll go with that for the moment
$ok = $ok -and (Test-RequiredModule -Name "Microsoft.PowerShell.ConsoleGuiTools")

if (-not $ok) {
  Write-Host ""
  Write-Host "❌ One or more required modules are missing. Exiting." -ForegroundColor Red
  exit
}

## Optional Modules
$adAvailable = Test-RequiredModule -Name "ActiveDirectory"
if (-not $adAvailable) {
    Write-Host "⚠️  ActiveDirectory module missing. Falling back to DEMO mode..." -ForegroundColor Yellow
    $Global:DemoMode = $true
}

$wcAvailable = Test-RequiredModule -Name "PSWriteColor"
if (-not $wcAvailable) {
    Write-Host "⚠️  PSWriteColor module missing..." -ForegroundColor Yellow
}

$tiAvailable = Test-RequiredModule -Name "Terminal-Icons"
if (-not $tiAvailable) {
    Write-Host "⚠️  Terminal-Icons module missing..." -ForegroundColor Yellow
}

<#
## -----------------------{ Check For Nerd font }----------------------
# NOTE:
# The most reliable way in Windows/Pwsh is to check whether a Nerd Font
# appears in the installed font list. Most Nerd Fonts have names like:
#   "Cascadia Code NF"
#   "FiraCode Nerd Font"
# etc.

$installedFonts = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name

$nerdFonts = $installedFonts | Where-Object {
    $_ -match "Nerd Font|NF"
}

if ($nerdFonts.Count -gt 0) {
    Write-Host "✔️  Nerd Font detected: $($nerdFonts -join ', ')"
} else {
    Write-Host "⚠️  No Nerd Font detected. Terminal-Icons may not display properly."
    Write-Host "    See: https://ardalis.com/install-nerd-fonts-terminal-icons-pwsh-7-win-11/"
}
#>


## ------------------------- Load Terminal.Gui ------------------------
Write-Host "Checking Terminal.Gui assembly..."
if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Terminal.Gui' })) {
    $mod = Get-Module -ListAvailable Microsoft.PowerShell.ConsoleGuiTools | Select-Object -First 1
    if ($mod) {
        $dll = Join-Path $mod.ModuleBase 'Terminal.Gui.dll'
        if (Test-Path $dll) { Add-Type -Path $dll -ErrorAction Stop; Write-Host "Loaded Terminal.Gui from $dll" } 
        else { Write-Error "Terminal.Gui.dll not found. Install Microsoft.PowerShell.ConsoleGuiTools."; return }
    } else { Write-Error "Microsoft.PowerShell.ConsoleGuiTools module not found."; return }
} else { Write-Host "Terminal.Gui assembly already loaded." }

## Set global demo mode flag immediately
$Global:DemoMode = $DemoMode

## Set global flags immediately after param block - themes here
$Global:ThemeMode = $Theme

Write-Host "Starting $($Global:ProjectName) in $(if($DemoMode){'DEMO'}else{'PRODUCTION'}) mode..."

### now that's done, set up a tree
$Global:tree = [Terminal.Gui.TreeView]::new()

## ------------------------- Globals ------------------------
## At the start of script - GLOBALS section

# Set up forest/domain structure
if ($Global:DemoMode) {
    ## Demo mode - hardcode forest structure
    ## RFC 2606 and 6761 compliant
    $Global:ForestName = "jukebox.example"
    $Global:RootDomain = "example.com"
    $Global:Domains = @('example.com', 'example.net', 'example.org')
    $Global:Sites = @('GLA', 'EDI', 'LND', 'CPH', 'KGE', 'ODE', 'BON', 'BRL', 'MUC', 'NEW', 'LIV')
} else {
    ## Production mode - query AD
    try {
        ## If user specified a domain, query THAT domain's forest
        if ($Domain) {
            Write-Host "Querying domain: $Domain"
            $targetDomain = Get-ADDomain -Server $Domain -ErrorAction Stop
            $forest = Get-ADForest -Server $targetDomain.Forest -ErrorAction Stop
        } else {
            # No domain specified - use current user's domain
            $targetDomain = Get-ADDomain -ErrorAction Stop
            $forest = Get-ADForest -ErrorAction Stop
        }
        
        $Global:ForestName = $forest.Name.Split('.')[0].ToUpper()
        $Global:RootDomain = $forest.RootDomain
        $Global:Domains = $forest.Domains
        $Global:Sites = $forest.Sites | ForEach-Object { $_.Name }
        
        ## Set current domain to what user specified, or root domain
        $Global:CurrentDomain = if ($Domain) { $Domain } else { $Global:RootDomain }
        
    } catch {
        Write-Error "Failed to query domain/forest: $_"
        Write-Error "Make sure you have connectivity to '$Domain' and appropriate permissions"
        
        # Fallback for Azure AD or disconnected
        $Global:ForestName = "DOMAIN"
        $Global:RootDomain = if ($Domain) { $Domain } else { $env:USERDNSDOMAIN }
        $Global:Domains = @($Global:RootDomain)
        $Global:Sites = @()
        $Global:CurrentDomain = $Global:RootDomain
    }
}

$Global:Domain = $Global:CurrentDomain  ## For compatibility

## Initialize other globals
$Global:CurrentDC = $null
$Global:Users = @()
$Global:Groups = @()
$Global:DCs = @()
$Global:ADObjects = @()
$Global:SelectedObjects = @()
$Global:SelectionMode = $false

## Global Search filters:
$Global:FilterOptions = @{
  ShowDisabledUsers = $true
  ShowEnabledUsers  = $true
  ShowLockedUsers    = $true
  ShowGroups        = $true
  ShowDCs           = $true
  ShowComputers     = $true
  ShowOUs           = $true
  NameFilter        = ""
  SortBy            = "Name"
  SortDescending    = $false
}

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

## .----------------------{ Functons Start Here }---------------------.
## | Any functions you add in here. Make sure to keep Chronology when |
## | calling them from inside other funcitons...                      |
## '------------------------------------------------------------------'

## --------------------------{ Debug Logging }-------------------------
function Debug-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Warn','Error','Success')]
        [string]$Type = 'Info'
    )
    $ts = Get-Date -Format 'HH:mm:ss'
    # Emoji + colors for each type
    switch ($Type) {
    'Info'    { $emoji = 'ⓘ'; $color = 'Cyan' }      # Circled i
    'Warn'    { $emoji = '▲'; $color = 'Yellow' }    # Triangle
    'Error'   { $emoji = '✗'; $color = 'Red' }       # X mark
    'Success' { $emoji = '✓'; $color = 'Green' }     # Check mark
    }
    $line = "[$ts] $emoji $Type $Message"
    
    # Show in console when Logging switch is enabled
    if ($Global:Logging) {
        # Use PSWriteColor if available, otherwise fall back to Write-Host
        if ($Global:HasPSWriteColor) {
            try {
                Write-Color -Encoding UTF8 -Text $line -Color $color
            } catch {
                Write-Host $line -ForegroundColor $color  # Fallback
            }
        } else {
            Write-Host $line -ForegroundColor $color
        }
    }
    
    # Write to log file if enabled
    if ($Global:Logging -and $Global:LogStream) {
        try {
            $Global:LogStream.WriteLine($line)
        } catch { }
    }
}

## ----------------------------{ Get Theme }---------------------------
function Get-Theme {
  param([string]$mode)

  # Initialize color schemes and Ensure ColorSchemes are instantiated
  if (-not $globalCs)     { $globalCs     = [Terminal.Gui.ColorScheme]::new() }
  if (-not $mainWindowCs) { $mainWindowCs = [Terminal.Gui.ColorScheme]::new() }

  ## Normalize theme string: lowercase + ASCII
  $mode = $mode.Trim().ToLower()

  ####    $mode = $mode -replace "ae","ae"

  <# Leave me alone for I am documentation

    Adding Themes:
  
    Add an option above in the [ValidateSet() then define a theme below:

    "faxekondi" {
      $globalCs.Normal     <-- Foreground borders and background colour for all modals
      $globalCs.Focus      <-- Foreground and background for menus
      $mainWindowCs.Normal <-- Main opening dialog and foreground text colour
      $mainWindowCs.Focus  <-- Main opening window focus colours foreground nad background
      }

  Valid colours: Black, Blue, Green, Cyan, Red, Magenta, Brown, Gray, DarkGray, BrightBlue,
                 BrightGreen, BrightCyan, BrightRed, BrightMagenta, BrightYellow, White

  Also leave me alone for I am also documentaiton #>

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
    "panam" {
      $globalCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightBlue,[Terminal.Gui.Color]::White)
      $globalCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightBlue)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightBlue)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightBlue,[Terminal.Gui.Color]::White)
      }
    "dsb" {
      $globalCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
      $globalCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::White)
      }
    "gemstones" {
      $globalCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Green,[Terminal.Gui.Color]::White)
      $globalCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::BrightMagenta)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightGreen)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Green,[Terminal.Gui.Color]::White)
      }
    "scotrail" {
      $globalCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::Blue)
      $globalCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Blue,[Terminal.Gui.Color]::White)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::BrightBlue)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
      }
    "class91" {
      $globalCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Black)
      $globalCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::White)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightRed,[Terminal.Gui.Color]::White)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightBlue)
      }
    "default" {
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

## --------------------------{ Apply Colours }-------------------------
function Apply-Theme {
  param(
    [hashtable]$ThemeData,        # expects keys: Global, MainWindow
    [object]$TopLevel,
    [object]$MainWindow,
    [object]$Menu,
    [object]$StatusBar
    )

  if ($null -eq $ThemeData) { return }

  ## --- Global / TopLevel ---
  if ($TopLevel -and $TopLevel.PSObject.Properties.Name -contains 'ColorScheme') { $TopLevel.ColorScheme = $ThemeData.Global }
  ## --- Main window ---
  if ($MainWindow -and $MainWindow.PSObject.Properties.Name -contains 'ColorScheme') { $MainWindow.ColorScheme = $ThemeData.MainWindow }
  ## --- Menu ---
  if ($Menu -and $Menu.PSObject.Properties.Name -contains 'ColorScheme') { $Menu.ColorScheme = $ThemeData.Global }
  ## --- StatusBar ---
  if ($StatusBar -and $StatusBar.PSObject.Properties.Name -contains 'ColorScheme') { $StatusBar.ColorScheme = $ThemeData.Global }

  ## --- Terminal.Gui base colors ---
  [Terminal.Gui.Colors]::Base     = $ThemeData.Global
  [Terminal.Gui.Colors]::Dialog   = $ThemeData.Global
  [Terminal.Gui.Colors]::Menu     = $ThemeData.Global
  [Terminal.Gui.Colors]::Error    = $ThemeData.Global
  [Terminal.Gui.Colors]::TopLevel = $ThemeData.Global
}

## --------------------------{ Debug Themes }--------------------------
## Diagnostics helper to show what's inside a ColorScheme
function Dump-ColorScheme {
  param([Terminal.Gui.ColorScheme]$Scheme)
  if ($null -eq $Scheme) { Debug-Log ("ColorScheme is null") -Type "warn" ; return }
  Debug-Log ("Normal    : $($Scheme.Normal)") -Type "info"
  Debug-Log ("Focus     : $($Scheme.Focus)") -Type "info"
  Debug-Log ("HotNormal : $($Scheme.HotNormal)") -Type "info"
  Debug-Log ("HotFocus  : $($Scheme.HotFocus)") -Type "info"
  Debug-Log ("Disabled  : $($Scheme.Disabled)") -Type "info"
}

## Select theme before proceeding. Save the mode string
$Global:ThemeMode = $Theme

## Get the selected colour scheme
$cs = Get-Theme -mode $Theme
Apply-Theme -ThemeData $themeData -TopLevel $TopLevel -MainWindow $MainWindow -Menu $Menu -Status $StatusBar

## ------------------------{ Show progress bar }-----------------------
## Helper: Show a simple loading/progress dialog with spinner
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
    
    ## CRITICAL FIX: Store frames in GLOBAL so timer thread can access it
    $Global:spinnerFrames = @("|", "/", "-", "\")
    $Global:spinnerFrameIndex = 0
    
    # Timer callback - ONLY access globals and invoke on MainLoop
    $timer = [System.Threading.Timer]::new({
        [Terminal.Gui.Application]::MainLoop.Invoke({
            # Now safely on the UI thread with PowerShell runspace
            $Global:spinnerFrameIndex = ($Global:spinnerFrameIndex + 1) % 4
            $spinner.Text = $Global:spinnerFrames[$Global:spinnerFrameIndex]
            ## Pretend we're polling real data for demo mode os oyu see progressbars
            if ($Global:DemoMode) {
              Start-Sleep -Milliseconds 350   # Adjust to taste
            }
        })
    }, $null, 0, 150)
    
    # Start non-blocking dialog
    [Terminal.Gui.Application]::Begin($dlg)
    
    # Return both dialog and timer so caller can close/stop cleanly
    return [PSCustomObject]@{ Dialog = $dlg; Timer = $timer }
}

## ------------------------{ Close Progressbar }-----------------------
function Show-ThemeSelector {

    # --- Theme list ---
    $themes = @("british","dark","dsb","light","matrix","panam","gemstones","class91","scotrail")

    # Split into two columns
    $half = [math]::Ceiling($themes.Count / 2)
    $leftThemes  = $themes[0..($half-1)]
    $rightThemes = $themes[$half..($themes.Count-1)]

    # --- Determine current theme (case-insensitive) ---
    $currentTheme = $Global:ThemeMode
    Debug-Log ("DEBUG: Global ThemeMode = ${Global:ThemeMode}") -Type "info"
    Debug-Log ("DEBUG: Current theme for selection = ${currentTheme}") -Type "info"

    $currentIndex = -1
    for ($i = 0; $i -lt $themes.Count; $i++) {
        if ($themes[$i].ToLower() -eq $currentTheme.ToLower()) {
            $currentIndex = $i
            break
        }
    }

    Debug-Log ("DEBUG: Index of current theme in $themes array = ${currentIndex}") -Type "info"

    # Calculate which column gets the selection
    $leftSelected  = if ($currentIndex -ge 0 -and $currentIndex -lt $leftThemes.Count) { $currentIndex } else { -1 }
    $rightSelected = if ($currentIndex -ge $leftThemes.Count) { $currentIndex - $leftThemes.Count } else { -1 }

    Debug-Log ("DEBUG: LeftSelected = ${leftSelected}, RightSelected = ${rightSelected}") -Type "info"

    # --- Create dialog ---
    $dlg = [Terminal.Gui.Dialog]::new("Select Theme", 60, 16)

    $lbl = [Terminal.Gui.Label]::new("Choose a color theme:")
    $lbl.X = 2; $lbl.Y = 1
    $dlg.Add($lbl)

    # --- Left column RadioGroup ---
    $rdoLeft = [Terminal.Gui.RadioGroup]::new($leftThemes)
    $rdoLeft.X = 2
    $rdoLeft.Y = 3
    $rdoLeft.SelectedItem = $leftSelected
    $dlg.Add($rdoLeft)

    # --- Right column RadioGroup ---
    $rdoRight = [Terminal.Gui.RadioGroup]::new($rightThemes)
    $rdoRight.X = 32
    $rdoRight.Y = 3
    $rdoRight.SelectedItem = $rightSelected
    $dlg.Add($rdoRight)

    Debug-Log ("DEBUG: rdoLeft.SelectedItem = ${rdoLeft.SelectedItem}, rdoRight.SelectedItem = ${rdoRight.SelectedItem}") -Type "info"

    # --- Sync columns so only one can be selected ---
    $rdoLeft.add_SelectedItemChanged({
        if ($rdoLeft.SelectedItem -ge 0) { $rdoRight.SelectedItem = -1 }
    })
    $rdoRight.add_SelectedItemChanged({
        if ($rdoRight.SelectedItem -ge 0) { $rdoLeft.SelectedItem = -1 }
    })

    # --- Apply Button ---
    $btnApply = [Terminal.Gui.Button]::new("Apply")
    $btnApply.add_Clicked({
        $sel = if ($rdoLeft.SelectedItem -ge 0) {
            $leftThemes[$rdoLeft.SelectedItem]
        } elseif ($rdoRight.SelectedItem -ge 0) {
            $rightThemes[$rdoRight.SelectedItem]
        } else {
            "dark"  # default if nothing selected
        }

        Debug-Log ("DEBUG: Theme selected on Apply = ${sel}") -Type "info"

        Debug-Log ("Switching to theme: ${sel}") -Type "info"
        $script:ThemeMode = $sel

        $newTheme = Get-Theme -mode $sel
        Apply-Theme -ThemeData $newTheme -TopLevel [Terminal.Gui.Application]::Top -MainWindow $win -Menu $menu -Status $StatusBar

        # Force redraw
        $menu.ColorScheme = $newTheme.Global
        $menu.SetNeedsDisplay()
        $win.SetNeedsDisplay()
        [Terminal.Gui.Application]::Top.SetNeedsDisplay()
        [Terminal.Gui.Application]::Refresh()

        Show-Modal "Theme Changed" "Theme changed to: ${sel}"
        [Terminal.Gui.Application]::RequestStop()
    })
    $dlg.AddButton($btnApply)

    # --- Cancel Button ---
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.AddButton($btnCancel)

    # --- Run the dialog ---
    [Terminal.Gui.Application]::Run($dlg)
}

############################### AD FUNCTIONS BELOW HERE ##################################
function Invoke-AD {
    param(
        [scriptblock]$Script,
        [switch]$SuppressError
    )

    try {
        if (-not (Get-Module -Name ActiveDirectory)) {
            Import-Module ActiveDirectory -ErrorAction Stop
        }

        return & $Script
    } catch {
        Debug-Log ("AD call failed: $($_.Exception.Message)") -Type "Error"

        if (-not $SuppressError) {
            Show-Modal "Error" "Active Directory query failed:`n$($_.Exception.Message)"
        }

        # Fallback to demo mode
        $Global:DemoMode = $true
        return $null
    }
}

## ------------------------{ Load Domain Data }------------------------
function Get-ADObjectsByType {
  param([string]$domain)
  $objTypes = @("user","computer","group","organizationalUnit","contact")
  $allObjects = @()
  foreach ($type in $objTypes) {
    try {
      $objs = if ($Global:DemoMode) {
      ## Demo objects already structured
      @()
      } else {
        Get-ADObject -Filter "ObjectClass -eq '$type'" -Server $domain -Properties Name,ObjectClass,DistinguishedName |
        ForEach-Object { @{ Name=$_.Name; Type=$_.ObjectClass; DN=$_.DistinguishedName } }
      }
      $allObjects += $objs
      } catch {
      ## minimal fix: string interpolation of exception object done via ToString()
      Debug-Log ("DEBUG: Failed to enumerate ${type}: $($_.ToString())") -Type "error"
      }
  }
  return $allObjects
}

function Get-CleanObjectInfo {
  param([string]$treeText)  # ← Remove $Global:
    
  Debug-Log ("DEBUG: Get-CleanObjectInfo called with: '$treeText'") -Type "info"  # ← Use $treeText
    
  ## Determine type FIRST (before cleaning)
  $objectType = if ($treeText -like "(U)*") { "user" }  # ← Use $treeText
                elseif ($treeText -like "(DC)*") { "dc" }  # ← Use $treeText
                elseif ($treeText -like "(G)*") { "group" }  # ← Use $treeText
                elseif ($treeText -like "(PC)*") { "computer" }  # ← Use $treeText
                else { "unknown" }
                
  Debug-Log ("DEBUG: Detected type: $objectType") -Type "info"
    
  ## Remove prefixes like "(U) " or "(DC) " or "(G) "
  $cleanName = $treeText -replace '^\([^)]+\)\s*', ''  # ← Use $treeText
  Debug-Log ("DEBUG: After removing prefix: $cleanName") -Type "info"
    
  ## Remove any non-letter, non-space characters from the start (status icons)
  $cleanName = $cleanName -replace '^[^a-zA-Z]+\s*', ''
  Debug-Log ("DEBUG: After removing icons: $cleanName") -Type "info"
    
  ## Extract just the name if it has [SITE] suffix (for DCs)
  if ($cleanName -match '^(.+?)\s+\[.+\]$') {
    $cleanName = $matches[1].Trim()
    Debug-Log ("DEBUG: Extracted name from [SITE] format: $cleanName") -Type "info"
  }
    
  Debug-Log ("DEBUG: Final cleaned name: $cleanName") -Type "info"
    
  return @{
    Type = $objectType
    Name = $cleanName
  }
}

## -----------------------{ Refresh Domain Data }----------------------
function Refresh-Data {
    param(
        [string]$domain = $Global:CurrentDomain,
        [switch]$RebuildTree = $true
    )
    
    Debug-Log ("DEBUG: === Refresh-Data START ===") -Type "info"
    Debug-Log ("DEBUG: Domain: $domain") -Type "info"
    
    try {
        # STEP 1: Load data (this is the ONLY place we care about demo vs prod)
        if ($Global:DemoMode) {
            Update-Status "Loading demo data..." -spinner
            Debug-Log ("DEBUG: Converting demo data...") -Type "info"
            
            $converted = Convert-DataToADObjects `
                -Users $Global:rawUsers `
                -DCs $Global:rawDCs `
                -Groups $Global:rawDemoGroups `
                -Computers $Global:rawComputers `
                -Domain $domain
            
            $Global:Users = $converted.Users
            $Global:DCs = $converted.DCs
            $Global:Groups = $converted.Groups
            $Global:Computers = $converted.Computers
            
        } else {
            Update-Status "Querying Active Directory..." -spinner
            Debug-Log ("DEBUG: Loading from AD...") -Type "info"
            
            Load-DomainData -domain $domain
        }
        
        Debug-Log ("DEBUG: Data loaded - Users: $($Global:Users.Count), Groups: $($Global:Groups.Count)") -Type "success"
        
        # STEP 2: Rebuild tree (same for both demo and prod - just uses PSCustomObjects)
        if ($RebuildTree) {
            Update-Status "Rebuilding tree..." -spinner
            Debug-Log ("DEBUG: Rebuilding tree...") -Type "info"
            
            [Terminal.Gui.Application]::MainLoop.Invoke({
                try {
                    Build-Tree -domain $domain
                    
                    if ($Global:FilterStatusLabel) {
                        Update-FilterStatusLabel -label $Global:FilterStatusLabel
                    }
                    
                    Debug-Log ("DEBUG: Tree rebuilt successfully") -Type "success"
                    
                } catch {
                    Debug-Log ("ERROR: Tree rebuild failed: $($_.Exception.Message)") -Type "error"
                    Update-Status "Tree rebuild failed" -final
                }
            })
        }

        Update-Status "Refresh complete" -final
        Debug-Log ("DEBUG: === Refresh-Data END (SUCCESS) ===") -Type "success"
        return $true
        
    } catch {
        Debug-Log ("ERROR: Refresh failed: $($_.Exception.Message)") -Type "error"
        Update-Status "Refresh failed" -final
        
        [Terminal.Gui.Application]::MainLoop.Invoke({
            Update-Status "Refresh failed: $($_.Exception.Message)" -final
        })
        
        return $false
    }
}

## -----------------------{ Apply Group Changes }----------------------
function Apply-GroupChanges {
    param($group, $fields)

    try {
        if ($Global:DemoMode) {
            # Apply changes to demo object
            $group.Description = $fields.txtDesc.Text.ToString()
            $group.Mail        = $fields.txtEmail.Text.ToString()
            $group.ManagedBy   = $fields.txtManagedBy.Text.ToString()

            # Update raw demo group
            $rawGroup = $Global:rawDemoGroups |
                Where-Object { $_.Name -eq $group.Name } |
                Select-Object -First 1

            if ($rawGroup) {
                $rawGroup.Description = $fields.txtDesc.Text.ToString()
                $rawGroup.Email       = $fields.txtEmail.Text.ToString()
                $rawGroup.ManagedBy   = $fields.txtManagedBy.Text.ToString()
            }

            Debug-Log ("SUCCESS: Group changes applied (demo mode)") -Type "success"
            Show-Modal "Success" "Changes applied successfully (demo mode)"

            Refresh-Data -domain $Global:CurrentDomain
        }
        else {
            # AD mode
            $setParams = @{
                Identity    = $group.Name
                Description = $fields.txtDesc.Text.ToString()
            }

            if ($fields.txtEmail.Text.ToString()) {
                $setParams['Mail'] = $fields.txtEmail.Text.ToString()
            }

            if ($fields.txtManagedBy.Text.ToString()) {
                $setParams['ManagedBy'] = $fields.txtManagedBy.Text.ToString()
            }

            Set-ADGroup @setParams -ErrorAction Stop

            Debug-Log ("SUCCESS: Group changes applied to AD") -Type "success"
            Show-Modal "Success" "Changes applied successfully"

            Refresh-Data -domain $Global:CurrentDomain
        }

        $script:groupChangesMade = $false
        return $true
    }
    catch {
        Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)"
        return $false
    }
}

## ------------------------{ Danske Soda vand }------------------------
## This is a theme now. Danish Fruit soda based fun. Method to the
## madness
function Show-BlaabaerInfo {
  $dlg = [Terminal.Gui.Dialog]::new("Why $($Global:FruitName)? 🫐", 60, 12)
    
  $message = @"
$($Global:ProjectName) is codenamed $Global:FruitName because:

- I was drinking blueberry soda when writing the code
- $($Global:FruitName) is Danish for blueberry
- Føtex sells a rather nice $($Global:FruitName) soda
- Every great project needs a forest-fruit mascot!
"@
    
  $label = [Terminal.Gui.Label]::new(1, 1, $message)
  $dlg.Add($label)
    
  Show-Modal "Why $($Global:FruitName)...? 🫐" $message
}

## ------------------------{ Load Domain Data }------------------------
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

    if ($Logging) { Write-Debug "DEBUG: Loading domain data for: $domain" }

    # If DemoMode already enabled, just log it
    if ($Global:DemoMode) {
        Debug-Log ("Starting $($Global:ProjectName) in DEMO mode...") -Type "info"

    ## ------------------ Define Demo Users ------------------
    $Global:rawUsers = @(
    ## ========== Simple Minds (UK/Scotland/Glasgow) ==========
    @{ Name = 'Jim Kerr'; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds'); Groups = @('Simple Minds','Vocalists'); Title = 'Lead Vocalist'; Email = 'jim.kerr@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Glasgow Office'; Phone = '+44 141 111 1111'; MobilePhone = '+44 7700 111111'; Street = '1 High Street'; City = 'Glasgow'; PostalCode = 'G1 1AA'; Company = 'Example Music Ltd'; Manager = ''; Description = 'Lead vocalist for Simple Minds' },
    @{ Name = 'Charlie Burchill'; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds'); Groups = @('Simple Minds','Guitarists'); Title = 'Lead Guitarist'; Email = 'charlie.b@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Glasgow Office'; Phone = '+44 141 111 1112'; MobilePhone = '+44 7700 111112'; Street = '1 High Street'; City = 'Glasgow'; PostalCode = 'G1 1AA'; Company = 'Example Music Ltd'; Manager = 'Jim Kerr'; Description = 'Guitarist and founding member of Simple Minds' },
    @{ Name = 'Mel Gaynor'; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds'); Groups = @('Simple Minds','Percussion'); Title = 'Drummer'; Email = 'mel.gaynor@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Glasgow Office'; Phone = '+44 141 111 1113'; MobilePhone = '+44 7700 111113'; Street = '1 High Street'; City = 'Glasgow'; PostalCode = 'G1 1AA'; Company = 'Example Music Ltd'; Manager = 'Jim Kerr'; Description = 'Drummer for Simple Minds' },
    @{ Name = 'Mick MacNeil'; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds'); Groups = @('Simple Minds','Musicians','Former Staff'); Title = 'Keyboardist (Former)'; Email = 'mick.macneil@example.com'; Country = 'UK'; Disabled = $true; Locked = $true; MustChangePassword = $false; Department = 'Music'; Office = 'Glasgow Office'; Phone = '+44 141 111 1120'; MobilePhone = '+44 7700 111120'; Street = '20 High Street'; City = 'Glasgow'; PostalCode = 'G1 1AA'; Company = 'Example Music Ltd'; Manager = ''; Description = 'Former keyboardist for Simple Minds (1977-1990)' },
    @{ Name = 'Derek Forbes'; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds'); Groups = @('Simple Minds','Musicians','Former Staff'); Title = 'Bassist (Former)'; Email = 'derek.forbes@example.com'; Country = 'UK'; Disabled = $true; Locked = $true; MustChangePassword = $false; Department = 'Music'; Office = 'Glasgow Office'; Phone = '+44 141 111 1121'; MobilePhone = '+44 7700 111121'; Street = '21 High Street'; City = 'Glasgow'; PostalCode = 'G1 1AA'; Company = 'Example Music Ltd'; Manager = ''; Description = 'Former bassist for Simple Minds (1977-1985)' },

    ## ========== Marillion (UK/Scotland/Edinburgh) ==========
    @{ Name = 'Derek Dick'; OU = @('Locations','UK','Scotland','Edinburgh','Marillion'); Groups = @('Marillion','Vocalists'); Title = 'Lead Vocalist'; Email = 'fish@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $true; Department = 'Music'; Office = 'Edinburgh Office'; Phone = '+44 131 222 2221'; MobilePhone = '+44 7700 222221'; Street = '22 Queens Road'; City = 'Edinburgh'; PostalCode = 'EH1 2BB'; Company = 'Example Music Ltd'; Manager = ''; Description = 'Former lead vocalist (Fish) for Marillion (1981-1988)' },
    @{ Name = 'Steve Rothery'; OU = @('Locations','UK','Scotland','Edinburgh','Marillion'); Groups = @('Marillion','Guitarists'); Title = 'Lead Guitarist'; Email = 'steve.rothery@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Edinburgh Office'; Phone = '+44 131 222 2222'; MobilePhone = '+44 7700 222222'; Street = '22 Queens Road'; City = 'Edinburgh'; PostalCode = 'EH1 2BB'; Company = 'Example Music Ltd'; Manager = 'Derek Dick'; Description = 'Lead guitarist and founding member of Marillion' },
    @{ Name = 'Pete Trewavas'; OU = @('Locations','UK','Scotland','Edinburgh','Marillion'); Groups = @('Marillion','Guitarists'); Title = 'Bassist'; Email = 'pete.trewavas@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Edinburgh Office'; Phone = '+44 131 222 2223'; MobilePhone = '+44 7700 222223'; Street = '22 Queens Road'; City = 'Edinburgh'; PostalCode = 'EH1 2BB'; Company = 'Example Music Ltd'; Manager = 'Derek Dick'; Description = 'Bassist and founding member of Marillion' },
    @{ Name = 'Mark Kelly'; OU = @('Locations','UK','Scotland','Edinburgh','Marillion'); Groups = @('Marillion','Keyboards'); Title = 'Keyboardist'; Email = 'mark.kelly@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Edinburgh Office'; Phone = '+44 131 222 2224'; MobilePhone = '+44 7700 222224'; Street = '22 Queens Road'; City = 'Edinburgh'; PostalCode = 'EH1 2BB'; Company = 'Example Music Ltd'; Manager = 'Derek Dick'; Description = 'Keyboardist and founding member of Marillion' },
    @{ Name = 'Ian Mosley'; OU = @('Locations','UK','Scotland','Edinburgh','Marillion'); Groups = @('Marillion','Percussion'); Title = 'Drummer'; Email = 'ian.mosley@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Edinburgh Office'; Phone = '+44 131 222 2225'; MobilePhone = '+44 7700 222225'; Street = '22 Queens Road'; City = 'Edinburgh'; PostalCode = 'EH1 2BB'; Company = 'Example Music Ltd'; Manager = 'Derek Dick'; Description = 'Drummer for Marillion (joined 1984)' },

    ## ========== Erasure (UK/England/London) ==========
    @{ Name = 'Andy Bell'; OU = @('Locations','UK','England','London','Erasure'); Groups = @('Erasure','Vocalists'); Title = 'Lead Vocalist'; Email = 'andy.bell@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'London Office'; Phone = '+44 20 7000 1111'; MobilePhone = '+44 7700 333333'; Street = '15 River Street'; City = 'London'; PostalCode = 'E1 7AA'; Company = 'Example Music Ltd'; Manager = ''; Description = 'Lead vocalist for Erasure' },
    @{ Name = 'Vince Clarke'; OU = @('Locations','UK','England','London','Erasure'); Groups = @('Erasure','Synth','Keyboards'); Title = 'Synth / Keyboardist'; Email = 'vince.clarke@example.com'; Country = 'UK'; Disabled = $false; Locked = $true; MustChangePassword = $false; Department = 'Music'; Office = 'London Office'; Phone = '+44 20 7000 1112'; MobilePhone = '+44 7700 333334'; Street = '15 River Street'; City = 'London'; PostalCode = 'E1 7AA'; Company = 'Example Music Ltd'; Manager = 'Andy Bell'; Description = 'Synthesizer pioneer - founding member of Depeche Mode and Erasure' },

    ## ========== Depeche Mode (UK/England/London) ==========
    @{ Name = 'Martin Gore'; OU = @('Locations','UK','England','London','Depeche Mode'); Groups = @('Depeche Mode','Guitarists','Keyboards'); Title = 'Guitarist/Keyboardist'; Email = 'martin.gore@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'London Office'; Phone = '+44 20 7000 2221'; MobilePhone = '+44 7700 444441'; Street = '32 Abbey Lane'; City = 'London'; PostalCode = 'EC2 1AA'; Company = 'Example Music Ltd'; Manager = ''; Description = 'Guitarist, keyboardist and primary songwriter for Depeche Mode' },
    @{ Name = 'Dave Gahan'; OU = @('Locations','UK','England','London','Depeche Mode'); Groups = @('Depeche Mode','Vocalists'); Title = 'Lead Vocalist'; Email = 'dave.gahan@example.com'; Country = 'UK'; Disabled = $false; Locked = $false; MustChangePassword = $true; Department = 'Music'; Office = 'London Office'; Phone = '+44 20 7000 2222'; MobilePhone = '+44 7700 444442'; Street = '32 Abbey Lane'; City = 'London'; PostalCode = 'EC2 1AA'; Company = 'Example Music Ltd'; Manager = 'Martin Gore'; Description = 'Lead vocalist for Depeche Mode' },
    @{ Name = 'Alan Wilder'; OU = @('Locations','UK','England','London','Depeche Mode'); Groups = @('Depeche Mode','Keyboards','Percussion'); Title = 'Keyboardist/Drummer'; Email = 'alan.wilder@example.com'; Country = 'UK'; Disabled = $false; Locked = $true; MustChangePassword = $false; Department = 'Music'; Office = 'London Office'; Phone = '+44 20 7000 2224'; MobilePhone = '+44 7700 444444'; Street = '32 Abbey Lane'; City = 'London'; PostalCode = 'EC2 1AA'; Company = 'Example Music Ltd'; Manager = 'Martin Gore'; Description = 'Multi-instrumentalist for Depeche Mode (1982-1995, departed)' },
    @{ Name = 'Andrew Fletcher'; OU = @('Locations','UK','England','London','Depeche Mode'); Groups = @('Depeche Mode','Keyboards'); Title = 'Keyboards/Bass Synth'; Email = 'andrew.fletcher@example.com'; Country = 'UK'; Disabled = $true; Locked = $true; MustChangePassword = $false; Department = 'Music'; Office = 'London Office'; Phone = '+44 20 7000 2223'; MobilePhone = '+44 7700 444443'; Street = '32 Abbey Lane'; City = 'London'; PostalCode = 'EC2 1AA'; Company = 'Example Music Ltd'; Manager = 'Martin Gore'; Description = 'Keyboard and bass synthesizer for Depeche Mode (deceased)' },

    ## ========== TV-2 (Denmark/Copenhagen) ==========
    @{ Name = 'Steffen Brandt'; OU = @('Locations','Denmark','Copenhagen','TV-2'); Groups = @('TV-2','Vocalists','Guitarists'); Title = 'Lead Vocalist / Guitarist'; Email = 'steffen.brandt@example.com'; Country = 'DK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Copenhagen Office'; Phone = '+45 0000 2222'; MobilePhone = '+45 5012 3457'; Street = '1 Raadhuspladsen'; City = 'Copenhagen'; PostalCode = '1550'; Company = 'Example Music ApS'; Manager = ''; Description = 'Frontman of TV-2' },
    @{ Name = 'Hans Erik Lerchenfeldt'; OU = @('Locations','Denmark','Copenhagen','TV-2'); Groups = @('TV-2','Musicians'); Title = 'Bassist'; Email = 'hans.lerchenfeldt@example.com'; Country = 'Denmark'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Copenhagen Office'; Phone = '+45 33 12 3457'; MobilePhone = '+45 20 11 1157'; Street = 'Nørrebrogade 2'; City = 'Copenhagen'; PostalCode = '2200'; Company = 'Example Music Ltd'; Manager = 'Steffen Brandt'; Description = 'Bassist for TV-2' },
    @{ Name = 'Sven Gaul'; OU = @('Locations','Denmark','Copenhagen','TV-2'); Groups = @('TV-2','Musicians'); Title = 'Drummer'; Email = 'sven.gaul@example.com'; Country = 'Denmark'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Copenhagen Office'; Phone = '+45 33 12 3458'; MobilePhone = '+45 20 11 1158'; Street = 'Nørrebrogade 3'; City = 'Copenhagen'; PostalCode = '2200'; Company = 'Example Music Ltd'; Manager = 'Steffen Brandt'; Description = 'Drummer for TV-2' },
    @{ Name = 'Georg Olesen'; OU = @('Locations','Denmark','Aarhus','TV-2'); Groups = @('TV-2','Musicians','Former Staff'); Title = 'Guitarist (Former)'; Email = 'georg.olesen@example.com'; Country = 'Denmark'; Disabled = $true; Locked = $true; MustChangePassword = $false; Department = 'Music'; Office = 'Aarhus Office'; Phone = '+45 86 12 3470'; MobilePhone = '+45 20 11 1170'; Street = 'Åboulevarden 20'; City = 'Aarhus'; PostalCode = '8000'; Company = 'Example Music Ltd'; Manager = ''; Description = 'Former guitarist and co-founder of TV-2 (1981-2003)' },

    ## ========== Rocazino (Denmark/Koge) ==========
    @{ Name = 'Ulla Kjaer'; OU = @('Locations','Denmark','Koge','Rocazino'); Groups = @('Rocazino','Vocalists'); Title = 'Lead Vocalist'; Email = 'ulla.kjaer@example.com'; Country = 'DK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Koge Office'; Phone = '+45 0000 2234'; MobilePhone = '+45 3012 3456'; Street = '7 Torvet'; City = 'Koge'; PostalCode = '4600'; Company = 'Example Music ApS'; Manager = ''; Description = 'Lead vocalist for Rocazino' },
    @{ Name = 'Michael Bruun'; OU = @('Locations','Denmark','Koge','Rocazino'); Groups = @('Rocazino','Guitarists'); Title = 'Guitarist'; Email = 'michael.bruun@example.com'; Country = 'DK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Koge Office'; Phone = '+45 0000 2235'; MobilePhone = '+45 3012 3457'; Street = '7 Torvet'; City = 'Koge'; PostalCode = '4600'; Company = 'Example Music ApS'; Manager = 'Ulla Kjaer'; Description = 'Guitarist and songwriter for Rocazino' },
    @{ Name = 'Jan Sivertsen'; OU = @('Locations','Denmark','Koge','Rocazino'); Groups = @('Rocazino','Percussion'); Title = 'Drummer'; Email = 'jan.sivertsen@example.com'; Country = 'DK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Koge Office'; Phone = '+45 0000 2236'; MobilePhone = '+45 3012 3458'; Street = '7 Torvet'; City = 'Koge'; PostalCode = '4600'; Company = 'Example Music ApS'; Manager = 'Ulla Kjaer'; Description = 'Drummer for Rocazino' },

    ## ========== MØ (Denmark/Odense) ==========
    @{ Name = 'Karen Marie Orsted'; OU = @('Locations','Denmark','Odense','Mo'); Groups = @('Mo','Vocalists'); Title = 'Singer / Songwriter'; Email = 'karen.orsted@example.com'; Country = 'DK'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Odense Office'; Phone = '+45 0000 3234'; MobilePhone = '+45 4012 3456'; Street = '22 Vestergade'; City = 'Odense'; PostalCode = '5000'; Company = 'Example Music ApS'; Manager = ''; Description = 'Danish singer-songwriter known internationally as MØ' }

    ## ========== Kraftwerk (West Germany/Bonn) ==========
    @{ Name = 'Ralf Hutter'; OU = @('Locations','Germany','Bonn','Kraftwerk'); Groups = @('Kraftwerk','Vocalists','Musicians'); Title = 'Vocals/Synthesizer'; Email = 'ralf.hutter@example.net'; Country = 'Germany'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Bonn Office'; Phone = '+49 228 1111 111'; MobilePhone = '+49 170 1111111'; Street = 'Adenauerallee 1'; City = 'Bonn'; PostalCode = '53113'; Company = 'Example Music GmbH'; Manager = ''; Description = 'Co-founder and frontman of Kraftwerk' },
    @{ Name = 'Florian Schneider'; OU = @('Locations','Germany','Bonn','Kraftwerk'); Groups = @('Kraftwerk','Musicians'); Title = 'Synthesizer/Flute'; Email = 'florian.schneider@example.net'; Country = 'Germany'; Disabled = $true; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Bonn Office'; Phone = '+49 228 1111 112'; MobilePhone = '+49 170 1111112'; Street = 'Adenauerallee 2'; City = 'Bonn'; PostalCode = '53113'; Company = 'Example Music GmbH'; Manager = 'Ralf Hutter'; Description = 'Co-founder of Kraftwerk (1947-2020) - Account disabled' },
    @{ Name = 'Wolfgang Flur'; OU = @('Locations','Germany','Bonn','Kraftwerk'); Groups = @('Kraftwerk','Musicians','Percussionists'); Title = 'Electronic Drums'; Email = 'wolfgang.flur@example.net'; Country = 'Germany'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Bonn Office'; Phone = '+49 228 1111 113'; MobilePhone = '+49 170 1111113'; Street = 'Adenauerallee 3'; City = 'Bonn'; PostalCode = '53113'; Company = 'Example Music GmbH'; Manager = 'Ralf Hutter'; Description = 'Electronic percussionist for Kraftwerk' },
    @{ Name = 'Karl Bartos'; OU = @('Locations','Germany','Bonn','Kraftwerk'); Groups = @('Kraftwerk','Musicians','Percussionists'); Title = 'Electronic Percussion'; Email = 'karl.bartos@example.net'; Country = 'Germany'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Bonn Office'; Phone = '+49 228 1111 114'; MobilePhone = '+49 170 1111114'; Street = 'Adenauerallee 4'; City = 'Bonn'; PostalCode = '53113'; Company = 'Example Music GmbH'; Manager = 'Ralf Hutter'; Description = 'Electronic percussionist and composer for Kraftwerk' },

    ## ==========  Nena - Germany/West Berlin ==========
    @{ Name = 'Gabriele Kerner'; OU = @('Locations','Germany','West Berlin','Nena'); Groups = @('Nena','Vocalists'); Title = 'Lead Vocalist'; Email = 'nena@example.net'; Country = 'Germany'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Berlin Office'; Phone = '+49 30 2222 111'; MobilePhone = '+49 171 2221111'; Street = 'Kurfurstendamm 100'; City = 'Berlin'; PostalCode = '10709'; Company = 'Example Music GmbH'; Manager = ''; Description = 'Lead vocalist of Nena, known as Nena' },
    @{ Name = 'Carlo Karges'; OU = @('Locations','Germany','West Berlin','Nena'); Groups = @('Nena','Musicians'); Title = 'Guitarist'; Email = 'carlo.karges@example.net'; Country = 'Germany'; Disabled = $true; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Berlin Office'; Phone = '+49 30 2222 112'; MobilePhone = '+49 171 2221112'; Street = 'Kurfurstendamm 101'; City = 'Berlin'; PostalCode = '10709'; Company = 'Example Music GmbH'; Manager = ''; Description = 'Guitarist for Nena (1951-2002) - Account disabled' },
    @{ Name = 'Uwe Fahrenkrog-Petersen'; OU = @('Locations','Germany','West Berlin','Nena'); Groups = @('Nena','Musicians'); Title = 'Keyboardist'; Email = 'uwe.fahrenkrog@example.net'; Country = 'Germany'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Berlin Office'; Phone = '+49 30 2222 113'; MobilePhone = '+49 171 2221113'; Street = 'Kurfurstendamm 102'; City = 'Berlin'; PostalCode = '10709'; Company = 'Example Music GmbH'; Manager = 'Gabriele Kerner'; Description = 'Keyboardist and songwriter for Nena' },
    @{ Name = 'Jurgen Dehmel'; OU = @('Locations','Germany','West Berlin','Nena'); Groups = @('Nena','Musicians'); Title = 'Bassist'; Email = 'jurgen.dehmel@example.net'; Country = 'Germany'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Berlin Office'; Phone = '+49 30 2222 114'; MobilePhone = '+49 171 2221114'; Street = 'Kurfurstendamm 103'; City = 'Berlin'; PostalCode = '10709'; Company = 'Example Music GmbH'; Manager = 'Gabriele Kerner'; Description = 'Bassist for Nena' },
    @{ Name = 'Rolf Brendel'; OU = @('Locations','Germany','West Berlin','Nena'); Groups = @('Nena','Musicians','Percussionists'); Title = 'Drummer'; Email = 'rolf.brendel@example.net'; Country = 'Germany'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Berlin Office'; Phone = '+49 30 2222 115'; MobilePhone = '+49 171 2221115'; Street = 'Kurfurstendamm 104'; City = 'Berlin'; PostalCode = '10709'; Company = 'Example Music GmbH'; Manager = 'Gabriele Kerner'; Description = 'Drummer for Nena' },

    ## ========== Tangerine Dream - Germany/West Berlin ==========
    @{ Name = 'Edgar Froese'; OU = @('Locations','Germany','West Berlin','Tangerine Dream'); Groups = @('Tangerine Dream','Musicians'); Title = 'Synthesizer/Guitar'; Email = 'edgar.froese@example.net'; Country = 'Germany'; Disabled = $true; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Berlin Office'; Phone = '+49 30 2222 121'; MobilePhone = '+49 171 2221121'; Street = 'Kantstrasse 50'; City = 'Berlin'; PostalCode = '10625'; Company = 'Example Music GmbH'; Manager = ''; Description = 'Founder of Tangerine Dream (1944-2015) - Account disabled' },
    @{ Name = 'Christopher Franke'; OU = @('Locations','Germany','West Berlin','Tangerine Dream'); Groups = @('Tangerine Dream','Musicians'); Title = 'Synthesizer/Drums'; Email = 'christopher.franke@example.net'; Country = 'Germany'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Berlin Office'; Phone = '+49 30 2222 122'; MobilePhone = '+49 171 2221122'; Street = 'Kantstrasse 51'; City = 'Berlin'; PostalCode = '10625'; Company = 'Example Music GmbH'; Manager = ''; Description = 'Synthesizer and electronic drums for Tangerine Dream' },
    @{ Name = 'Peter Baumann'; OU = @('Locations','Germany','West Berlin','Tangerine Dream'); Groups = @('Tangerine Dream','Musicians'); Title = 'Synthesizer'; Email = 'peter.baumann@example.net'; Country = 'Germany'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Berlin Office'; Phone = '+49 30 2222 123'; MobilePhone = '+49 171 2221123'; Street = 'Kantstrasse 52'; City = 'Berlin'; PostalCode = '10625'; Company = 'Example Music GmbH'; Manager = ''; Description = 'Synthesizer player for Tangerine Dream' },

    # ========== Fancy - Germany/Munich ==========
    @{ Name = 'Manfred Segieth'; OU = @('Locations','Germany','Bayern','Munich','Fancy'); Groups = @('Fancy Solo','Vocalists'); Title = 'Solo Artist'; Email = 'fancy@example.net'; Country = 'Germany'; Disabled = $false; Locked = $false; MustChangePassword = $false; Department = 'Music'; Office = 'Munich Office'; Phone = '+49 89 3333 111'; MobilePhone = '+49 172 3331111'; Street = 'Leopoldstrasse 100'; City = 'Munich'; PostalCode = '80802'; Company = 'Example Music GmbH'; Manager = ''; Description = 'Solo artist Fancy from Munich, known for Lady of Ice and Italo disco hits' }

    # ========== Echo And The Bunnymen - UK/England/Liverpool ==========
    @{ Name = 'Ian McCulloch' ; OU = @('Locations','UK','England','Liverpool','Echo and The Bunnymen') ; Groups = @('Echo and The Bunnymen','Vocalists') ; Title = 'Lead Vocalist' ; Email = 'ian.mcculloch@example.org' ; Country = 'United Kingdom' ; Disabled = $false ; Locked = $false ; MustChangePassword = $true ; Department = 'Music' ; Office = 'Liverpool Office' ; Phone = '+44 151 555 1001' ; MobilePhone = '+44 7700 900001' ; Street = 'Mathew Street 12' ; City = 'Liverpool' ; PostalCode = 'L1 4ED' ; Company = 'Example Music UK' ; Manager = '' ; Description = 'Lead vocalist of Echo and The Bunnymen' }
    @{ Name = 'Will Sergeant' ; OU = @('Locations','UK','England','Liverpool','Echo and The Bunnymen') ; Groups = @('Echo and The Bunnymen','Guitarists') ; Title = 'Guitarist' ; Email = 'will.sergeant@example.org' ; Country = 'United Kingdom' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Liverpool Office' ; Phone = '+44 151 555 1002' ; MobilePhone = '+44 7700 900002' ; Street = 'Bold Street 22' ; City = 'Liverpool' ; PostalCode = 'L1 4HR' ; Company = 'Example Music UK' ; Manager = 'Ian McCulloch' ; Description = 'Guitarist for Echo and The Bunnymen' }
    @{ Name = 'Les Pattinson' ; OU = @('Locations','UK','England','Liverpool','Echo and The Bunnymen') ; Groups = @('Echo and The Bunnymen','Bassists') ; Title = 'Bass Guitarist' ; Email = 'les.pattinson@example.org' ; Country = 'United Kingdom' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Liverpool Office' ; Phone = '+44 151 555 1003' ; MobilePhone = '+44 7700 900003' ; Street = 'Seel Street 8' ; City = 'Liverpool' ; PostalCode = 'L1 4BE' ; Company = 'Example Music UK' ; Manager = 'Ian McCulloch' ; Description = 'Bass guitarist for Echo and The Bunnymen' }
    @{ Name = 'Pete de Freitas' ; OU = @('Locations','UK','England','Liverpool','Echo and The Bunnymen') ; Groups = @('Echo and The Bunnymen','Drummers') ; Title = 'Drummer' ; Email = 'pete.defreitas@example.org' ; Country = 'United Kingdom' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Liverpool Office' ; Phone = '+44 151 555 1004' ; MobilePhone = '+44 7700 900004' ; Street = 'Dale Street 5' ; City = 'Liverpool' ; PostalCode = 'L2 2EH' ; Company = 'Example Music UK' ; Manager = 'Ian McCulloch' ; Description = 'Drummer for Echo and The Bunnymen' }

    @{ Name = 'Sting' ; OU = @('Locations','UK','England','Newcastle','The Police') ; Groups = @('The Police','Vocalists','Bassists') ; Title = 'Lead Vocalist & Bassist' ; Email = 'sting@example.org' ; Country = 'United Kingdom' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Newcastle Office' ; Phone = '+44 191 555 2001' ; MobilePhone = '+44 7700 910001' ; Street = 'Grey Street 14' ; City = 'Newcastle' ; PostalCode = 'NE1 6BH' ; Company = 'Example Music UK' ; Manager = '' ; Description = 'Lead vocalist and bassist of The Police' }
    @{ Name = 'Andy Summers' ; OU = @('Locations','UK','England','Newcastle','The Police') ; Groups = @('The Police','Guitarists') ; Title = 'Guitarist' ; Email = 'andy.summers@example.org' ; Country = 'United Kingdom' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Newcastle Office' ; Phone = '+44 191 555 2002' ; MobilePhone = '+44 7700 910002' ; Street = 'Dean Street 11' ; City = 'Newcastle' ; PostalCode = 'NE1 1PG' ; Company = 'Example Music UK' ; Manager = 'Sting' ; Description = 'Guitarist for The Police' }
    @{ Name = 'Stewart Copeland' ; OU = @('Locations','UK','England','Newcastle','The Police') ; Groups = @('The Police','Drummers') ; Title = 'Drummer' ; Email = 'stewart.copeland@example.org' ; Country = 'United Kingdom' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Newcastle Office' ; Phone = '+44 191 555 2003' ; MobilePhone = '+44 7700 910003' ; Street = 'Collingwood Street 5' ; City = 'Newcastle' ; PostalCode = 'NE1 1JF' ; Company = 'Example Music UK' ; Manager = 'Sting' ; Description = 'Drummer for The Police' }

    ## Demo Users ending stanza
    )

    ## ------------------ Define Demo Groups ------------------
  $Global:rawDemoGroups = @(
    @{ Name = 'Simple Minds' ; Description = 'Scottish rock band formed in Glasgow in 1977' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Jim Kerr' ; Email = 'simpleminds@example.com' },
    @{ Name = 'Depeche Mode' ; Description = 'English electronic music band formed in Basildon in 1980' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Dave Gahan' ; Email = 'depechemode@example.com' },
    @{ Name = 'Erasure' ; Description = 'English synth-pop duo formed in London in 1985' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Andy Bell' ; Email = 'erasure@example.com' },
    @{ Name = 'Marillion' ; Description = 'British rock band formed in Edinburgh in 1979' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Steve Hogarth' ; Email = 'marillion@example.com' },
    @{ Name = 'TV-2' ; Description = 'Danish rock band formed in 1981' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Steffen Brandt' ; Email = 'tv2@example.com' },
    @{ Name = 'Rocazino' ; Description = 'Danish pop band from Koge' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Stine Bramsen' ; Email = 'alphabeat@example.com' },
    @{ Name = 'MØ Solo' ; Description = 'Solo artist MØ from Odense' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Karen Marie Ørsted' ; Email = 'mo@example.com' },
    @{ Name = 'Kraftwerk' ; Description = 'German electronic music pioneers from Bonn, formed 1970' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Ralf Hutter' ; Email = 'kraftwerk@example.net' },
    @{ Name = 'Nena' ; Description = 'German Neue Deutsche Welle band from West Berlin, formed 1982' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Gabriele Kerner' ; Email = 'nena@example.net' },
    @{ Name = 'Tangerine Dream' ; Description = 'German electronic music group from West Berlin, formed 1967' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Christopher Franke' ; Email = 'tangerinedream@example.net' },
    @{ Name = 'Fancy Solo' ; Description = 'Solo artist Fancy from Munich, Italo disco performer' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Manfred Segieth' ; Email = 'fancy@example.net' }
    @{ Name = 'bunnymen' ; Description = 'Echo And The Bunnymen Groupr' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Ian McCulloch' ; Email = 'bunnymen@example.net' }
    @{ Name = 'The Police' ; Description = 'Group Catch-all for The Police' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Sting' ; Email = 'thepolice@example.net' }
    @{ Name = 'Vocalists' ; Description = 'Lead singers and vocalists across all bands' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = '' ; Email = '' },
    @{ Name = 'Keyboards' ; Description = 'Keyboard And synthesizers' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = '' ; Email = '' },
    @{ Name = 'Musicians' ; Description = 'Instrumentalists - guitarists, bassists, drummers, keyboardists' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = '' ; Email = '' },
    @{ Name = 'Guitarists' ; Description = 'Guitar and bass players' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = '' ; Email = '' },
    @{ Name = 'Percussionists' ; Description = 'Drummers and percussion specialists' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = '' ; Email = '' },
    @{ Name = 'Former Staff' ; Description = 'Former band members and staff' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = '' ; Email = '' },
    @{ Name = 'Disabled Users' ; Description = 'Users with administratively disabled accounts' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = '' ; Email = '' }
  )

  ## ------------------ Define Demo Domain Controllers ------------------
  $Global:rawDCs = @(
    @{ Name = 'EXAGLADC01' ; Site = 'GLA' ; Location = 'Glasgow, Scotland' ; OS = 'Windows Server 2022 Standard' ; IPAddress = '192.168.4.20' ; IsGlobalCatalog = $true ; FSMORoles = @('Schema Master', 'Domain Naming Master', 'PDC Emulator') ; LastReplication = (Get-Date).AddMinutes(-12) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-45) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{  'C:' = @{Total = '120 GB'; Used = '45 GB'; Free = '75 GB'; PercentFree = 62} ; 'SYSVOL' = @{Total = '50 GB'; Used = '8 GB'; Free = '42 GB'; PercentFree = 84} } ; ReplicationPartners = @('EXALNDDC01', 'EXAEDIDC01') },
    @{ Name = 'EXAEDIDC01' ; Site = 'EDI' ; Location = 'Edinburgh, Scotland' ; OS = 'Windows Server 2019 Standard' ; IPAddress = '192.168.3.20' ; IsGlobalCatalog = $true ; FSMORoles = @() ; LastReplication = (Get-Date).AddMinutes(-18) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-67) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '38 GB'; Free = '82 GB'; PercentFree = 68} ; 'SYSVOL' = @{Total = '50 GB'; Used = '6 GB'; Free = '44 GB'; PercentFree = 88} } ; ReplicationPartners = @('EXAGLADC01', 'EXALNDDC01','EXANEWDC01', 'EXALIVDC01') },
    @{ Name = 'EXALNDDC01' ; Site = 'LND' ; Location = 'London, England' ; OS = 'Windows Server 2022 Standard' ; IPAddress = '192.168.2.20' ; IsGlobalCatalog = $true ; FSMORoles = @('RID Master', 'Infrastructure Master') ; LastReplication = (Get-Date).AddMinutes(-8) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-23) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '52 GB'; Free = '68 GB'; PercentFree = 57} ; 'SYSVOL' = @{Total = '50 GB'; Used = '12 GB'; Free = '38 GB'; PercentFree = 76} } ; ReplicationPartners = @('EXAGLADC01', 'EXAEDIDC01', 'EXAKGEDC01', 'EXACPHDC01') },
    @{ Name = 'EXACPHDC01' ; Site = 'CPH' ; Location = 'Copenhagen, Denmark' ; OS = 'Windows Server 2022 Standard' ; IPAddress = '192.168.6.20' ; IsGlobalCatalog = $true ; FSMORoles = @() ; LastReplication = (Get-Date).AddMinutes(-10) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-34) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '41 GB'; Free = '79 GB'; PercentFree = 66} ; 'SYSVOL' = @{Total = '50 GB'; Used = '9 GB'; Free = '41 GB'; PercentFree = 82} } ; ReplicationPartners = @('EXALNDDC01', 'EXAKGEDC01') },
    @{ Name = 'EXAKGEDC01' ; Site = 'KGE' ; Location = 'Køge, Denmark' ; OS = 'Windows Server 2016 Standard' ; IPAddress = '192.168.5.20' ; IsGlobalCatalog = $false ; FSMORoles = @() ; LastReplication = (Get-Date).AddHours(-4).AddMinutes(-35) ; ReplicationHealth = 'Warning - Out of Sync' ; LastBoot = (Get-Date).AddDays(-156) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '100 GB'; Used = '78 GB'; Free = '22 GB'; PercentFree = 22} ; 'SYSVOL' = @{Total = '40 GB'; Used = '28 GB'; Free = '12 GB'; PercentFree = 30} } ; ReplicationPartners = @('EXALNDDC01', 'EXACPHDC01') },
    @{ Name = 'EXAODEDC01' ; Site = 'ODE' ; Location = 'Odense, Denmark' ; OS = 'Windows Server 2022 Standard' ; IPAddress = '192.168.7.20' ; IsGlobalCatalog = $true ; FSMORoles = @() ; LastReplication = (Get-Date).AddMinutes(-14) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-28) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '42 GB'; Free = '78 GB'; PercentFree = 65}; 'SYSVOL' = @{Total = '50 GB'; Used = '7 GB'; Free = '43 GB'; PercentFree = 86} } ; ReplicationPartners = @('EXACPHDC01', 'EXAKGEDC01') },
    @{ Name = 'EXABONDC01' ; Site = 'BON' ; Location = 'Bonn, Germany' ; OS = 'Windows Server 2022 Standard' ; IPAddress = '192.168.22.20' ; IsGlobalCatalog = $true ; FSMORoles = @('Schema Master', 'Domain Naming Master') ; LastReplication = (Get-Date).AddMinutes(-11) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-38) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '48 GB'; Free = '72 GB'; PercentFree = 60} ; 'SYSVOL' = @{Total = '50 GB'; Used = '10 GB'; Free = '40 GB'; PercentFree = 80} } ; ReplicationPartners = @('EXABRLDC01') },
    @{ Name = 'EXABRLDC01' ; Site = 'BRL' ; Location = 'West Berlin, Germany' ; OS = 'Windows Server 2019 Standard' ; IPAddress = '192.168.30.20' ; IsGlobalCatalog = $true ; FSMORoles = @('PDC Emulator', 'RID Master', 'Infrastructure Master') ; LastReplication = (Get-Date).AddMinutes(-9) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-51) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '55 GB'; Free = '65 GB'; PercentFree = 54} ; 'SYSVOL' = @{Total = '50 GB'; Used = '14 GB'; Free = '36 GB'; PercentFree = 72} } ; ReplicationPartners = @('EXABONDC01', 'EXAMUCDC01') },
    @{ Name = 'EXAMUCDC01' ; Site = 'MUC' ; Location = 'Munich, Germany' ; OS = 'Windows Server 2022 Standard' ; IPAddress = '192.168.89.20' ; IsGlobalCatalog = $true ; FSMORoles = @() ; LastReplication = (Get-Date).AddMinutes(-13) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-42) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '44 GB'; Free = '76 GB'; PercentFree = 63} ; 'SYSVOL' = @{Total = '50 GB'; Used = '8 GB'; Free = '42 GB'; PercentFree = 84} ; } ; ReplicationPartners = @('EXABONDC01', 'EXABRLDC01') }    
    @{ Name = 'EXANEWDC01' ; Site = 'NEW' ; Location = 'Newcastle, UK' ; OS = 'Windows Server 2022 Standard' ; IPAddress = '192.168.91.20' ; IsGlobalCatalog = $true ; FSMORoles = @() ; LastReplication = (Get-Date).AddMinutes(-13) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-42) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '24 GB'; Free = '96 GB'; PercentFree = 63} ; 'SYSVOL' = @{Total = '50 GB'; Used = '8 GB'; Free = '42 GB'; PercentFree = 84} ; } ; ReplicationPartners = @('EXALIVDC01', 'EXAEDIDC01', 'EXAGLADC01', 'EXALNDDC01' ) }
    @{ Name = 'EXALIVDC01' ; Site = 'LIV' ; Location = 'Liverpool, UK' ; OS = 'Windows Server 2025 Advanced' ; IPAddress = '192.168.51.20' ; IsGlobalCatalog = $true ; FSMORoles = @() ; LastReplication = (Get-Date).AddMinutes(-13) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-42) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '46 GB'; Free = '76 GB'; PercentFree = 63} ; 'SYSVOL' = @{Total = '50 GB'; Used = '8 GB'; Free = '42 GB'; PercentFree = 84} ; } ; ReplicationPartners = @('EXABONDC01', 'EXABRLDC01') }
  )

  ## --------{ Demo Computers (Workstations, Laptops, Printers) }--------
  $Global:rawComputers = @(
    @{ Name = 'GLA-WS-001' ; Type = 'Workstation' ; Location = 'Glasgow, Scotland' ; OU = @('Locations','UK','Scotland','Glasgow','Computers') ; OS = 'Windows 11 Pro' ; OSVersion = '10.0.22631' ; LastLogon = (Get-Date).AddHours(-2) ; IPv4Address = '192.168.4.50' ; Description = 'Hot desk workstation - Glasgow office'; Enabled = $true },
    @{ Name = 'GLA-WS-002' ; Type = 'Workstation' ; Location = 'Glasgow, Scotland' ; OU = @('Locations','UK','Scotland','Glasgow','Computers') ; OS = 'Windows 11 Pro' ; OSVersion = '10.0.22631' ; LastLogon = (Get-Date).AddDays(-1) ; IPv4Address = '192.168.4.51' ; Description = 'Hot desk workstation - Glasgow office' ; Enabled = $true },
    @{ Name = 'GLA-LT-001' ; Type = 'Laptop' ; Location = 'Glasgow, Scotland' ; OU = @('Locations','UK','Scotland','Glasgow','Computers') ; OS = 'Windows 11 Pro' ; OSVersion = '10.0.22631' ; LastLogon = (Get-Date).AddHours(-5) ; IPv4Address = '192.168.4.80' ; Description = 'Dell Latitude laptop - Pool device' ; Enabled = $true },
    @{ Name = 'GLA-PR-001' ; Type = 'Printer' ; Location = 'Glasgow, Scotland' ; OU = @('Locations','UK','Scotland','Glasgow','Computers') ; OS = 'Printer' ; OSVersion = 'N/A' ; LastLogon = (Get-Date).AddMinutes(-30) ; IPv4Address = '192.168.4.100' ; Description = 'HP LaserJet Pro - Main floor' ; Enabled = $true }, 
    @{ Name = 'LND-WS-001' ; Type = 'Workstation' ; Location = 'London, England' ; OU = @('Locations','UK','England','London','Computers') ; OS = 'Windows 11 Pro' ; OSVersion = '10.0.22631' ; LastLogon = (Get-Date).AddHours(-1) ; IPv4Address = '192.168.2.50' ; Description = 'Hot desk workstation - London office' ; Enabled = $true },
    @{ Name = 'LND-PR-001' ; Type = 'Printer' ; Location = 'London, England' ; OU = @('Locations','UK','England','London','Computers') ; OS = 'Printer' ; OSVersion = 'N/A' ; LastLogon = (Get-Date).AddMinutes(-15) ; IPv4Address = '192.168.2.100' ; Description = 'Xerox WorkCentre - Reception' ; Enabled = $true },
    @{ Name = 'BON-WS-001' ; Type = 'Workstation' ; Location = 'Bonn, Germany' ; OU = @('Locations','Germany','Bonn','Computers') ; OS = 'Windows 11 Pro' ; OSVersion = '10.0.22631' ; LastLogon = (Get-Date).AddHours(-3) ; IPv4Address = '192.168.22.50' ; Description = 'Hot desk workstation - Bonn office' ; Enabled = $true },
    @{ Name = 'BON-LT-001' ; Type = 'Laptop' ; Location = 'Bonn, Germany' ; OU = @('Locations','Germany','Bonn','Computers') ; OS = 'Windows 11 Pro' ; OSVersion = '10.0.22631' ; LastLogon = (Get-Date).AddDays(-2) ; IPv4Address = '192.168.22.75' ; Description = 'Lenovo ThinkPad - Pool device' ; Enabled = $false  } ## Disabled for maintenance
    @{ Name = 'LIV-MBP-001' ; Type = 'Macbook' ; Location = 'Liverpool, UK' ; OU = @('Locations','UK','Liverpool','Computers') ; OS = 'MacOS Tahoe' ; OSVersion = '26.0.1' ; LastLogon = (Get-Date).AddDays(-14) ; IPv4Address = '192.168.51.75' ; Description = 'Macbook Pro 2024 - Pool device' ; Enabled = $true  }
    @{ Name = 'LIV-MAC-001' ; Type = 'iMac' ; Location = 'Newcastle, UK' ; OU = @('Locations','UK','Newcastle','Computers') ; OS = 'MacOS Tahoe' ; OSVersion = '26.0.1' ; LastLogon = (Get-Date).AddDays(-40) ; IPv4Address = '192.168.91.75' ; Description = 'Macbook Pro 2024 - Pool device' ; Enabled = $false  } ## Disabled for maintenance
    
  )

  ## Load demo data
  $converted = Convert-DataToADObjects -Users $Global:rawUsers -DCs $Global:rawDCs -Groups $Global:rawDemoGroups -Computers $Global:rawComputers -Domain $Global:CurrentDomain
  return
  }
  
  ## Production
## Attempt to load ActiveDirectory module and fetch data
try {
    Debug-Log "Attempting to load production AD data for domain: $Global:CurrentDomain" -Type "info"
    
    # Get Domain Controllers
    $dcs = Invoke-AD {
        Get-ADDomainController -Filter * -Server $Global:CurrentDomain -ErrorAction Stop
    } -SuppressError
    
    # Get Users
    $users = Invoke-AD {
        Get-ADUser -Filter * -Server $Global:CurrentDomain -Properties DisplayName,EmailAddress,Title,Department,Enabled,LockedOut,DistinguishedName -ErrorAction Stop
    } -SuppressError
    
    # Get Groups
    $groups = Invoke-AD {
        Get-ADGroup -Filter * -Server $Global:CurrentDomain -Properties Description,GroupCategory,GroupScope,Members,DistinguishedName -ErrorAction Stop
    } -SuppressError
    
    # Get Computers
    $computers = Invoke-AD {
        Get-ADComputer -Filter * -Server $Global:CurrentDomain -Properties OperatingSystem,OperatingSystemVersion,Enabled,LastLogonDate,DistinguishedName -ErrorAction Stop
    } -SuppressError
    
    # Check if we got any data
    if ($null -eq $dcs -or $null -eq $users -or $null -eq $groups -or $null -eq $computers) {
        Debug-Log "One or more AD queries returned null - falling back to DEMO mode" -Type "warn"
        $Global:DemoMode = $true
        $converted = Convert-DataToADObjects `
            -Users $Global:rawUsers `
            -DCs $Global:rawDCs `
            -Groups $Global:rawDemoGroups `
            -Computers $Global:rawComputers `
            -Domain $Global:CurrentDomain
        return $converted
    }
    
    Debug-Log "Successfully retrieved AD data: $($dcs.Count) DCs, $($users.Count) Users, $($groups.Count) Groups, $($computers.Count) Computers" -Type "info"
    
    # Convert to your object format (you may need to create this function or adapt Convert-DataToADObjects)
    $converted = Convert-DataToADObjects -Users $users -DCs $dcs -Groups $groups -Computers $computers -Domain $Global:CurrentDomain
    return $converted
    
} catch {
    Debug-Log "Unexpected error loading production data: $_" -Type "error"
    Write-Warning "Failed to load production Active Directory data. Falling back to DEMO mode."
    $Global:DemoMode = $true
    
    # Load demo data
    $converted = Convert-DataToADObjects -Users $Global:rawUsers -DCs $Global:rawDCs -Groups $Global:rawDemoGroups -Computers $Global:rawComputers -Domain $Global:CurrentDomain
    return $converted
  }
  
}



## ----------------------{ Convert Domain Data }-----------------------
function Convert-DataToADObjects {
  param(
    [array]$Users,
    [array]$DCs = @(),
    [array]$Computers = @(),
    [array]$Groups = @(),
    [string]$Domain = "example.com",
    [string]$BaseDN = "DC=example,DC=com"
  )

  Write-Debug "DEBUG: Converting demo data to AD-like objects..."

  ## Helper functions
  function New-FakeGuid { [guid]::NewGuid().ToString() }
  function New-FakeSid { $rid = Get-Random -Minimum 1000 -Maximum 65535 ; "S-1-5-21-{0}-{1}-{2}-{3}" -f (Get-Random -Max 999999999), (Get-Random -Max 999999999), (Get-Random -Max 999999999), $rid }

  ## Convert Users
  $convertedUsers = @()
  foreach ($user in $Users) {
    $sam = ($user.Name -replace '\s+', '.').ToLower()
    $upn = if ($user.Email) { $user.Email } else { "$sam@$Domain" }

    # Derive domain from email BEFORE creating object
    $userDomain = if ($user.Email -and $user.Email -match '@(.+)$') {
        $matches[1]
    } else {
        'example.com'  # Default fallback
    }
    
    # Build DN BEFORE creating object
    if ($user.OU) {
      $ouChain = $user.OU | ForEach-Object { "OU=$_" }
      $dn = "CN=$($user.Name)," + ($ouChain[-1..0] -join ',') + ",$BaseDN"
    } else {
      $dn = "CN=$($user.Name),$BaseDN"
    }

    # NOW create the object with calculated values
    $adUser = [PSCustomObject]@{
      ObjectClass       = 'user'
      Name              = $user.Name
      Domain            = $userDomain  ## Automatically set from email
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

      ## Contact info
      Title       = $user.Title
      Department  = $user.Department
      Company     = $user.Company
      Manager     = $user.Manager
      EmailAddress= $user.Email
      OfficePhone = $user.Phone
      MobilePhone = $user.MobilePhone
      Office      = $user.Office

      ## Address
      StreetAddress = $user.Street
      City          = $user.City
      PostalCode    = $user.PostalCode
      Country       = $user.Country

      ## Other
      Description = $user.Description
      MemberOf    = $user.Groups
      CanonicalName = ($user.OU -join '/') + "/$($user.Name)"
      whenCreated   = (Get-Date).AddDays(-90)
      whenChanged   = (Get-Date).AddDays(-5)

      ## Original demo properties
      OU       = $user.OU
      Groups   = $user.Groups
      Disabled = $user.Disabled
      Locked   = $user.Locked
    }

    ## Make it look like AD user object
    $adUser.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADUser')
    $convertedUsers += $adUser
  }

  ## Convert DCs
  $convertedDCs = @()
  foreach ($dc in $DCs) {
    $dn = "CN=$($dc.Name),OU=Domain Controllers,$BaseDN"
    
    ## Derive domain from DC location or IP subnet
    $dcDomain = if ($dc.Location -match 'Germany') {
        'example.net'
    } else {
        'example.com'
    }
    
    $adDC = [PSCustomObject]@{
      ObjectClass               = 'computer'
      Name                      = $dc.Name
      Domain                    = $dcDomain
      DNSHostName               = "$($dc.Name).$dcDomain"
      DistinguishedName         = $dn
      ObjectGUID                = New-FakeGuid
      SID                       = New-FakeSid
      Enabled                   = $true
      Site                      = $dc.Site
      Location                  = $dc.Location
      OperatingSystem           = $dc.OS
      OperatingSystemVersion    = if ($dc.OS -match '2022') { '10.0 (20348)' } elseif ($dc.OS -match '2019') { '10.0 (17763)' } else { '10.0 (14393)' }
      IPv4Address               = $dc.IPAddress
      IsGlobalCatalog           = $dc.IsGlobalCatalog
      FSMORoles                 = $dc.FSMORoles
      LastReplication           = $dc.LastReplication
      ReplicationHealth         = $dc.ReplicationHealth
      LastBootUpTime            = $dc.LastBoot
      Services                  = $dc.Services
      DiskSpace                 = $dc.DiskSpace
      ReplicationPartners       = $dc.ReplicationPartners
      whenCreated               = (Get-Date).AddDays(-180)
      OU                        = 'Domain Controllers'
    }
    $adDC.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADComputer')
    $convertedDCs += $adDC
  }

  ## Convert Groups
  $convertedGroups = @()
  foreach ($group in $Groups) {
    $sam = ($group.Name -replace '\s+', '.').ToLower()
    $dn = "CN=$($group.Name),OU=Groups,$BaseDN"
    
    ## Derive domain from email
    $groupDomain = if ($group.Email -and $group.Email -match '@(.+)$') {
        $matches[1]
    } else {
        'example.com'
    }
      
    $adGroup = [PSCustomObject]@{
      ObjectClass       = 'group'
      Name              = $group.Name
      Domain            = $groupDomain
      SamAccountName    = $sam
      DistinguishedName = $dn
      ObjectGUID        = New-FakeGuid
      SID               = New-FakeSid
      GroupCategory     = $group.Type
      GroupScope        = $group.Scope
      Description       = $group.Description
      ManagedBy         = $group.ManagedBy
      Mail              = $group.Email
      whenCreated       = (Get-Date).AddDays(-180)
      whenChanged       = (Get-Date).AddDays(-10)
    }
      
    # Make it look like AD group object
    $adGroup.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADGroup')
    $convertedGroups += $adGroup
  }

  ## Convert Computers
  $convertedComputers = @()
  foreach ($computer in $Computers) {
    $dn = if ($computer.OU) {
        $ouChain = $computer.OU | ForEach-Object { "OU=$_" }
        "CN=$($computer.Name)," + ($ouChain[-1..0] -join ',') + ",$BaseDN"
    } else {
        "CN=$($computer.Name),CN=Computers,$BaseDN"
    }
    
    # Derive domain from location
    $compDomain = if ($computer.Location -match 'Germany') {
        'example.net'
    } else {
        'example.com'
    }
    
    $adComputer = [PSCustomObject]@{
        ObjectClass            = 'computer'
        Name                   = $computer.Name
        Domain                 = $compDomain
        DNSHostName            = "$($computer.Name).$compDomain"
        DistinguishedName      = $dn
        ObjectGUID             = New-FakeGuid
        SID                    = New-FakeSid
        Enabled                = $computer.Enabled
        OperatingSystem        = $computer.OS
        OperatingSystemVersion = $computer.OSVersion
        IPv4Address            = $computer.IPv4Address
        LastLogonDate          = $computer.LastLogon
        Location               = $computer.Location
        Description            = $computer.Description
        ComputerType           = $computer.Type  # Custom property
        whenCreated            = (Get-Date).AddDays(-120)
        OU                     = $computer.OU
    }
    
    $adComputer.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADComputer')
    $convertedComputers += $adComputer
  }

  ## Set global variables
  $Global:Users     = $convertedUsers
  $Global:Computers = $convertedComputers
  $Global:DCs       = $convertedDCs
  $Global:Groups    = $convertedGroups
  $Global:ADObjects = $convertedUsers + $convertedDCs + $convertedGroups + $convertedComputers

  Write-Debug "DEBUG: Converted $($convertedUsers.Count) users, $($convertedDCs.Count) DCs, Computers: $($convertedComputers.Count) and $($convertedGroups.Count) groups to AD-like objects"

  ## Return hashtable
  return @{
    Users = $convertedUsers
    DCs = $convertedDCs
    Groups = $convertedGroups
    Computers = $convertedComputers
  }
}

##-------------------{ Convwert Object to Tree Items }-------------------
function Convert-ToTreeNode {
  param([OUNode]$node)

  $tn = [Terminal.Gui.Trees.TreeNode]::new($node, $node.Name)

  foreach ($child in $node.Children) {
    $tn.Children.Add((Convert-ToTreeNode $child))
  }

  return $tn
}

##-------------------{ Build Domain Content }-------------------
function Build-DomainContent {
    param(
        [Terminal.Gui.Trees.TreeNode]$domainNode,
        [string]$domain
    )
    
    Debug-Log ("DEBUG: Building content for domain: $domain") -Type "info"
    
    ## Apply filters - filter users by domain
    $nameFilter = $Global:FilterOptions.NameFilter.Trim()
    
    # Filter users for this specific domain
    $domainUsers = $Global:Users | Where-Object { $_.Domain -eq $domain }
    
    $filteredUsers = $domainUsers | Where-Object { 
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

    Debug-Log ("DEBUG: Filtered to $($filteredUsers.Count) users for domain $domain") -Type "info"

    ## Cache for finding nodes by path
    $nodeCache = @{}

    ## Helper function to get or create node
    function Get-OrCreateChildNode {
        param(
            [Terminal.Gui.Trees.TreeNode]$Parent,
            [string]$Name,
            [string]$FullPath
        )

        if ($nodeCache.ContainsKey($FullPath)) { return $nodeCache[$FullPath] }

        $newNode = [Terminal.Gui.Trees.TreeNode]::new($Name)
        
        try {
            if ($null -eq $Parent.Children) {
                $Parent | Add-Member -Force -MemberType NoteProperty -Name Children -Value (New-Object 'System.Collections.ObjectModel.Collection[Terminal.Gui.Trees.ITreeNode]')
            }
            $Parent.Children.Add($newNode)
        } catch {
            Debug-Log ("WARNING: Could not add node $Name to parent: $_") -Type "warn"
        }

        $nodeCache[$FullPath] = $newNode
        return $newNode
    }

    ## Build OU hierarchy for this domain

    $userCount = 0
    foreach ($user in $filteredUsers) {
        $userCount++
        
        # Update every 25 users to avoid spam
        if ($userCount % 25 -eq 0) {
            Update-Status "Building $domain - $userCount/$($filteredUsers.Count) users" -spinner
        }

        if (-not $user.OU) { continue }

        $ouPath = @($user.OU)
        if ($ouPath.Count -eq 0) { continue }

        $currentNode = $domainNode
        $pathSoFar = ""

        foreach ($ouLevel in $ouPath) {
            $pathSoFar = if ($pathSoFar) { "$pathSoFar/$ouLevel" } else { $ouLevel }
            $currentNode = Get-OrCreateChildNode -Parent $currentNode -Name $ouLevel -FullPath "$domain/$pathSoFar"
        }

        $statusIcon = if ($user.Locked) { "🔒" } elseif ($user.Disabled) { "⊗" } else { "○" }
        $userNode = [Terminal.Gui.Trees.TreeNode]::new("(U) $statusIcon $($user.Name)")
        $userNode.Tag = $user
       
        if ($null -eq $currentNode.Children) {
            $currentNode | Add-Member -Force -MemberType NoteProperty -Name Children -Value (New-Object 'System.Collections.ObjectModel.Collection[Terminal.Gui.Trees.ITreeNode]')
        }
        $currentNode.Children.Add($userNode)
    }

    ## Add Groups for this domain
    if ($Global:FilterOptions.ShowGroups) {
        Update-Status "Building $domain - adding groups" -spinner
        $domainGroups = $Global:Groups | Where-Object { $_.Domain -eq $domain }
        
if ($domainGroups.Count -gt 0) {
        $groupsNode = Get-OrCreateChildNode -Parent $domainNode -Name "Groups" -FullPath "$domain/_Groups"
        
        foreach ($group in ($domainGroups | Sort-Object -Property Name)) {
            $groupNode = [Terminal.Gui.Trees.TreeNode]::new("(G) $($group.Name)")  # ADD PREFIX HERE
            $groupNode.Tag = $group  # ADD THIS to store group object
            $members = $filteredUsers | Where-Object { $_.Groups -contains $group.Name } | Sort-Object -Property Name

                foreach ($member in $members) {
                    $statusIcon = if ($member.Locked) { "🔒" } elseif ($member.Disabled) { "⊗" } else { "○" }
                    $memberNode = [Terminal.Gui.Trees.TreeNode]::new("(U) $statusIcon $($member.Name)")
                    $memberNode.Tag = $member
                    
                    if ($null -eq $groupNode.Children) {
                        $groupNode | Add-Member -Force -MemberType NoteProperty -Name Children -Value (New-Object 'System.Collections.ObjectModel.Collection[Terminal.Gui.Trees.ITreeNode]')
                    }
                    $groupNode.Children.Add($memberNode)
                }
                
                if ($groupNode.Children -and $groupNode.Children.Count -gt 0) {
                    $groupsNode.Children.Add($groupNode)
                }
            }
        }
    }

    ## Add DCs for this domain
    if ($Global:FilterOptions.ShowDCs) {
        Update-Status "Building $domain - adding DCs" -spinner
        $domainDCs = $Global:DCs | Where-Object { $_.Domain -eq $domain }
        
        if ($domainDCs.Count -gt 0) {
            $dcNode = Get-OrCreateChildNode -Parent $domainNode -Name "Domain Controllers" -FullPath "$domain/_DCs"
            
            foreach ($dc in ($domainDCs | Sort-Object -Property Name)) {
                $dcChildNode = [Terminal.Gui.Trees.TreeNode]::new("(DC) $($dc.Name) [$($dc.Site)]")
                $dcChildNode.Tag = $dc
                $dcNode.Children.Add($dcChildNode)
            }
        }
    }

         ## Add Computers
    if ($Global:FilterOptions.ShowComputers) {  # You'll need to add this filter option
                  Update-Status "Building $domain - adding computers" -spinner
    
      $domainComputers = $Global:Computers | Where-Object { $_.Domain -eq $domain }
    
     if ($domainComputers.Count -gt 0) {
      $computerNode = Get-OrCreateChildNode -Parent $domainNode -Name "Computers" -FullPath "$domain/_Computers"
        
        # Group by type
        $byType = $domainComputers | Group-Object ComputerType
        
        foreach ($typeGroup in $byType) {
            $typeNode = [Terminal.Gui.Trees.TreeNode]::new("$($typeGroup.Name)s")  # "Workstations", "Laptops", etc.
            
            foreach ($comp in ($typeGroup.Group | Sort-Object -Property Name)) {
                $icon = if (-not $comp.Enabled) { "⊗" } else { "💻" }
                $compNode = [Terminal.Gui.Trees.TreeNode]::new("(PC) $icon $($comp.Name)")
                $compNode.Tag = $comp
                
                if ($null -eq $typeNode.Children) {
                    $typeNode | Add-Member -Force -MemberType NoteProperty -Name Children -Value (New-Object 'System.Collections.ObjectModel.Collection[Terminal.Gui.Trees.ITreeNode]')
                }
                $typeNode.Children.Add($compNode)
            }
            
            if ($typeNode.Children -and $typeNode.Children.Count -gt 0) {
                $computerNode.Children.Add($typeNode)
            }
        }
    }
}
    Update-Status "$domain complete" -final    
    Debug-Log ("DEBUG: Finished building content for domain $domain - nodes in cache: $($nodeCache.Count)") -Type "success"
}

##---------------------------{ Build The Tree }--------------------------
function Build-Tree {
    param([string]$domain)

    # If no domain was provided, fall back to the global selected one
    if (-not $domain) {
        $domain = $Global:CurrentDomain
    }

    Debug-Log ("DEBUG: Building tree...") -Type "info"
    $Global:tree.ClearObjects()

    $root = $null

    ## ----------------------------------------------------
    ## Multi-domain forest
    ## ----------------------------------------------------
    if ($Global:Domains.Count -gt 1) {

        Debug-Log ("DEBUG: Creating multi-domain forest tree with root: $Global:ForestName") -Type "info"
        $root = [Terminal.Gui.Trees.TreeNode]::new($Global:ForestName)

        foreach ($dom in $Global:Domains) {

            Debug-Log ("DEBUG: Adding domain node: $($dom)") -Type "info"
            $domainNode = [Terminal.Gui.Trees.TreeNode]::new($dom)

            # Ensure children collection exists
            if ($null -eq $root.Children) {
                $root | Add-Member -Force -MemberType NoteProperty -Name Children -Value (
                    New-Object 'System.Collections.ObjectModel.Collection[Terminal.Gui.Trees.ITreeNode]'
                )
            }

            $root.Children.Add($domainNode)

            # Build domain subtree
            Build-DomainContent -domainNode $domainNode -domain $dom
        }

        ## ----------------------------------------------------
        ## FIX: Auto-select the correct domain for UI use
        ## ----------------------------------------------------
        if (-not $Global:CurrentDomain) {
            $Global:CurrentDomain = $Global:Domains[0]
            Debug-Log ("DEBUG: CurrentDomain was empty. Auto-selected: $Global:CurrentDomain") -Type "info"
        }

    }

    ## ----------------------------------------------------
    ## Single-domain mode
    ## ----------------------------------------------------
    else {
        Debug-Log ("DEBUG: Creating single-domain tree: $($Global:Domains[0])") -Type "info"
        $root = [Terminal.Gui.Trees.TreeNode]::new($Global:Domains[0])

        Build-DomainContent -domainNode $root -domain $Global:Domains[0]

        ## FIX: Set active domain if not set
        if (-not $Global:CurrentDomain) {
            $Global:CurrentDomain = $Global:Domains[0]
            Debug-Log ("DEBUG: CurrentDomain set (single-domain): $Global:CurrentDomain") -Type "info"
        }
    }

    ## ----------------------------------------------------
    ## Add root to TUI tree control
    ## ----------------------------------------------------
    $Global:tree.AddObject($root)
    Debug-Log ("DEBUG: Tree built successfully") -Type "success"

    ## ----------------------------------------------------
    ## FIX: Auto-select the correct tree node
    ## ----------------------------------------------------
    try {
        if ($Global:Domains.Count -gt 1) {
            # Multi-domain: select the domain node that matches CurrentDomain
            $selected = $root.Children | Where-Object { $_.Text -eq $Global:CurrentDomain } | Select-Object -First 1

            if ($selected) {
                $Global:tree.SelectedObject = $selected
                Debug-Log ("DEBUG: Tree auto-selected domain node: $($selected.Text)") -Type "info"
            }
            else {
                Debug-Log ("WARN: Could not auto-select domain: $($Global:CurrentDomain)") -Type "warn"
            }
        }
        else {
            # Single-domain: root is the selected node
            $Global:tree.SelectedObject = $root
            Debug-Log ("DEBUG: Tree auto-selected single-domain root: $($root.Text)") -Type "info"
        }
    }
    catch {
        Debug-Log ("ERROR: Failed to auto-select tree node: $_") -Type "error"
    }
}

## -----------------------{ Update Filter Label }----------------------
function Create-FilterStatusLabel {
    Debug-Log ("DEBUG: Entered Create-FilterStatusLabel") -Type "info"
    try {
        $lblStatus = [Terminal.Gui.Label]::new("")
        Debug-Log ("DEBUG: Created label object: $lblStatus") -Type "info"
        $lblStatus.X = 80
        $lblStatus.Y = 6
        $lblStatus.Width = 40
        return $lblStatus
    }
    catch {
        Debug-Log ("ERROR in Create-FilterStatusLabel: $($_.Exception.Message)") -Type "error"
        return $null
    }
}


## -----------------------{ Update Filter Label }----------------------
function Update-FilterStatusLabel {
  param($label)
    
  if (-not $label) {
    Debug-Log ("WARNING: label parameter is null in Update-FilterStatusLabel") -Type "warn"
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

## ------------------------{ Update Status Bar }-----------------------
function Update-Status {
    param(
        [string]$message,
        [switch]$spinner,
        [switch]$final
    )
    
    Debug-Log ("DEBUG: Update-Status called - message='$message', spinner=$spinner, final=$final") -Type "info"
    
    # Check if UI is running before trying to update it
    if (-not [Terminal.Gui.Application]::MainLoop) {
        Debug-Log ("DEBUG: MainLoop not running, skipping UI update") -Type "info"
        return
    }
    
    # Static prefix with forest and object counts
    $staticPrefix = "Forest: $($Global:ForestName) | Objects: $($Global:ADObjects.Count)"
    
    if ($final) {
        $displayText = "$staticPrefix | ✓ $message"
        Debug-Log ("DEBUG: Setting final status to: '$displayText'") -Type "success"
        
        [Terminal.Gui.Application]::MainLoop.Invoke({
            $Global:StatusItem.Title = $displayText
            $Global:statusBar.SetNeedsDisplay()
        })
        
    } elseif ($spinner) {
        [Terminal.Gui.Application]::MainLoop.Invoke({
            $Global:statusSpinnerIndex = ($Global:statusSpinnerIndex + 1) % 4
            $spinChar = $Global:statusSpinner[$Global:statusSpinnerIndex]
            $displayText = "$staticPrefix | $spinChar $message"
            $Global:StatusItem.Title = $displayText
            $Global:statusBar.SetNeedsDisplay()
        })
        Debug-Log ("STATUS: $message") -Type "info"
        
    } else {
        Debug-Log ("DEBUG: Setting static status to: '$message'") -Type "info"
        
        [Terminal.Gui.Application]::MainLoop.Invoke({
            $Global:StatusItem.Title = "$staticPrefix | $message"
            $Global:statusBar.SetNeedsDisplay()
        })
    }
}

# ------------------------- Filter Panel (Add to main window) ------------------------
function Create-FilterPanel {
  ## Create a frame for filters
  $filterFrame = [Terminal.Gui.FrameView]::new("Filters")
  $filterFrame.X = 32  # Right of the tree
  $filterFrame.Y = 1
  $filterFrame.Width = 40
  $filterFrame.Height = 12
  $y = 0
    
  ## Name filter
  $lblNameFilter = [Terminal.Gui.Label]::new("Name contains:"); $lblNameFilter.X=1; $lblNameFilter.Y=$y; $filterFrame.Add($lblNameFilter)
  $txtNameFilter = [Terminal.Gui.TextField]::new($Global:FilterOptions.NameFilter)
  $txtNameFilter.X=1; $txtNameFilter.Y=$y+1; $txtNameFilter.Width=35
  $txtNameFilter.add_TextChanged({ $Global:FilterOptions.NameFilter = $txtNameFilter.Text.ToString() })
  $filterFrame.Add($txtNameFilter)
  $y+=3
    
  ## Show/Hide checkboxes
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
    
  ## Sort options
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
    
## Apply/Reset buttons
$btnApplyFilter = [Terminal.Gui.Button]::new("Apply Filter")
$btnApplyFilter.X=1; $btnApplyFilter.Y=$y
$btnApplyFilter.add_Clicked({
  Debug-Log ("DEBUG: Applying filters...") -Type "info"
  Build-Tree -domain $Global:CurrentDomain
  })
  $filterFrame.Add($btnApplyFilter)
    
  $btnResetFilter = [Terminal.Gui.Button]::new("Reset")
  $btnResetFilter.X=17; $btnResetFilter.Y=$y
  $btnResetFilter.add_Clicked({
    Debug-Log ("DEBUG: Resetting filters...") -Type "info"
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
        
    ## Reset UI controls
    $chkEnabled.Checked = $true
    $chkDisabled.Checked = $true
    $chkGroups.Checked = $true
    $chkDCs.Checked = $true
    $txtNameFilter.Text = ""
    $rdoSort.SelectedItem = 0
       
    Build-Tree -domain $Global:CurrentDomain
  })

  $filterFrame.Add($btnResetFilter)
  return $filterFrame
}

## -------------------{ Quick Filter Menu (for Menu Bar) -------------------
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
    Debug-Log ("DEBUG: Applying quick filter: $selected") -Type "info"
        
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
        ## This would require additional logic to track last logon
        Show-Modal "Filter" "Filter applied: $selected"
        }
      "Users with No Manager" {
        ## Filter users with no manager
        $Global:FilterOptions.NameFilter = ""
        }
      "Empty Groups" {
        ## Show only groups with no members
        Show-Modal "Filter" "Filter applied: $selected"
        }
      "Domain Admins Only" {
        $Global:FilterOptions.NameFilter = ""
        }
      }
        
      Build-Tree -domain $Global:CurrentDomain
      [Terminal.Gui.Application]::RequestStop()
  })

  $dlg.AddButton($btnApply)
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $dlg.AddButton($btnCancel)
    
  [Terminal.Gui.Application]::Run($dlg)
}

## .----------------------------{ Get-ADHealth }----------------------------.
## | - Tabbed AD Health modal for DSA-TUI (Final: Tools row, Health icons,  |
## |   Alt-key accelerators, help line)                                     |
## | - Uses $Global:CurrentDomain for the domain to query                   |
## | - auto-refresh Domain if it changes                                    |
## | - No F-keys, no Ctrl+Q/Ctrl+D; uses Alt+... + Esc as requested         |
## '------------------------------------------------------------------------'
function Get-ADHealth {

  ## --------------------------------{ Layout }-------------------------------
  $width  = 118
  $height = 36

  ## Ensure cache / state exists
  if (-not $script:ADHealth) { $script:ADHealth = @{} }
  if (-not $script:ADHealth.Cache) { $script:ADHealth.Cache = @{} }

  ## Safe external invocation helper
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

  ## Initialize / Refresh caches and detect tools
  function Initialize-ADHealth {
    if (-not $Global:CurrentDomain -or [string]::IsNullOrWhiteSpace($Global:CurrentDomain)) {
      throw "Global variable `$Global:CurrentDomain is not set. Set `$Global:CurrentDomain before calling Get-ADHealth."
    }
    $domain = $Global:CurrentDomain

    ## If domain changed since last run, clear cached data
    if ($script:ADHealth.Cache.Domain -ne $domain) {
      $script:ADHealth.Cache = @{
        Domain = $domain
        Timestamp = (Get-Date).ToString("s")
      }
    }

  ## Prepare blank data containers for fresh run
  $script:ADHealth.Data = @{
    Domain = $domain
    Timestamp   = (Get-Date)
    DCStatus    = @{ Summary = @(); Details = "" ; Health = "UNKNOWN" }
    Replication = @{ Summary = @(); Details = "" ; Health = "UNKNOWN" }
    DNS         = @{ Summary = @(); Details = "" ; Health = "UNKNOWN" }
    SYSVOL      = @{ Summary = @(); Details = "" ; Health = "UNKNOWN" }
    FSMO        = @{ Summary = @(); Details = "" ; Health = "UNKNOWN" }
    GPO         = @{ Summary = @(); Details = "" ; Health = "UNKNOWN" }
        }

    ## Detect available external utilities (like CertUI checks)
    $script:ADHealth.Tools = @{
      Repadmin = (Get-Command repadmin.exe -ErrorAction SilentlyContinue) -ne $null
      DcDiag   = (Get-Command dcdiag.exe -ErrorAction SilentlyContinue) -ne $null
      DfsrDiag = (Get-Command dfsrdiag.exe -ErrorAction SilentlyContinue) -ne $null
      Nltest   = (Get-Command nltest.exe -ErrorAction SilentlyContinue) -ne $null
      Dnscmd   = (Get-Command dnscmd.exe -ErrorAction SilentlyContinue) -ne $null
    }
  }

  ## DC Status
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
      $ip   = $dc.IPAddress -join ","
      $site = $dc.Site
      $os   = $dc.OperatingSystem
      $isGC = $dc.IsGlobalCatalog

      ## Reachability
      try { $ping = Test-Connection -ComputerName $name -Count 1 -Quiet -ErrorAction SilentlyContinue } catch { $ping = $false }
      $reachable = if ($ping) { "OK" } else { "FAIL" }

      ## Uptime
      try {
        $wmi = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $name -ErrorAction Stop
        $lastBoot = $wmi.LastBootUpTime
        $uptimeDays = ((Get-Date) - [Management.ManagementDateTimeConverter]::ToDateTime($lastBoot)).Days
      } catch { $uptimeDays = "?" }

      ## Services
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

    ## Determine simple health: if any Reachable=FAIL -> WARN, if many fails -> FAIL
    $fails = ($summary | Where-Object { $_.Reachable -ne "OK" }).Count
    if ($fails -eq 0) { $health = "OK" } elseif ($fails -lt ($summary.Count / 2)) { $health = "WARN" } else { $health = "FAIL" }
      $script:ADHealth.Data.DCStatus.Summary = $summary
      $script:ADHealth.Data.DCStatus.Details = $sb.ToString()
      $script:ADHealth.Data.DCStatus.Health = $health
    }

  ## Replication
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
        ## Simple health: if the output contains "failed" or "error" mark WARN
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

  ## DNS
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

    ## Per-DC DNS service status
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

  ## Health deduction
  $fails = ($summary -match "NOT FOUND|Err|Error").Count
  if ($fails -eq 0) { $health = "OK" } elseif ($fails -lt 2) { $health = "WARN" } else { $health = "FAIL" }
    $script:ADHealth.Data.DNS.Summary = $summary
    $script:ADHealth.Data.DNS.Details = $sb.ToString()
    $script:ADHealth.Data.DNS.Health = $health
  }

  ## SYSVOL / DFSR
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


######################## cleaned to here ############################

  ## FSMO roles
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
                Show-Modal "Details" "No details available."
            }
        })
        $Parent.Add($btnDetails)
    }

    # ----------------------------
    # UI: Build modal & controls
    # ----------------------------
    $modal = [Terminal.Gui.Window]::new("Active Directory Health Check - $($Global:CurrentDomain)")
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
                    Show-Modal "Tool Found" "$nameCopy detected - will be used where applicable."
                } else {
                    Show-Modal "Tool Missing" "$nameCopy not found. Some detail modes will be disabled or fall back to PowerShell cmdlets."
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
        try { Run-AllChecks -Domain $Global:CurrentDomain } finally { [Terminal.Gui.Application]::RequestStop() }
        Render-ADHealthTabs; Render-ADHealthContent -Tab $script:ActiveHealthTab
    })

    $btnRefreshTab.add_Click({
        # Run only the check for the active tab
        switch ($script:ActiveHealthTab) {
            0 { Run-DCStatusCheck -Domain $Global:CurrentDomain }
            1 { Run-ReplicationCheck -Domain $Global:CurrentDomain }
            2 { Run-DNSCheck -Domain $Global:CurrentDomain }
            3 { Run-SYSVOLCheck -Domain $Global:CurrentDomain }
            4 { Run-FSMOCheck -Domain $Global:CurrentDomain }
            5 { Run-GPOCheck -Domain $Global:CurrentDomain }
        }
        Render-ADHealthTabs; Render-ADHealthContent -Tab $script:ActiveHealthTab
    })

    # Export (plain-text)
    $btnExport.add_Click({
        $ts = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $fname = "ADHealth_${($Global:CurrentDomain -replace '[^a-zA-Z0-9\.-]','_')}_$ts.txt"
        $full = Join-Path -Path (Get-Location) -ChildPath $fname
        $sb = New-Object System.Text.StringBuilder
        $sb.AppendLine("AD Health Report for $($Global:CurrentDomain) - $ts") | Out-Null
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
            Show-Modal "Export" "Report exported to`n$full"
        } catch {
            Show-Modal "Export Error" "Failed to write file: $($_.Exception.Message)"
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
                if ($d -and $d.Length -gt 0) { Show-DetailsModal -Title "Details" -Content $d } else { Show-Modal "Details" "No details available for this tab." }
                $args.Handled = $true
            }
        } catch { }
    })

    # ----------------------------
    # Initial run & render
    # ----------------------------
    Run-AllChecks -Domain $Global:CurrentDomain
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

    # --- Generate-RandomPassworild UI ---
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
            Show-Modal "Invalid Length" "Password length must be 1-127."
            return
        }

        $pool = @()
        if ($chkUpper.Checked) { $pool += $UpperCase }
        if ($chkLower.Checked) { $pool += $LowerCase }
        if ($chkNums.Checked)  { $pool += $Numbers }
        if ($chkSyms.Checked)  { $pool += $Symbols }

        if ($pool.Count -eq 0) {
            Show-Modal "No Character Types" "Select at least one character type."
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
        Show-Modal "Copied" "Password copied to clipboard."
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

# First, create the Apply-UserChanges function (put this before Show-UserPropertiesDialog):

function Apply-UserChanges {
    param($user, $fields)
    
    Debug-Log ("DEBUG: Applying changes for user: $($user.Name)") -Type "info"
    
    try {
        if ($Global:DemoMode) {
            # Update demo user object
            $user.Name = $fields.txtName.Text.ToString()
            $user.Description = $fields.txtDesc.Text.ToString()
            $user.Office = $fields.txtOffice.Text.ToString()
            $user.OfficePhone = $fields.txtPhone.Text.ToString()
            $user.MobilePhone = $fields.txtMobile.Text.ToString()
            $user.EmailAddress = $fields.txtEmail.Text.ToString()
            $user.Disabled = $fields.chkDisabled.Checked
            $user.Locked = $fields.chkLocked.Checked
            $user.StreetAddress = $fields.txtStreet.Text.ToString()
            $user.City = $fields.txtCity.Text.ToString()
            $user.PostalCode = $fields.txtPostal.Text.ToString()
            $user.Country = $fields.txtCountry.Text.ToString()
            $user.Title = $fields.txtTitle.Text.ToString()
            $user.Department = $fields.txtDept.Text.ToString()
            $user.Company = $fields.txtCompany.Text.ToString()
            $user.Manager = $fields.txtManager.Text.ToString()
            
            Debug-Log ("SUCCESS: User changes applied (demo mode)") -Type "info"
            Show-Modal "Success" "Changes applied successfully (demo mode)"
            
            # Rebuild tree to reflect changes
            [Terminal.Gui.Application]::MainLoop.Invoke({
                Build-Tree -domain $Global:CurrentDomain
                if ($filterStatusLabel) {
                    Update-FilterStatusLabel -label $Global:FilterStatusLabel
                }
            })
            
        } else {
            # Production mode - use Set-ADUser
            $setParams = @{
                Identity = $user.SamAccountName
                DisplayName = $fields.txtName.Text.ToString()
                Description = $fields.txtDesc.Text.ToString()
                Office = $fields.txtOffice.Text.ToString()
                OfficePhone = $fields.txtPhone.Text.ToString()
                MobilePhone = $fields.txtMobile.Text.ToString()
                EmailAddress = $fields.txtEmail.Text.ToString()
                StreetAddress = $fields.txtStreet.Text.ToString()
                City = $fields.txtCity.Text.ToString()
                PostalCode = $fields.txtPostal.Text.ToString()
                Country = $fields.txtCountry.Text.ToString()
                Title = $fields.txtTitle.Text.ToString()
                Department = $fields.txtDept.Text.ToString()
                Company = $fields.txtCompany.Text.ToString()
                Manager = $fields.txtManager.Text.ToString()
            }
            
            Set-ADUser @setParams -ErrorAction Stop
            
            # Handle account status separately
            if ($fields.chkDisabled.Checked -and $user.Enabled) {
                Disable-ADAccount -Identity $user.SamAccountName -ErrorAction Stop
            } elseif (-not $fields.chkDisabled.Checked -and -not $user.Enabled) {
                Enable-ADAccount -Identity $user.SamAccountName -ErrorAction Stop
            }
            
            if ($fields.chkLocked.Checked -eq $false -and $user.LockedOut) {
                Unlock-ADAccount -Identity $user.SamAccountName -ErrorAction Stop
            }
            
            Debug-Log ("SUCCESS: User changes applied to AD") -Type "info"
            Show-Modal "Success" "Changes applied successfully"
            
            # Reload AD data
Refresh-Data -domain $Global:CurrentDomain
            if ($filterStatusLabel) {
                Update-FilterStatusLabel -label $Global:FilterStatusLabel
            }
        }
        
        $script:changesMade = $false
        return $true
        
    } catch {
        Debug-Log ("ERROR: Failed to apply changes: $($_.Exception.Message)") -Type "error"
        Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)"
        return $false
    }
}

function Show-UserPropertiesDialog {
    param($user)

   # Safety checks
    if (-not $user) { 
        Debug-Log ("ERROR: User object is null") -Type "warn"
        return 
    }
    
    Debug-Log ("DEBUG: Show-UserPropertiesDialog starting for: $($user.Name)") -Type "info"

    #    # Safety checks
    #    if (-not $user) { Write-Error "User object is null"; return }
    #    if (-not $Global:CurrentDomain) { $Global:CurrentDomain = "" }
    #    if ($Global:DemoMode -and -not $Global:Users) { $Global:Users = @() }

    # ----- Create buttons FIRST -----
    $btnOK = [Terminal.Gui.Button]::new("OK")
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnApply = [Terminal.Gui.Button]::new("Apply")

    # ----- Create dialog WITH buttons -----
    $dlg = [Terminal.Gui.Dialog]::new("User Properties", 100, 40, $btnOK, $btnCancel, $btnApply)

    # ----- TabView -----
    $tabView = [Terminal.Gui.TabView]::new()
    $tabView.X=0; $tabView.Y=0; $tabView.Width=[Terminal.Gui.Dim]::Fill(); $tabView.Height = [Terminal.Gui.Dim]::Fill(1)  # leave 1 line at bottom

    # ==================== General Tab ====================
    $generalTab = [Terminal.Gui.TabView+Tab]::new()
    $generalTab.Text = "General"
    $generalView = [Terminal.Gui.View]::new()

    $y = 1
    # Display Name
    $lbl = [Terminal.Gui.Label]::new("Display Name:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtName = [Terminal.Gui.TextField]::new($user.Name ?? ""); $txtName.X=20; $txtName.Y=$y; $txtName.Width=70
    $txtName.add_TextChanged({ $script:changesMade = $true }); $generalView.Add($txtName)
    $y+=2

    # Description
    $lbl = [Terminal.Gui.Label]::new("Description:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtDesc = [Terminal.Gui.TextField]::new($user.Description ?? ""); $txtDesc.X=20; $txtDesc.Y=$y; $txtDesc.Width=70
    $txtDesc.add_TextChanged({ $script:changesMade = $true }); $generalView.Add($txtDesc)
    $y+=2

    # Office
    $lbl = [Terminal.Gui.Label]::new("Office:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtOffice = [Terminal.Gui.TextField]::new($user.Office ?? ""); $txtOffice.X=20; $txtOffice.Y=$y; $txtOffice.Width=70
    $txtOffice.add_TextChanged({ $script:changesMade = $true }); $generalView.Add($txtOffice)
    $y+=2

    # Telephone
    Write-Debug "DEBUG: User Phone value='$($selUser.OfficePhone)'"
    $lbl = [Terminal.Gui.Label]::new("Telephone:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtPhone = [Terminal.Gui.TextField]::new($user.OfficePhone ?? ""); $txtPhone.X=20; $txtPhone.Y=$y; $txtPhone.Width=70
    $txtPhone.add_TextChanged({ $script:changesMade = $true }); $generalView.Add($txtPhone)
    $y+=2

    # Mobile Phone
    $lbl = [Terminal.Gui.Label]::new("Mobile Phone:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtMobile = [Terminal.Gui.TextField]::new($user.MobilePhone ?? ""); $txtMobile.X=20; $txtMobile.Y=$y; $txtMobile.Width=70
    $txtMobile.add_TextChanged({ $script:changesMade = $true }); $generalView.Add($txtMobile)
    $y+=2

    # E-mail 
    Write-Debug "DEBUG: User Email value='$($selUser.EmailAddress)'"
    $lbl = [Terminal.Gui.Label]::new("E-mail:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtEmail = [Terminal.Gui.TextField]::new($user.EmailAddress ?? ""); $txtEmail.X=20; $txtEmail.Y=$y; $txtEmail.Width=70
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
        $confirm = [Terminal.Gui.MessageBox]::Query(70, 10, "Apply Password", "Apply the following password to user:`n`n$($user.Name)`n`nPassword:`n$newPwd`n", "Apply", "Cancel" )
        if ($confirm -ne 0) { Show-Modal "Cancelled" "Password reset cancelled."; return }
        Write-Debug "DEBUG: Password reset for $($user.Name) to: $newPwd"
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
$txtStreet = [Terminal.Gui.TextField]::new($User.StreetAddress); $txtStreet.X=20; $txtStreet.Y=$y; $txtStreet.Width=70
$txtStreet.add_TextChanged({ $script:changesMade = $true })
$addressView.Add($txtStreet)
$y+=2

$lbl = [Terminal.Gui.Label]::new("City:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
$txtCity = [Terminal.Gui.TextField]::new($User.City); $txtCity.X=20; $txtCity.Y=$y; $txtCity.Width=70
$txtCity.add_TextChanged({ $script:changesMade = $true })
$addressView.Add($txtCity)
$y+=2

$lbl = [Terminal.Gui.Label]::new("Postal Code:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
$txtPostal = [Terminal.Gui.TextField]::new($User.PostalCode); $txtPostal.X=20; $txtPostal.Y=$y; $txtPostal.Width=20
$txtPostal.add_TextChanged({ $script:changesMade = $true })
$addressView.Add($txtPostal)
$y+=2

$lbl = [Terminal.Gui.Label]::new("Country:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
$txtCountry = [Terminal.Gui.TextField]::new($User.Country); $txtCountry.X=20; $txtCountry.Y=$y; $txtCountry.Width=70
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
$txtTitle = [Terminal.Gui.TextField]::new($user.Title); $txtTitle.X=20; $txtTitle.Y=$y; $txtTitle.Width=70
$txtTitle.add_TextChanged({ $script:changesMade = $true })
$orgView.Add($txtTitle)
$y+=2

$lbl = [Terminal.Gui.Label]::new("Department:"); $lbl.X=2; $lbl.Y=$y; $orgView.Add($lbl)
$txtDept = [Terminal.Gui.TextField]::new($user.Department); $txtDept.X=20; $txtDept.Y=$y; $txtDept.Width=70
$txtDept.add_TextChanged({ $script:changesMade = $true })
$orgView.Add($txtDept)
$y+=2

$lbl = [Terminal.Gui.Label]::new("Company:"); $lbl.X=2; $lbl.Y=$y; $orgView.Add($lbl)
$txtCompany = [Terminal.Gui.TextField]::new($user.Company); $txtCompany.X=20; $txtCompany.Y=$y; $txtCompany.Width=70
$txtCompany.add_TextChanged({ $script:changesMade = $true })
$orgView.Add($txtCompany)
$y+=2

$lbl = [Terminal.Gui.Label]::new("Manager:"); $lbl.X=2; $lbl.Y=$y; $orgView.Add($lbl)
$txtManager = [Terminal.Gui.TextField]::new($user.Manager); $txtManager.X=20; $txtManager.Y=$y; $txtManager.Width=70
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
    
    Write-Debug "DEBUG: Add to group clicked for user: $($currentUser.Name)"
    
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
            Show-Modal "Error" "Failed to retrieve groups"
            return
        }
    }
    
    # Filter out current groups
    $currentGroups = if ($currentUser['Groups']) { $currentUser['Groups'] } else { @() }
    $groupsToAdd = $availableGroups | Where-Object { $currentGroups -notcontains $_ }
    
    Write-Debug "DEBUG: Available groups to add: $($groupsToAdd.Count)"
    
    if ($groupsToAdd.Count -eq 0) {
      Show-Modal "No Groups" "User is already in all groups!"
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
        Write-Debug "DEBUG: Add OK clicked, SelectedItem = $($lstAvailGroups.SelectedItem)"
        Write-Debug "DEBUG: Current user in handler: '$($currentUser.Name)'"
        
        if ($lstAvailGroups.SelectedItem -eq -1) {
            Show-Modal "No Selection" "Please select a group"
            return
        }
        
        if (-not $groupsToAdd -or $groupsToAdd.Count -eq 0) {
            Debug-Log ("ERROR: groupsToAdd is null or empty") -Type "warn"
            Show-Modal "Error" "No groups available"
            return
        }
        
        if ($lstAvailGroups.SelectedItem -ge $groupsToAdd.Count) {
            Debug-Log ("ERROR: SelectedItem ($($lstAvailGroups.SelectedItem)) >= groupsToAdd.Count ($($groupsToAdd.Count))") -Type "error"
            Show-Modal "Error" "Invalid selection index"
            return
        }
        
        $selectedGroup = $groupsToAdd[$lstAvailGroups.SelectedItem]
        Write-Debug "DEBUG: Selected group: $selectedGroup"
        
        try {
            if ($Global:DemoMode) {
                # Initialize Groups if needed
                if (-not $currentUser.ContainsKey('Groups') -or $null -eq $currentUser['Groups']) {
                    Write-Debug "DEBUG: Initializing Groups array"
                    $currentUser['Groups'] = @()
                }
                
                # Ensure it's an array
                if ($currentUser['Groups'] -isnot [array]) {
                    Write-Debug "DEBUG: Converting Groups to array"
                    $currentUser['Groups'] = @($currentUser['Groups'])
                }
                
                # Add the group
                Write-Debug "DEBUG: Current groups before add: $($currentUser['Groups'] -join ', ')"
                $currentUser['Groups'] += $selectedGroup
                $currentUser['Groups'] = $currentUser['Groups'] | Sort-Object -Unique
                Write-Debug "DEBUG: Current groups after add: $($currentUser['Groups'] -join ', ')"
      Write-Debug "DEBUG: User $($currentUser.Name) added to group $selectedGroup (demo mode)"
                
                # Close dialog first
                [Terminal.Gui.Application]::RequestStop()
                
                # Then update UI with error handling
                [Terminal.Gui.Application]::MainLoop.Invoke({
                    try {
                        Write-Debug "DEBUG: Updating parent groups list..."
                        $parentGroupsList.SetSource($currentUser['Groups'])
                        Write-Debug "DEBUG: Parent groups list updated"
                        
                        # Skip tree rebuild - will happen when properties dialog closes
                        # Build-Tree -domain $Global:CurrentDomain
                        
                        Write-Debug "DEBUG: Updating filter status label..."
                        if ($filterStatusLabel) {
                            Update-FilterStatusLabel -label $Global:FilterStatusLabel
                            Write-Debug "DEBUG: Filter status label updated"
                        } else {
                            Debug-Log ("WARNING: filterStatusLabel is null, skipping update") -Type "info"
                        }
                    } catch {
                        Debug-Log ("ERROR in MainLoop.Invoke: $($_.Exception.Message)") -Type "error"
                        Debug-Log ("ERROR: Stack trace: $($_.ScriptStackTrace)") -Type "error"
                    }
                })
                
                $script:changesMade = $true
                
                # Success message
                Debug-Log ("SUCCESS: Added $($currentUser.Name) to group $selectedGroup") -Type "success"

                } else {
                # Production mode
                Add-ADGroupMember -Identity $selectedGroup -Members $currentUser.Name -ErrorAction Stop
                
                Write-Debug "DEBUG: User $($currentUser.Name) added to group $selectedGroup in AD"
                
                # Reload from AD
                $currentUser['Groups'] = @(Get-ADPrincipalGroupMembership -Identity $currentUser.Name | Select-Object -ExpandProperty Name | Sort-Object)
                
                # Close dialog first
                [Terminal.Gui.Application]::RequestStop()
                
                # Then update UI with error handling
                [Terminal.Gui.Application]::MainLoop.Invoke({
                    try {
                        Write-Debug "DEBUG: Updating parent groups list..."
                        $parentGroupsList.SetSource($currentUser['Groups'])
                        Write-Debug "DEBUG: Parent groups list updated"
                        
                        Write-Debug "DEBUG: Rebuilding tree..."
                        Build-Tree -domain $Global:CurrentDomain
                        Write-Debug "DEBUG: Tree rebuilt"
                        
                        Write-Debug "DEBUG: Updating filter status label..."
                        if ($filterStatusLabel) {
                            Update-FilterStatusLabel -label $Global:FilterStatusLabel
                            Write-Debug "DEBUG: Filter status label updated"
                        } else {
                            Debug-Log ("WARNING: filterStatusLabel is null, skipping update") -Type "warn"
                        }
                    } catch {
                        Debug-Log ("ERROR in MainLoop.Invoke: $($_.Exception.Message)") -Type "error"
                        Debug-Log ("ERROR: Stack trace: $($_.ScriptStackTrace)") -Type "error"
                    }
                })
                
                $script:changesMade = $true
                
                Debug-Log ("SUCCESS: Added $($currentUser.Name) to group $selectedGroup") -Type "success"
            }
       } catch {
            Debug-Log ("ERROR: Failed to add to group: $($_.Exception.Message)") -Type "error"
            Debug-Log ("ERROR: Exception type: $($_.Exception.GetType().FullName)") -Type "error"
            Debug-Log ("ERROR: Stack trace: $($_.ScriptStackTrace)")  -Type "error"
            Show-Modal "Error" "Failed to add to group:`n$($_.Exception.Message)"
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
    
    Write-Debug "DEBUG: Remove from group clicked for user: $($currentUser.Name)"
    
    # Check if a group is selected
    if ($lstGroups.SelectedItem -eq -1) {
        Show-Modal "No Selection" "Please select a group to remove"
        return
    }
    
    # Verify user has groups
    if (-not $currentUser.ContainsKey('Groups') -or -not $currentUser['Groups'] -or $currentUser['Groups'].Count -eq 0) {
     Show-Modal "No Groups" "User is not a member of any groups"
        return
    }
    
    # Verify selected index is valid
    if ($lstGroups.SelectedItem -ge $currentUser['Groups'].Count) {
        Debug-Log ("ERROR: SelectedItem out of range") -Type "error"
       Show-Modal "Error" "Invalid group selection"
        return
    }
    
    # Get the selected group
    $selectedGroup = $currentUser['Groups'][$lstGroups.SelectedItem]
    Write-Debug "DEBUG: Selected group to remove: $selectedGroup"
    
    # Confirm removal - FIXED: Store message first
    $confirmMessage = "Remove $($currentUser.Name) from group '$selectedGroup'?"
    
    try {
        $result = [Terminal.Gui.MessageBox]::Query(60, 8, "Confirm Removal", $confirmMessage, @("Yes", "No"))
    } catch {
        Debug-Log ("ERROR: MessageBox failed: $($_.Exception.Message)") -Type "error"
        # Fall back to no confirmation, just do it
        $result = 0
    }
    
    if ($result -eq 0) {
        Write-Debug "DEBUG: User confirmed removal"
        
        try {
            if ($Global:DemoMode) {
                # Demo mode: remove from array
                Write-Debug "DEBUG: Current groups before remove: $($currentUser['Groups'] -join ', ')"
                $currentUser['Groups'] = $currentUser['Groups'] | Where-Object { $_ -ne $selectedGroup }
                Write-Debug "DEBUG: Current groups after remove: $($currentUser['Groups'] -join ', ')"
                
                Write-Debug "DEBUG: User $($currentUser.Name) removed from group $selectedGroup (demo mode)"
                
                # Update the groups list display
                # Update the groups list display
                [Terminal.Gui.Application]::MainLoop.Invoke({
                    try {
                        Write-Debug "DEBUG: Updating groups list..."
                        $lstGroups.SetSource($currentUser['Groups'])
                        $lstGroups.SetNeedsDisplay()
                        $lstGroups.Redraw($lstGroups.Bounds)
                        Write-Debug "DEBUG: Groups list updated"
                    } catch {
                        Debug-Log ("ERROR in MainLoop.Invoke: $($_.Exception.Message)") -Type "error"
                    }
                })                
                $script:changesMade = $true
                
                Debug-Log ("SUCCESS: Removed $($currentUser.Name) from group $selectedGroup") -Type "error"
                
            } else {
                # Production mode: remove from AD
                Remove-ADGroupMember -Identity $selectedGroup -Members $currentUser.Name -Confirm:$false -ErrorAction Stop
                
                Write-Debug "DEBUG: User $($currentUser.Name) removed from group $selectedGroup in AD"
                
                # Reload group membership from AD
                $currentUser['Groups'] = @(Get-ADPrincipalGroupMembership -Identity $currentUser.Name | Select-Object -ExpandProperty Name | Sort-Object)
                
                # Update the groups list display
                [Terminal.Gui.Application]::MainLoop.Invoke({
                    try {
                        Write-Debug "DEBUG: Updating groups list..."
                        $lstGroups.SetSource($currentUser['Groups'])
                        $lstGroups.SetNeedsDisplay()
                        $lstGroups.Redraw($lstGroups.Bounds)
                        Write-Debug "DEBUG: Groups list updated"
                    } catch {
                        Debug-Log ("ERROR in MainLoop.Invoke: $($_.Exception.Message)") -Type "error"
                    }
                })
                
                $script:changesMade = $true
                
                Debug-Log ("SUCCESS: Removed $($currentUser.Name) from group $selectedGroup") -Type "success"
            }
            
        } catch {
            Debug-Log ("ERROR: Failed to remove from group: $($_.Exception.Message)") -Type "error"
            Debug-Log ("ERROR: Stack trace: $($_.ScriptStackTrace)") -Type "error"
             Show-Modal "Error" "Failed to remove from group:`n$($_.Exception.Message)"
        }
    } else {
        Write-Debug "DEBUG: User cancelled removal"
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
    $txtSearchDomain = [Terminal.Gui.TextField]::new($Global:CurrentDomain ?? ""); $txtSearchDomain.X=15; $txtSearchDomain.Y=$y; $txtSearchDomain.Width=30; $searchView.Add($txtSearchDomain)
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

    $txtSearchFilter.add_TextChanged({
    if ($script:currentSearchOutputLines) {
      $search = $txtSearchFilter.Text.ToString().Trim()
      if ($search) {
        $txtSearchOutput.Text = ($script:currentSearchOutputLines | Where-Object { $_ -match "(?i)$search" }) -join "`n"
     } else {
              $txtSearchOutput.Text = $script:currentSearchOutputLines -join "`n"
          }
      }
  })


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
            $chkSearchDisabled.Checked = [bool]($user.Disabled)
            $chkSearchDisabled.Data    = $user.Name
        }
    })

#####################################################################################################
## placeholder menu items. Place tabs chronologically as to see them in order, 1.16
## doesn't support any other type of ordering

# ----- Profile -----
$profileTab = [Terminal.Gui.TabView+Tab]::new()
$profileTab.Text = "Proifle"
$profileView = [Terminal.Gui.View]::new()

$y = 1
$lbl = [Terminal.Gui.Label]::new("Profile Dir:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
$txtStreet = [Terminal.Gui.TextField]::new('Under Construction'); $txtStreet.X=20; $txtStreet.Y=$y; $txtStreet.Width=70
$txtStreet.add_TextChanged({ $script:changesMade = $true })
$addressView.Add($txtStreet)
$y+=2

$profileTab.View = $profileView
$tabView.AddTab($profileTab, $false)

# ----- Remote Desktop Services -----
$RDPTab = [Terminal.Gui.TabView+Tab]::new()
$RDPTab.Text = "Remote Desktop Services"
$RDPView = [Terminal.Gui.View]::new()

$y = 1
$lbl = [Terminal.Gui.Label]::new("RDP Profile Dir:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
$txtStreet = [Terminal.Gui.TextField]::new('Under Construction'); $txtStreet.X=20; $txtStreet.Y=$y; $txtStreet.Width=70
$txtStreet.add_TextChanged({ $script:changesMade = $true })
$addressView.Add($txtStreet)
$y+=2

$RDPTab.View = $RDPView
$tabView.AddTab($RDPTab, $false)

# ----- Sessions -----
$SessionsTab = [Terminal.Gui.TabView+Tab]::new()
$SessionsTab.Text = "Profile"
$SessionsView = [Terminal.Gui.View]::new()

$y = 1
$lbl = [Terminal.Gui.Label]::new("Sessions:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
$txtStreet = [Terminal.Gui.TextField]::new('Under Construction'); $txtStreet.X=20; $txtStreet.Y=$y; $txtStreet.Width=70
$txtStreet.add_TextChanged({ $script:changesMade = $true })
$addressView.Add($txtStreet)
$y+=2

$SessionsTab.View = $SessionsView
$tabView.AddTab($SessionsTab, $false)


#####################################################################################################

    # ----- Add TabView to Dialog -----
    $dlg.Add($tabView)

# ----- Wire up button actions (DON'T add them again!) -----
$btnOK.add_Clicked({
    # Apply changes and close
    if (Apply-UserChanges -user $user -fields $fields) {
        [Terminal.Gui.Application]::RequestStop()
    }
}.GetNewClosure())

$btnCancel.add_Clicked({
    # Check if changes were made
    if ($script:changesMade) {
        $result = [Terminal.Gui.MessageBox]::Query(60, 8, "Unsaved Changes", "You have unsaved changes. Discard them?",  @("Yes", "No"))
        if ($result -eq 0) {
            $script:changesMade = $false
            [Terminal.Gui.Application]::RequestStop()
        }
    } else {
        [Terminal.Gui.Application]::RequestStop()
    }
}.GetNewClosure())

$btnApply.add_Clicked({
    # Apply changes but keep dialog open
    Apply-UserChanges -user $user -fields $fields
}.GetNewClosure())

# Run the dialog (ONLY ONCE!)
Debug-Log ("DEBUG: Show-UserPropertiesDialog running") -Type "info"
[Terminal.Gui.Application]::Run($dlg)
Debug-Log ("DEBUG: Show-UserPropertiesDialog completed") -Type "info"
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
           Show-Modal "Error" "Name is required!"
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
            
            # Add to RAW user data
            $newUser = @{
                Name = $name
                OU = @($ou)  # OU as array
                Groups = @()
                Title = ""
                Email = $email
                Country = ""
                Disabled = $false
                Locked = $false
                MustChangePassword = $true
                Department = ""
                Office = ""
                Phone = ""
                MobilePhone = ""
                Street = ""
                City = ""
                PostalCode = ""
                Company = ""
                Manager = ""
                Description = $displayName
            }
            
            # Add to raw users array
            $Global:rawUsers += $newUser
            
            # Reconvert to update $Global:Users with AD-like objects
            $converted = Convert-DataToADObjects -Users $Global:rawUsers -DCs $Global:rawDCs -Groups $Global:rawDemoGroups -Domain $Global:CurrentDomain
            $Global:Users = $converted.Users
            
            Debug-Log ("DEBUG: Created user $name in demo mode") -Type "info"
        }
        
        "Group" {
            # Add to RAW group data
            $newGroup = @{
                Name = $name
                Description = $displayName
                Type = 'Security'
                Scope = 'Global'
                ManagedBy = ''
                Email = ''
            }
            
            # Add to raw groups array
            $Global:rawDemoGroups += $newGroup
            
            # Reconvert to update $Global:Groups with AD-like objects
            $converted = Convert-DataToADObjects -Users $Global:rawUsers -DCs $Global:rawDCs -Groups $Global:rawDemoGroups -Domain $Global:CurrentDomain
            $Global:Groups = $converted.Groups
            
            Debug-Log ("DEBUG: Created group $name in demo mode") -Type "info"
        }
        
        "OrganizationalUnit" {
            # For OUs, we need to track them in a structure
            # OUs are built from the OU arrays in users, so we could either:
            # 1. Add a $Global:rawOUs array (cleaner)
            # 2. Just ensure the OU path exists when we rebuild the tree
            
            # For now, let's just ensure it's tracked
            if (-not $Global:rawOUs) {
                $Global:rawOUs = @()
            }
            
            $newOU = @{
                Name = $name
                Path = $ou
                Description = $displayName
            }
            
            $Global:rawOUs += $newOU
            
            Debug-Log ("DEBUG: Created OU $name in demo mode") -Type "info"
        }
        
        "Computer" {
            # Similar to users, add to a computers array
            if (-not $Global:rawComputers) {
                $Global:rawComputers = @()
            }
            
            $newComputer = @{
                Name = $name
                OU = $ou
                Description = $displayName
            }
            
            $Global:rawComputers += $newComputer
            
            Debug-Log ("DEBUG: Created computer $name in demo mode") -Type "info"
        }
    }
    
    Show-Modal "Success" "$objType '$name' created successfully (demo mode)"
    Build-Tree -domain $Global:CurrentDomain
    Update-FilterStatusLabel -label $Global:FilterStatusLabel
    [Terminal.Gui.Application]::RequestStop()
    
            } else {
                # Production mode - create in AD
                switch ($objType) {
                    "User" {
                        $sam = $txtSam.Text.ToString().Trim()
                        $email = $txtEmail.Text.ToString().Trim()
                        $pwd = $txtPassword.Text.ToString()
                        
                        if (-not $sam) {
                            Show-Modal "Error" "Username (SAM) is required for users!"
                            return
                        }
                        
                        if (-not $pwd) {
                            Show-Modal "Error" "Password is required for users!"
                            return
                        }
                        
                        $secPwd = ConvertTo-SecureString -String $pwd -AsPlainText -Force
                        
                        $params = @{
                            Name = $name
                            SamAccountName = $sam
                            UserPrincipalName = "$sam@$($Global:CurrentDomain)"
                            AccountPassword = $secPwd
                            Enabled = $true
                            Path = $ou
                            ChangePasswordAtLogon = $true
                        }
                        
                        if ($displayName) { $params['DisplayName'] = $displayName }
                        if ($email) { $params['EmailAddress'] = $email }
                        
                        New-ADUser @params -ErrorAction Stop
                        Debug-Log ("DEBUG: Created user $name in AD") -Type "info"
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
                        Debug-Log ("DEBUG: Created group $name in AD") -Type "info"
                    }
                    "OrganizationalUnit" {
                        $params = @{
                            Name = $name
                            Path = $ou
                        }
                        
                        if ($displayName) { $params['Description'] = $displayName }
                        
                        New-ADOrganizationalUnit @params -ErrorAction Stop
                        Debug-Log ("DEBUG: Created OU $name in AD") -Type "info"
                    }
                    "Computer" {
                        $params = @{
                            Name = $name
                            Path = $ou
                        }
                        
                        New-ADComputer @params -ErrorAction Stop
                        Debug-Log ("DEBUG: Created computer $name in AD") -Type "info"
                    }
                    "Contact" {
                        $params = @{
                            Name = $name
                            Type = "Contact"
                            Path = $ou
                        }
                        
                        if ($displayName) { $params['DisplayName'] = $displayName }
                        
                        New-ADObject @params -ErrorAction Stop
                        Debug-Log ("DEBUG: Created contact $name in AD") -Type "info"
                    }
                }
                
                Show-Modal "Success" "$objType '$name' created successfully"
                
                # Refresh data
                Refresh-Data -domain $Global:CurrentDomain
                Update-FilterStatusLabel -label $Global:FilterStatusLabel
                [Terminal.Gui.Application]::RequestStop()
            }
            
        } catch {
            $errMsg = $_.Exception.Message
            Show-Modal "Error" "Failed to create $objType`:`n$errMsg"
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
    
           # Remove prefixes and ALL leading non-alphanumeric characters
# Remove prefixes like "(U) " or "(DC) "

    # Parse the tree text to get clean name and type
    $objInfo = Get-CleanObjectInfo -treeText $objectName
    $cleanName = $objInfo.Name

Debug-Log ("DEBUG: After removing prefix: '$cleanName'") -Type "info"

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
                        Debug-Log ("DEBUG: Deleted user $cleanName (demo mode)") -Type "info"
                    }
                    "group" {
                        # Remove group from all users
                        foreach ($u in $Global:Users) {
                            $u.Groups = $u.Groups | Where-Object { $_ -ne $cleanName }
                        }
                        Debug-Log ("DEBUG: Deleted group $cleanName (demo mode)") -Type "info"
                    }
                    default {
                        Debug-Log ("DEBUG: Deleted $objectType $cleanName (demo mode)") -Type "info"
                    }
                }
                
                Show-Modal "Deleted" "$objectType '$cleanName' deleted (demo mode)"
                Build-Tree -domain $Global:CurrentDomain
                Update-FilterStatusLabel -label $Global:FilterStatusLabel
                
            } else {
                # Production mode - delete from AD
                switch ($objectType.ToLower()) {
                    "user" {
                        Remove-ADUser -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log ("DEBUG: Deleted user $cleanName from AD") -Type "info"
                    }
                    "group" {
                        Remove-ADGroup -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log ("DEBUG: Deleted group $cleanName from AD") -Type "info"
                    }
                    "ou" {
                        Remove-ADOrganizationalUnit -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log ("DEBUG: Deleted OU $cleanName from AD") -Type "info"
                    }
                    "computer" {
                        Remove-ADComputer -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log ("DEBUG: Deleted computer $cleanName from AD") -Type "info"
                    }
                    default {
                        Remove-ADObject -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log ("DEBUG: Deleted $objectType $cleanName from AD") -Type "info"
                    }
                }
                
                Show-Modal "Deleted" "$objectType '$cleanName' deleted successfully"
                
                # Refresh data
                Refresh-Data -domain $Global:CurrentDomain
                Update-FilterStatusLabel -label $Global:FilterStatusLabel
            }
            
        } catch {
            $errMsg = $_.Exception.Message
           Show-Modal "Delete Failed" "Failed to delete $objectType`:`n$errMsg"
        }
    }
}

# ------------------------- Move Object ------------------------
function Show-MoveObjectDialog {
    param([string]$objectName, [string]$objectType)
    
          # Parse the tree text to get clean name and type
          $objInfo = Get-CleanObjectInfo -treeText $objectName
          $cleanName = $objInfo.Name
    
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
            Show-Modal "Error" "Please select a target OU"
            return
        }
        
        $targetOU = $ouList[$lstOU.SelectedItem]
        
        if ($targetOU -eq $currentOU) {
          Show-Modal "Error" "Object is already in that OU"
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
                            Debug-Log ("DEBUG: Moved user $cleanName to $targetOU (demo mode)") -Type "info"
                        }
                    }
                    
                    Show-Modal "Success" "Object moved successfully (demo mode)"
                    Build-Tree -domain $Global:CurrentDomain
                    Update-FilterStatusLabel -label $Global:FilterStatusLabel
                    [Terminal.Gui.Application]::RequestStop()
                    
                } else {
                    # Production mode - move in AD
                    $adObject = Get-ADObject -Filter "Name -eq '$cleanName'" -ErrorAction Stop
                    Move-ADObject -Identity $adObject.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
                    
                    Debug-Log ("DEBUG: Moved $cleanName to $targetOU in AD") -Type "info"
                   Show-Modal "Success" "Object moved successfully"
                    
                    # Refresh data                    
                    Refresh-Data -domain $Global:CurrentDomain
                    Update-FilterStatusLabel -label $Global:FilterStatusLabel
                    [Terminal.Gui.Application]::RequestStop()
                }
                
            } catch {
                $errMsg = $_.Exception.Message
                Show-Modal "Move Failed" "Failed to move object:`n$errMsg"
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
    $txtDomain = [Terminal.Gui.TextField]::new($Global:CurrentDomain); $txtDomain.X=15; $txtDomain.Y=0
    $dlg.Add($txtDomain)
    $okBtn = [Terminal.Gui.Button]::new("OK"); $okBtn.X=10; $okBtn.Y=2
    $okBtn.add_Clicked({
        $domainString = -join ($txtDomain.Text | ForEach-Object { [char]$_ })
        Debug-Log ("DEBUG: OK pressed, Domain = $domainString") -Type "info"
        $Global:CurrentDomain = $domainString
        Refresh-Data -domain $Global:CurrentDomain
        # Add after Build-Tree calls:
        Update-FilterStatusLabel -label $Global:FilterStatusLabel
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
## Check tree exists first or it blows things up If we're initalising, it is created later
if ($null -ne $Global:tree) {
  ## backup will remove for prod code
  ##  $tree.add_KeyPress({ param($sender,$keyArgs) if ($keyArgs.KeyEvent.Key -eq [Terminal.Gui.Key]::Enter -and $tree.SelectedObject) { $tree.SelectedObject.Expanded = -not $tree.SelectedObject.Expanded; $tree.SetNeedsDisplay(); $keyArgs.Handled = $true } })
  $Global:tree.Add_KeyPress({
    param($sender, $keyArgs)

    if ($keyArgs.KeyEvent.Key -ne [Terminal.Gui.Key]::Enter) { return }

    $node = $Global:tree.SelectedNode
    if (-not $node) {
        Debug-Log ("Enter pressed but no selected node — ignoring safely.") -Type "info"
        return
    }

    # Toggle expand/collapse properly
    if ($node.IsExpanded) {
        $Global:tree.CollapseNode($node)
    } else {
        $Global:tree.ExpandNode($node)
    }

    $Global:tree.SetNeedsDisplay()
    $keyArgs.Handled = $true
  })
} else {
    Debug-Log "Attempted to attach KeyPress handler but $Global:tree is null" -Type "warn"
}


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
    $dlg = [Terminal.Gui.Dialog]::new("Advanced Search - Active Directory",90,33)
    $dlg.X = 2; $dlg.Y = 02

    # Create TabView for Basic vs Advanced search
    $tabView = [Terminal.Gui.TabView]::new()
    $tabView.X = 0
    $tabView.Y = 0
    $tabView.Width = [Terminal.Gui.Dim]::Fill()
    ## “Take ALL remaining vertical space, minus Fill(n) rows. In other words the larger this number, the smaller the modal”
    ##$tabView.Height = [Terminal.Gui.Dim]::Fill(12)
    $tabView.Height = 16

    # Store search results globally so export can access them
    $script:lastSearchResults = @()
    $script:lastSearchType = ""

    # ----- Basic Search Tab -----
    $basicTab = [Terminal.Gui.TabView+Tab]::new()
    $basicTab.Text = "Basic Search"
    $basicView = [Terminal.Gui.View]::new()

    ## --- Domain Label ---
    $y = 1 
    $lblDomain = [Terminal.Gui.Label]::new("Domain:")
    $lblDomain.X = 2
    $lblDomain.Y = $y
    $basicView.Add($lblDomain)

    ## Ensure domains is always an array
    $domainList = @($Global:Domains)

    ## --- Domain ComboBox ---
    $comboDomain = [Terminal.Gui.ComboBox]::new()
    $comboDomain.X = 10
    $comboDomain.Y = $y
    $comboDomain.Width = 35
    $comboDomain.Height = 5  # enough room for dropdown
    $comboDomain.ReadOnly = $true   # user must pick from list
    $comboDomain.SetSource($domainList)

    ## Preselect current domain if available
    $initialIndex = $domainList.IndexOf($Global:CurrentDomain)
    if ($initialIndex -ge 0) {
      $comboDomain.SelectedItem = $initialIndex
    }

    $basicView.Add($comboDomain)

    ## --- Hidden text field to store selection (if you still need it) ---
    $txtDomain = [Terminal.Gui.TextField]::new($Global:CurrentDomain)
    $txtDomain.Visible = $false  # not shown, just stores the value
    $basicView.Add($txtDomain)

    ## Update txtDomain whenever domain changes
    $comboDomain.add_SelectedItemChanged({
      param($args)
      $selected = $comboDomain.Text
      $txtDomain.Text = $selected
    })
    $y+=2

    $lblName = [Terminal.Gui.Label]::new("Name:"); $lblName.X=2; $lblName.Y=$y; $basicView.Add($lblName)
    $txtUser = [Terminal.Gui.TextField]::new(""); $txtUser.X=10; $txtUser.Y=$y; $txtUser.Width=35; $basicView.Add($txtUser)
    $y+=2

    $lblType = [Terminal.Gui.Label]::new("Type:"); $lblType.X=2; $lblType.Y=$y; $basicView.Add($lblType)
    $cmbObjType = [Terminal.Gui.ComboBox]::new(); $cmbObjType.X=10; $cmbObjType.Y=$y; $cmbObjType.Width=20
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
            Show-Modal "Loaded" "Loaded filter: $($selected.Name)"
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
                Debug-Log ("DEBUG: Saved search '$newName'") -Type "info"
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
    ## results dialog you can chneg sizes here
    $txtOutput.X=2; $txtOutput.Y=[Terminal.Gui.Pos]::Bottom($lblResults); $txtOutput.Width=[Terminal.Gui.Dim]::Fill(2); $txtOutput.Height=10; $txtOutput.ReadOnly=$true
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
          Show-Modal "No Results" "No search results to export. Run a search first."
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
                Debug-Log ("DEBUG: Exported $($script:lastSearchResults.Count) results to $filename") -Type "info"
                Show-Modal "Success" "Exported $($script:lastSearchResults.Count) results to:`n$filename"
                [Terminal.Gui.Application]::RequestStop()
            } catch {
                $errMsg = $_.Exception.Message
               Show-Modal "Export Failed" "Failed to export:`n$errMsg"
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
    Debug-Log ("DEBUG: Refreshing tree data...") -Type "info"
    
    # Show loading dialog
    $loadingDlg = Show-LoadingDialog -Message "Refreshing Active Directory data..."
    
    try {
        # Reload domain data
        Load-DomainData -domain $Global:CurrentDomain
        
        # Rebuild tree
        Build-Tree -domain $Global:CurrentDomain
        # Add after Build-Tree calls:
        Update-FilterStatusLabel -label $Global:FilterStatusLabel
        
        Debug-Log ("DEBUG: Tree refreshed successfully") -Type "info"
    } finally {
        Close-LoadingDialog $loadingDlg
    }
    
    Show-Modal "Refreshed" "Active Directory data refreshed successfully"
}

function Show-DCPropertiesDialog {
    param([string]$dcName)
    
    Debug-Log ("DEBUG: Showing DC properties for: $dcName") -Type "info"
    
    if ($Global:DemoMode) {
        $dc = $Global:DCs | Where-Object { $_.Name -eq $dcName } | Select-Object -First 1
        
        if ($dc) {
            Debug-Log ("DEBUG: DC object found: $($dc.Name)") -Type "info"
            
            # Safely get property values with fallbacks
            $fsmoRoles = if ($dc.FSMORoles -and $dc.FSMORoles.Count -gt 0) { 
                $dc.FSMORoles -join ', ' 
            } else { 
                'None' 
            }
            
            $partners = if ($dc.ReplicationPartners -and $dc.ReplicationPartners.Count -gt 0) {
                $dc.ReplicationPartners -join ', '
            } else {
                'None'
            }
            
            $lastRep = if ($dc.LastReplication) { $dc.LastReplication.ToString('g') } else { 'Unknown' }
            $lastBoot = if ($dc.LastBootUpTime) { $dc.LastBootUpTime.ToString('g') } else { 'Unknown' }
            
            # Disk space - show both C: and SYSVOL
            $diskC = if ($dc.DiskSpace -and $dc.DiskSpace.'C:') { 
                "$($dc.DiskSpace.'C:'.Free) free of $($dc.DiskSpace.'C:'.Total) ($($dc.DiskSpace.'C:'.PercentFree)%)" 
            } else { 
                'Unknown' 
            }
            
            $diskSysvol = if ($dc.DiskSpace -and $dc.DiskSpace.'SYSVOL') { 
                "$($dc.DiskSpace.'SYSVOL'.Free) free of $($dc.DiskSpace.'SYSVOL'.Total) ($($dc.DiskSpace.'SYSVOL'.PercentFree)%)" 
            } else { 
                'Unknown' 
            }
            
            $msg = "Name: $($dc.Name)`nSite: $($dc.Site)`nLocation: $($dc.Location)`nIP Address: $($dc.IPv4Address)`nOperating System: $($dc.OperatingSystem)`nGlobal Catalog: $($dc.IsGlobalCatalog)`nFSMO Roles: $fsmoRoles`nReplication Health: $($dc.ReplicationHealth)`nLast Replication: $lastRep`nLast Boot: $lastBoot`nC: Drive: $diskC`nSYSVOL: $diskSysvol`nReplication Partners: $partners`n`n(Demo Mode)"
            
            if ($Global:VerboseMode) {
                Debug-Log ("VERBOSE: MessageBox content:") -Type "info"
                Debug-Log ("VERBOSE: Title: DC Properties") -Type "info"
                Debug-Log ("VERBOSE: Message:`n$msg") -Type "info"
            }
            
            Show-Modal "DC Properties" $msg
        } else {
            Show-Modal "Not Found" "DC '$dcName' not found"
        }
    } else {
        # Production mode
        try {
            $dc = Get-ADDomainController -Identity $dcName -ErrorAction Stop
            
            $msg = "Name: $($dc.Name)`nHostname: $($dc.HostName)`nSite: $($dc.Site)`nDomain: $($dc.Domain)`nIPv4: $($dc.IPv4Address)`nOS: $($dc.OperatingSystem)`nGlobal Catalog: $($dc.IsGlobalCatalog)"
            
            if ($Global:VerboseMode) {
                Debug-Log ("VERBOSE: MessageBox content:") -Type "info"
                Debug-Log ("VERBOSE: Title: DC Properties") -Type "info"
                Debug-Log ("VERBOSE: Message:`n$msg") -Type "info"
            }
            
            Show-Modal "DC Properties" $msg
        } catch {
            Show-Modal "Error" "Failed to get DC properties:`n$_"
        }
    }
}

function Show-GroupPropertiesDialog {
    param([string]$groupName)
    
    Debug-Log ("DEBUG: Showing group properties dialog for: $groupName") -Type "info"
    
    if (-not $groupName) {
        Debug-Log ("ERROR: Group name is null") -Type "error"
        return
    }
    
    # Find the group
    $group = $Global:Groups | Where-Object { $_.Name -eq $groupName } | Select-Object -First 1
    
    if (-not $group) {
        Show-Modal "Not Found" "Group '$groupName' not found"
        return
    }
    
    # Create buttons FIRST
    $btnOK = [Terminal.Gui.Button]::new("OK")
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnApply = [Terminal.Gui.Button]::new("Apply")
    
    # Create dialog WITH buttons
    $dlg = [Terminal.Gui.Dialog]::new("Group Properties - $groupName", 90, 35, $btnOK, $btnCancel, $btnApply)
    
    # TabView
    $tabView = [Terminal.Gui.TabView]::new()
    $tabView.X = 0
    $tabView.Y = 0
    $tabView.Width = [Terminal.Gui.Dim]::Fill()
    $tabView.Height = [Terminal.Gui.Dim]::Fill(3)
    
    # ==================== General Tab ====================
    $generalTab = [Terminal.Gui.TabView+Tab]::new()
    $generalTab.Text = "General"
    $generalView = [Terminal.Gui.View]::new()
    
    $y = 1
    $lbl = [Terminal.Gui.Label]::new("Group name:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtName = [Terminal.Gui.TextField]::new($group.Name ?? ""); $txtName.X=20; $txtName.Y=$y; $txtName.Width=50; $txtName.ReadOnly=$true
    $generalView.Add($txtName)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("Description:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtDesc = [Terminal.Gui.TextField]::new($group.Description ?? ""); $txtDesc.X=20; $txtDesc.Y=$y; $txtDesc.Width=50
    $txtDesc.add_TextChanged({ $script:groupChangesMade = $true })
    $generalView.Add($txtDesc)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("E-mail:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtEmail = [Terminal.Gui.TextField]::new($group.Mail ?? ""); $txtEmail.X=20; $txtEmail.Y=$y; $txtEmail.Width=50
    $txtEmail.add_TextChanged({ $script:groupChangesMade = $true })
    $generalView.Add($txtEmail)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("Group type:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $lblType = [Terminal.Gui.Label]::new($group.GroupCategory ?? "Security"); $lblType.X=20; $lblType.Y=$y
    $generalView.Add($lblType)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("Group scope:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $lblScope = [Terminal.Gui.Label]::new($group.GroupScope ?? "Global"); $lblScope.X=20; $lblScope.Y=$y
    $generalView.Add($lblScope)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("Managed by:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtManagedBy = [Terminal.Gui.TextField]::new($group.ManagedBy ?? ""); $txtManagedBy.X=20; $txtManagedBy.Y=$y; $txtManagedBy.Width=50
    $txtManagedBy.add_TextChanged({ $script:groupChangesMade = $true })
    $generalView.Add($txtManagedBy)
    
    $generalTab.View = $generalView
    $tabView.AddTab($generalTab, $false)
    
    # ==================== Members Tab ====================
    $membersTab = [Terminal.Gui.TabView+Tab]::new()
    $membersTab.Text = "Members"
    $membersView = [Terminal.Gui.View]::new()
    
    $members = $Global:Users | Where-Object { $_.Groups -contains $groupName } | ForEach-Object { $_.Name } | Sort-Object
    
    $lbl = [Terminal.Gui.Label]::new("Members:"); $lbl.X=2; $lbl.Y=1; $membersView.Add($lbl)
    $lstMembers = [Terminal.Gui.ListView]::new()
    $lstMembers.SetSource($members)
    $lstMembers.X=2; $lstMembers.Y=3; $lstMembers.Width=[Terminal.Gui.Dim]::Fill(2); $lstMembers.Height=[Terminal.Gui.Dim]::Fill(4)
    $membersView.Add($lstMembers)
    
    $btnAddMember = [Terminal.Gui.Button]::new("Add...")
    $btnAddMember.X=2; $btnAddMember.Y=[Terminal.Gui.Pos]::Bottom($lstMembers)+1
    $btnAddMember.add_Clicked({
        Show-Modal "Add Member" "Add member functionality - to be implemented"
        $script:groupChangesMade = $true
    }.GetNewClosure())
    $membersView.Add($btnAddMember)
    
    $btnRemoveMember = [Terminal.Gui.Button]::new("Remove")
    $btnRemoveMember.X=[Terminal.Gui.Pos]::Right($btnAddMember)+2
    $btnRemoveMember.Y=[Terminal.Gui.Pos]::Bottom($lstMembers)+1
    $btnRemoveMember.add_Clicked({
        if ($lstMembers.SelectedItem -eq -1) {
            Show-Modal "No Selection" "Please select a member to remove"
            return
        }
        
        $selectedMember = $members[$lstMembers.SelectedItem]
        $result = [Terminal.Gui.MessageBox]::Query(60, 8, "Confirm", 
            "Remove '$selectedMember' from group '$groupName'?", 
            @("Yes", "No"))
        
        if ($result -eq 0) {
            $memberUser = $Global:Users | Where-Object { $_.Name -eq $selectedMember } | Select-Object -First 1
            if ($memberUser -and $memberUser.Groups) {
                $memberUser.Groups = $memberUser.Groups | Where-Object { $_ -ne $groupName }
                
                $members = $Global:Users | Where-Object { $_.Groups -contains $groupName } | ForEach-Object { $_.Name } | Sort-Object
                [Terminal.Gui.Application]::MainLoop.Invoke({
                    $lstMembers.SetSource($members)
                })
                
                $script:groupChangesMade = $true
                Debug-Log ("DEBUG: Removed $selectedMember from group $groupName") -Type "info"
            }
        }
    }.GetNewClosure())
    $membersView.Add($btnRemoveMember)
    
    $membersTab.View = $membersView
    $tabView.AddTab($membersTab, $false)
    
    $dlg.Add($tabView)
    
    $fields = @{
        txtDesc = $txtDesc
        txtEmail = $txtEmail
        txtManagedBy = $txtManagedBy
    }
    
    $btnOK.add_Clicked({
        if (Apply-GroupChanges -group $group -fields $fields) {
            [Terminal.Gui.Application]::RequestStop()
        }
    }.GetNewClosure())
    
    $btnCancel.add_Clicked({
        if ($script:groupChangesMade) {
            $result = [Terminal.Gui.MessageBox]::Query(60, 8, "Unsaved Changes", 
                "You have unsaved changes. Discard them?", 
                @("Yes", "No"))
            if ($result -eq 0) {
                $script:groupChangesMade = $false
                [Terminal.Gui.Application]::RequestStop()
            }
        } else {
            [Terminal.Gui.Application]::RequestStop()
        }
    }.GetNewClosure())
    
    $btnApply.add_Clicked({
        Apply-GroupChanges -group $group -fields $fields
    }.GetNewClosure())
    
    Debug-Log ("DEBUG: Show-GroupPropertiesDialog running") -Type "info"
    [Terminal.Gui.Application]::Run($dlg)
    Debug-Log ("DEBUG: Show-GroupPropertiesDialog completed") -Type "info"
}

# ----------------------- Show Computer Properties ----------------------
function Show-ComputerPropertiesDialog {
    param([string]$computerName)
    
    Debug-Log ("DEBUG: Showing computer properties for: $computerName") -Type "info"
    
    $computer = $Global:Computers | Where-Object { $_.Name -eq $computerName } | Select-Object -First 1
    
    if (-not $computer) {
        Show-Modal "Not Found" "Computer '$computerName' not found"
        return
    }
    
    $status = if ($computer.Enabled) { "Enabled" } else { "Disabled" }
    $lastLogon = if ($computer.LastLogonDate) { $computer.LastLogonDate.ToString('g') } else { "Never" }
    
    $msg = "Computer: $($computer.Name)`n" +
           "Type: $($computer.ComputerType)`n" +
           "Location: $($computer.Location)`n" +
           "IP Address: $($computer.IPv4Address)`n" +
           "Operating System: $($computer.OperatingSystem)`n" +
           "OS Version: $($computer.OperatingSystemVersion)`n" +
           "Status: $status`n" +
           "Last Logon: $lastLogon`n" +
           "Description: $($computer.Description)`n`n" +
           "(Demo Mode)"
    
    Show-Modal "Computer Properties" $msg
}

# ------------------------- Context Menu Handler ------------------------
function Show-ContextMenu {
    param(
        [string]$objectName,
        [string]$objectType
    )
    
$selName = $Global:tree.SelectedObject.Text
Debug-Log ("DEBUG: Selected object text: '$selName'") -Type "info"

# Determine what type of object this is FIRST (before cleaning the name)
$selType = if ($selName -like "(U)*") {"user"} 
           elseif ($selName -like "(DC)*") {"dc"} 
           else {"group"}
Debug-Log ("DEBUG: Detected type: $selType") -Type "info"

# NOW clean the name - remove prefixes like "(U) " or "(DC) " and status icons
$cleanName = $selName -replace '^\([^)]+\)\s*', '' -replace '^[○⊗🔒]\s*', ''
Debug-Log ("DEBUG: Cleaned name after removing prefix: '$cleanName'") -Type "info"

# Extract just the name part if it has [SITE] suffix
if ($cleanName -match '^(.+?)\s+\[.+\]$') {
    $cleanName = $matches[1].Trim()
    Debug-Log ("DEBUG: Extracted name from [SITE] format: '$cleanName'") -Type "info"
}

Debug-Log ("DEBUG: Final cleaned name: '$cleanName'") -Type "info"
    
    # Build menu items array for display
    $menuText = @()

    ################################### right-click or right click or rightclick This is where to clena the code ##########################
    ################################### so it calls the same functions as the menu for less arsing about/bugs    ##########################
    
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
    
    Debug-Log ("DEBUG: Menu item selected: $selected") -Type "info"
    
    [Terminal.Gui.Application]::RequestStop()
    
    if ($selected -ne "---") {
        Debug-Log ("DEBUG: Context menu selected: $selected") -Type "info"
        
        switch ($selected) {
            "Properties" { 
                Debug-Log ("DEBUG: Properties selected for type: isUser=$isUser, isGroup=$isGroup, isDC=$isDC") -Type "info"
                
                if ($isUser) {
                    Debug-Log ("DEBUG: Looking for user: $cleanName") -Type "info"
                    $user = $Global:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                    
                    if ($user) { 
                        Debug-Log ("DEBUG: Found user, calling Show-UserPropertiesDialog") -Type "info"
                        Debug-Log ("DEBUG: User type: $($user.GetType().Name)") -Type "info"
                        Debug-Log ("DEBUG: User Name: $($user.Name)") -Type "info"
                        
                        try {
                            Show-UserPropertiesDialog -user $user  # REMOVED -Global $Global
                            Debug-Log ("DEBUG: Show-UserPropertiesDialog returned") -Type "info"
                        } catch {
                            Debug-Log ("ERROR: Exception showing properties: $($_.Exception.Message)") -Type "error"
                            Debug-Log ("ERROR: Stack trace: $($_.ScriptStackTrace)") -Type "error"
                            Show-Modal "Error" "Failed to show properties:`n$($_.Exception.Message)"
                        }
                    } else {
                        Debug-Log ("ERROR: User '$cleanName' not found in Global:Users") -Type "error"
                        Show-Modal "Not Found" "User '$cleanName' not found"
                    }
                } elseif ($isGroup) {
                    Debug-Log ("DEBUG: Showing group properties") -Type "info"
                    Show-GroupPropertiesDialog -groupName $cleanName
                } elseif ($isDC) {
                    Debug-Log ("DEBUG: Showing DC properties") -Type "info"
                    Show-DCPropertiesDialog -dcName $cleanName
                } else {
                    Debug-Log ("DEBUG: Showing generic properties") -Type "info"
                    $msg = "Object: $cleanName`nType: $objectType"
                    Show-Modal "Properties" $msg
                }
            }
            "Reset Password" { Show-ResetPasswordDialog -userName $cleanName }
            "Disable Account" { Toggle-UserAccount -userName $cleanName -disable $true }
            "Enable Account" { Toggle-UserAccount -userName $cleanName -disable $false }
            "Move to OU..." { Show-MoveObjectDialog -objectName $cleanName -objectType "User" }
            "Delete" { Show-DeleteObjectDialog -objectName $cleanName -objectType $objectType }
            "Add Member..." { Show-AddGroupMemberDialog -groupName $cleanName }
            "Remove Member..." { Show-RemoveGroupMemberDialog -groupName $cleanName }
            "New Object..." { Show-NewObjectWizard }
            "Check Replication" { Check-DCReplication -dcName $cleanName }
            "Refresh" { Refresh-TreeData }
        }
    }
})
    
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $contextDialog.AddButton($btnCancel)
    
    [Terminal.Gui.Application]::Run($contextDialog)
}

# Debug version of Show-Properties
function Show-Properties {
    Debug-Log ("DEBUG: Show-Properties called") -Type "info"
    
    if (-not $Global:tree.SelectedObject) { 
        Debug-Log ("DEBUG: No object selected") -Type "info"
        Show-Modal "Debug" "No object selected in tree"
        return 
    }
    
    $selName = $Global:tree.SelectedObject.Text
    Debug-Log ("DEBUG: Selected object text: '$selName'") -Type "info"
    
    # Parse the tree text to get clean name and type
    $objInfo = Get-CleanObjectInfo -treeText $selName
    $cleanName = $objInfo.Name
    $selType = $objInfo.Type
    
    # Route to appropriate dialog
    if ($selType -eq "user") {
        $selUser = $Global:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
        
        if ($selUser) {
            Debug-Log ("DEBUG: User found, calling Show-UserPropertiesDialog") -Type "info"
            try {
                Show-UserPropertiesDialog -user $selUser
            } catch {
                Debug-Log ("ERROR: Exception in Show-UserPropertiesDialog: $_") -Type "info"
                Show-Modal "Error" "Failed to show properties`n$($_.Exception.Message)`n`nCheck console for details"
            }
        } else {
            Debug-Log ("DEBUG: User NOT found in Global:Users") -Type "info"
            Show-Modal "Not Found" "User '$cleanName' not found"
        }
    } elseif ($selType -eq "group") {
        Debug-Log ("DEBUG: Group type selected: $cleanName") -Type "info"
        Show-GroupPropertiesDialog -groupName $cleanName
    } elseif ($selType -eq "ou") {
        Debug-Log ("DEBUG: OU type selected: $cleanName") -Type "info"
        Show-OUPropertiesDialog -ouName $cleanName
    } elseif ($selType -eq "dc") {
            Debug-Log ("DEBUG: DC type selected: $cleanName") -Type "info"
            Show-DCPropertiesDialog -dcName $cleanName

    } elseif ($selType -eq "computer") {
            Debug-Log ("DEBUG: Computer type selected: $cleanName") -Type "info"
            Show-ComputerPropertiesDialog -computerName $cleanName
    } else {
        Debug-Log ("DEBUG: Selected object type $selType not handled yet.") -Type "info"
    }
}

# Additional debug helper - call this to verify your demo data loaded correctly
function Test-DemoData {
    Debug-Log ("========== DEMO DATA CHECK ==========") -Type "info"
    Debug-Log ("Global:Users count: $($Global:Users.Count)") -Type "info"
    Debug-Log ("Global:DCs count: $($Global:DCs.Count)") -Type "info"
    Debug-Log ("") -Type "info"
    Debug-Log ("Users in memory:") -Type "info"
    foreach ($u in $Global:Users) {
        $locked = if ($u.Locked) { "🔒" } else { "" }
        $disabled = if ($u.Disabled) { "⊗" } else { "○" }
        Debug-Log ("  $disabled$locked $($u.Name) - Groups: $($u.Groups -join ', ')") -Type "info"
    }
    Debug-Log ("=====================================") -Type "info"
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
            Show-Modal "Error" "Passwords do not match!"
            return
        }
        
        if ($pwd1.Length -lt 8) {
            Show-Modal "Error" "Password must be at least 8 characters!"
            return
        }
        
        try {
            if ($Global:DemoMode) {
                Debug-Log ("DEBUG: Password reset for $userName (demo mode)") -Type "info"
                Show-Modal "Success" "Password reset successfully (demo mode)"
            } else {
                $secPwd = ConvertTo-SecureString -String $pwd1 -AsPlainText -Force
                Set-ADAccountPassword -Identity $userName -NewPassword $secPwd -Reset -ErrorAction Stop
                if ($chkMustChange.Checked) {
                    Set-ADUser -Identity $userName -ChangePasswordAtLogon $true -ErrorAction Stop
                }
                Show-Modal "Success" "Password reset successfully"
            }
            [Terminal.Gui.Application]::RequestStop()
        } catch {
            $errMsg = $_.Exception.Message
            Show-Modal "Error" "Failed to reset password:`n$errMsg"
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
                    Debug-Log ("DEBUG: Account $userName $action`d (demo mode)") -Type "info"
                }
            } else {
                if ($disable) {
                    Disable-ADAccount -Identity $userName -ErrorAction Stop
                } else {
                    Enable-ADAccount -Identity $userName -ErrorAction Stop
                }
            }
            Show-Modal "Success" "Account $action`d successfully"
            Refresh-TreeData
        } catch {
            $errMsg = $_.Exception.Message
            Show-Modal "Error" "Failed to $action account:`n$errMsg"
        }
    }
}

function Show-DeleteObjectDialog {
    param([string]$objectName, [string]$objectType)
    
    $result = [Terminal.Gui.MessageBox]::Query(60, 9, "Delete $objectType", "WARNING: Are you sure you want to delete:`n$objectName`n`nThis action cannot be undone!", "Delete", "Cancel")
    
    if ($result -eq 0) {
      Show-Modal "Delete" "Delete functionality coming soon"
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
        Debug-Log ("DEBUG: Selection mode ENABLED") -Type "info"
        [Terminal.Gui.MessageBox]::Query(60, 8, "Selection Mode", 
            "Selection mode enabled!`n`nClick objects to select/deselect them.`nPress Ctrl+A to select all.`nPress Ctrl+D to deselect all.", 
            "OK") | Out-Null
    } else {
        Debug-Log ("DEBUG: Selection mode DISABLED") -Type "info"
        $Global:SelectedObjects = @()
        Build-Tree -domain $Global:CurrentDomain
        Update-FilterStatusLabel -label $Global:FilterStatusLabel
    }
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
    
    if (-not $Global:tree.SelectedObject) { return }
    
    $selName = $Global:tree.SelectedObject.Text
    
    # Check if in selection mode
    if ($Global:SelectionMode) {
        # Toggle selection
        if ($Global:SelectedObjects -contains $selName) {
            # Deselect
            $Global:SelectedObjects = $Global:SelectedObjects | Where-Object { $_ -ne $selName }
            Debug-Log ("DEBUG: Deselected $selName") -Type "info"
        } else {
            # Select
            $Global:SelectedObjects += $selName
            Debug-Log ("DEBUG: Selected $selName") -Type "info"
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
      Show-Modal "Selection Mode" "Enable selection mode first (Ctrl+S)"
        return
    }
    
    $Global:SelectedObjects = @()
    
    # Get all users from tree
    foreach ($user in $Global:Users) {
        $statusIcon = if ($user.Disabled) { "⊗" } else { "○" }
        $displayName = "(U) $statusIcon $($user.Name)"
        $Global:SelectedObjects += $displayName
    }
    
    Debug-Log ("DEBUG: Selected all users ($($Global:SelectedObjects.Count))") -Type "info"
    Update-SelectionPanel -panel $selectionPanel
    Show-Modal "Selected All" "Selected $($Global:SelectedObjects.Count) users"
}

function Deselect-AllObjects {
    $Global:SelectedObjects = @()
    Debug-Log ("DEBUG: Deselected all objects") -Type "info"
    Update-SelectionPanel -panel $selectionPanel
}

# ------------------------- Bulk Disable/Enable ------------------------
function Invoke-BulkDisableEnable {
    param([bool]$disable)
    
    if ($Global:SelectedObjects.Count -eq 0) {
      Show-Modal "No Selection" "No objects selected. Select objects first."
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
                        Debug-Log ("DEBUG: $action`d $cleanName (demo mode)") -Type "info"
                    }
                } else {
                    if ($disable) {
                        Disable-ADAccount -Identity $cleanName -ErrorAction Stop
                    } else {
                        Enable-ADAccount -Identity $cleanName -ErrorAction Stop
                    }
                    $successCount++
                    Debug-Log ("DEBUG: $action`d $cleanName in AD") -Type "info"
                }
            } catch {
                $failCount++
                $errors += "$cleanName`: $($_.Exception.Message)"
                Debug-Log ("DEBUG: Failed to $action $cleanName`: $_") -Type "info"
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
        
        Show-Modal "Bulk Action Complete" $msg
        
        # Refresh tree
        if (-not $Global:DemoMode) {
            Load-DomainData -domain $Global:CurrentDomain
        }
        Build-Tree -domain $Global:CurrentDomain
        Update-FilterStatusLabel -label $Global:FilterStatusLabel
        
        # Clear selection
        $Global:SelectedObjects = @()
        $Global:SelectionMode = $false
        Update-SelectionPanel -panel $selectionPanel
    }
}

# ------------------------- Bulk Move ------------------------
function Invoke-BulkMove {
    if ($Global:SelectedObjects.Count -eq 0) {
      Show-Modal "No Selection" "No objects selected. Select objects first."
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
          Show-Modal "Error" "Please select a target OU"
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
                            Debug-Log ("DEBUG: Moved $cleanName to $targetOU (demo mode)") -Type "info"
                        }
                    } else {
                        $adObject = Get-ADObject -Filter "Name -eq '$cleanName'" -ErrorAction Stop
                        Move-ADObject -Identity $adObject.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
                        $successCount++
                        Debug-Log ("DEBUG: Moved $cleanName to $targetOU in AD") -Type "info"
                    }
                } catch {
                    $failCount++
                    $errors += "$cleanName`: $($_.Exception.Message)"
                    Debug-Log ("DEBUG: Failed to move $cleanName`: $_") -Type "info"
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
            
            Show=-Modal "Bulk Move Complete" $msg
            
            # Refresh tree
            if (-not $Global:DemoMode) {
                Load-DomainData -domain $Global:CurrentDomain
            }
            Build-Tree -domain $Global:CurrentDomain
            Update-FilterStatusLabel -label $Global:FilterStatusLabel
            
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
      Show-Modal "No Selection" "No objects selected. Select objects first."
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
          Show-Modal "Error" "Please select a group"
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
                            Debug-Log ("DEBUG: Added $cleanName to $targetGroup (demo mode)") -Type "info"
                        }
                    } else {
                        Add-ADGroupMember -Identity $targetGroup -Members $cleanName -ErrorAction Stop
                        $successCount++
                        Debug-Log ("DEBUG: Added $cleanName to $targetGroup in AD") -Type "info"
                    }
                } catch {
                    $failCount++
                    Debug-Log ("DEBUG: Failed to add $cleanName`: $_") -Type "info"
                }
            }
            
            [Terminal.Gui.MessageBox]::Query(60, 10, "Bulk Add Complete", 
                "Successfully added $successCount user(s)`nFailed: $failCount", 
                "OK") | Out-Null
            
            # Refresh tree
            Build-Tree -domain $Global:CurrentDomain
            Update-FilterStatusLabel -label $Global:FilterStatusLabel
            
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
    
    Debug-Log ("DEBUG: Selected all users ($($Global:SelectedObjects.Count))") -Type "info"
    Update-SelectionPanel -panel $selectionPanel
    Show-Modal "Selected All" "Selected $($Global:SelectedObjects.Count) users"
}

function Deselect-AllObjects {
    $Global:SelectedObjects = @()
    Debug-Log ("DEBUG: Deselected all objects") -Type "info"
    Update-SelectionPanel -panel $selectionPanel
}

# ------------------------- Bulk Disable/Enable ------------------------
function Invoke-BulkDisableEnable {
    param([bool]$disable)
    
    if ($Global:SelectedObjects.Count -eq 0) {
      Sho-wModal "No Selection" "No objects selected. Select objects first."
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
                        Debug-Log ("DEBUG: $action`d $cleanName (demo mode)") -Type "info"
                    }
                } else {
                    if ($disable) {
                        Disable-ADAccount -Identity $cleanName -ErrorAction Stop
                    } else {
                        Enable-ADAccount -Identity $cleanName -ErrorAction Stop
                    }
                    $successCount++
                    Debug-Log ("DEBUG: $action`d $cleanName in AD") -Type "info"
                }
            } catch {
                $failCount++
                $errors += "$cleanName`: $($_.Exception.Message)"
                Debug-Log ("DEBUG: Failed to $action $cleanName`: $_") -Type "info"
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
            Load-DomainData -domain $Global:CurrentDomain
        }
        Build-Tree -domain $Global:CurrentDomain
        Update-FilterStatusLabel -label $Global:FilterStatusLabel
        
        # Clear selection
        $Global:SelectedObjects = @()
        $Global:SelectionMode = $false
        Update-SelectionPanel -panel $selectionPanel
    }
}

# ------------------------- Bulk Move ------------------------
function Invoke-BulkMove {
    if ($Global:SelectedObjects.Count -eq 0) {
      Show-Modal "No Selection" "No objects selected. Select objects first."
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
          Show-Modal "Error" "Please select a target OU"
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
                            Debug-Log ("DEBUG: Moved $cleanName to $targetOU (demo mode)") -Type "info"
                        }
                    } else {
                        $adObject = Get-ADObject -Filter "Name -eq '$cleanName'" -ErrorAction Stop
                        Move-ADObject -Identity $adObject.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
                        $successCount++
                        Debug-Log ("DEBUG: Moved $cleanName to $targetOU in AD") -Type "info"
                    }
                } catch {
                    $failCount++
                    $errors += "$cleanName`: $($_.Exception.Message)"
                    Debug-Log ("DEBUG: Failed to move $cleanName`: $_") -Type "info"
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
            
            Show-Modal "Bulk Move Complete" $msg
            
            # Refresh tree
            if (-not $Global:DemoMode) {
                Load-DomainData -domain $Global:CurrentDomain
            }
            Build-Tree -domain $Global:CurrentDomain
            Update-FilterStatusLabel -label $Global:FilterStatusLabel
            
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
      Show-Modal "No Selection" "No objects selected. Select objects first."
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
            Show-Modal "Error" "Please select a group"
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
                            Debug-Log ("DEBUG: Added $cleanName to $targetGroup (demo mode)") -Type "info"
                        }
                    } else {
                        Add-ADGroupMember -Identity $targetGroup -Members $cleanName -ErrorAction Stop
                        $successCount++
                        Debug-Log ("DEBUG: Added $cleanName to $targetGroup in AD") -Type "info"
                    }
                } catch {
                    $failCount++
                    Debug-Log ("DEBUG: Failed to add $cleanName`: $_") -Type "info"
                }
            }
            
            [Terminal.Gui.MessageBox]::Query(60, 10, "Bulk Add Complete", 
                "Successfully added $successCount user(s)`nFailed: $failCount", 
                "OK") | Out-Null
            
            # Refresh tree
            Build-Tree -domain $Global:CurrentDomain
            Update-FilterStatusLabel -label $Global:FilterStatusLabel
            
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
    
    # Add Tag property to store references
    $selPanel | Add-Member -MemberType NoteProperty -Name Tag -Value @{
        CountLabel = $lblCount
        ListView = $lstSelected
    } -Force
    
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

# Add/Remove Group Member Implementation
# Replace the placeholder functions in your script

# =====================================================
# ADD GROUP MEMBER
# =====================================================
function Show-AddGroupMemberDialog {
    param([string]$groupName)
    
    Debug-Log ("DEBUG: Add member to group: $groupName") -Type "info"
    
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
            Show-Modal "No Selection" "Please select at least one user (use SPACE to mark multiple, or just highlight one)"
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
    
    Debug-Log ("DEBUG: Remove member from group: $groupName") -Type "info"
    
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
            [Terminal.Gui.MessageBox]::Query(50, 7, "No Selection", "Please select at least one user (use SPACE key)", "OK") | Out-Null
            return
        }
        
        $usersToRemove = $markedIndices | ForEach-Object { $groupMembers[$_] }
        $confirm = [Terminal.Gui.MessageBox]::Query(60, 9, "Confirm Remove", "Remove $($usersToRemove.Count) user(s) from group '$groupName'?", "Yes", "No")
        
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
                        Debug-Log ("DEBUG: Removed $($user.Name) from $groupName (demo mode)") -Type "info"
                    } else {
                        # Production mode: remove from AD
                        Remove-ADGroupMember -Identity $groupName -Members $user.SamAccountName -Confirm:$false -ErrorAction Stop
                        $successCount++
                        Debug-Log ("DEBUG: Removed $($user.SamAccountName) from $groupName in AD") -Type "info"
                    }
                } catch {
                    $failCount++
                    $userName = if ($Global:DemoMode) { $user.Name } else { $user.SamAccountName }
                    $errors += "$userName`: $($_.Exception.Message)"
                    Debug-Log ("ERROR: Failed to remove $userName from group: $_") -Type "info"
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
            
            Show-Modal "Remove Members Complete" $msg
            
            # Rebuild tree
            [Terminal.Gui.Application]::MainLoop.Invoke({
                Build-Tree -domain $Global:CurrentDomain
                Update-FilterStatusLabel -label $Global:FilterStatusLabel
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
    Show-Modal "Replication Check" "Checking replication for $dcName`n(Coming soon)"
}

function Show-Modal { 
    param($title, $msg) 
    [Terminal.Gui.MessageBox]::Query($title, $msg, @("OK")) | Out-Null 
}

function Show-OUPropertiesDialog {
    param([string]$ouName)
    
    Debug-Log ("DEBUG: Showing OU properties dialog for: $ouName") -Type "info"
    
    if (-not $ouName) {
        Debug-Log ("ERROR: OU name is null") -Type "error"
        return
    }
    
    # Find the OU in raw data
    $ou = $Global:rawOUs | Where-Object { $_.Name -eq $ouName } | Select-Object -First 1
    
    # If not found, create a basic one (for existing OUs from user data)
    if (-not $ou) {
        $ou = @{
            Name = $ouName
            Path = ""
            Description = ""
        }
    }
    
    # Create buttons FIRST
    $btnOK = [Terminal.Gui.Button]::new("OK")
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnApply = [Terminal.Gui.Button]::new("Apply")
    
    # Create dialog WITH buttons
    $dlg = [Terminal.Gui.Dialog]::new("OU Properties - $ouName", 80, 20, $btnOK, $btnCancel, $btnApply)
    
    # Create a simple view (no tabs needed for OUs)
    $view = [Terminal.Gui.View]::new()
    $view.X = 0
    $view.Y = 0
    $view.Width = [Terminal.Gui.Dim]::Fill()
    $view.Height = [Terminal.Gui.Dim]::Fill(3)
    
    $y = 1
    
    # OU Name (editable for rename)
    $lbl = [Terminal.Gui.Label]::new("Name:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
    $txtName = [Terminal.Gui.TextField]::new($ou.Name ?? ""); $txtName.X=20; $txtName.Y=$y; $txtName.Width=50
    $txtName.add_TextChanged({ $script:ouChangesMade = $true })
    $view.Add($txtName)
    $y+=2
    
    # Description
    $lbl = [Terminal.Gui.Label]::new("Description:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
    $txtDesc = [Terminal.Gui.TextField]::new($ou.Description ?? ""); $txtDesc.X=20; $txtDesc.Y=$y; $txtDesc.Width=50
    $txtDesc.add_TextChanged({ $script:ouChangesMade = $true })
    $view.Add($txtDesc)
    $y+=2
    
    # Path (read-only)
    $lbl = [Terminal.Gui.Label]::new("Path:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
    $txtPath = [Terminal.Gui.TextField]::new($ou.Path ?? ""); $txtPath.X=20; $txtPath.Y=$y; $txtPath.Width=50; $txtPath.ReadOnly=$true
    $view.Add($txtPath)
    $y+=2
    
    # Show object count in this OU
    $objectCount = ($Global:Users | Where-Object { $_.OU -contains $ouName }).Count
    $lbl = [Terminal.Gui.Label]::new("Contains:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
    $lblCount = [Terminal.Gui.Label]::new("$objectCount objects"); $lblCount.X=20; $lblCount.Y=$y
    $view.Add($lblCount)
    
    $dlg.Add($view)
    
    # Store field references for Apply function
    $fields = @{
        txtName = $txtName
        txtDesc = $txtDesc
        originalName = $ou.Name
    }
    
    # Wire up button actions (NO .GetNewClosure() to avoid scope issues)
    $btnOK.add_Clicked({
        if (Apply-OUChanges -ou $ou -fields $fields) {
            [Terminal.Gui.Application]::RequestStop()
        }
    })
    
    $btnCancel.add_Clicked({
        if ($script:ouChangesMade) {
            $result = [Terminal.Gui.MessageBox]::Query(60, 8, "Unsaved Changes", 
                "You have unsaved changes. Discard them?", 
                @("Yes", "No"))
            if ($result -eq 0) {
                $script:ouChangesMade = $false
                [Terminal.Gui.Application]::RequestStop()
            }
        } else {
            [Terminal.Gui.Application]::RequestStop()
        }
    })
    
    $btnApply.add_Clicked({
        Apply-OUChanges -ou $ou -fields $fields
    })
    
    Debug-Log ("DEBUG: Show-OUPropertiesDialog running") -Type "info"
    [Terminal.Gui.Application]::Run($dlg)
    Debug-Log ("DEBUG: Show-OUPropertiesDialog completed") -Type "info"
}

function Apply-OUChanges {
    param($ou, $fields)
    
    $originalName = $fields.originalName
    $newName = $fields.txtName.Text.ToString()
    $newDesc = $fields.txtDesc.Text.ToString()
    
    Debug-Log ("DEBUG: Applying changes for OU: $originalName -> $newName") -Type "info"
    
    try {
        if ($Global:DemoMode) {
            # Check if renaming
            $isRename = $originalName -ne $newName
            
            if ($isRename) {
                Debug-Log ("DEBUG: Renaming OU from '$originalName' to '$newName'") -Type "info"
                
                # Update the OU itself
                $ou.Name = $newName
                $ou.Description = $newDesc
                
                # Update all users that reference this OU
                foreach ($user in $Global:Users) {
                    if ($user.OU -contains $originalName) {
                        Debug-Log ("DEBUG: Updating user $($user.Name) OU reference") -Type "info"
                        $user.OU = $user.OU | ForEach-Object { if ($_ -eq $originalName) { $newName } else { $_ } }
                    }
                }
                
                # Update raw users too
                foreach ($rawUser in $Global:rawUsers) {
                    if ($rawUser.OU -contains $originalName) {
                        $rawUser.OU = $rawUser.OU | ForEach-Object { if ($_ -eq $originalName) { $newName } else { $_ } }
                    }
                }
                
                # Update in rawOUs
                $rawOU = $Global:rawOUs | Where-Object { $_.Name -eq $originalName } | Select-Object -First 1
                if ($rawOU) {
                    $rawOU.Name = $newName
                    $rawOU.Description = $newDesc
                }
                
                # Update the fields for next apply
                $fields.originalName = $newName
                
                Show-Modal "Success" "OU renamed from '$originalName' to '$newName' (demo mode)"
            } else {
                # Just updating description
                $ou.Description = $newDesc
                
                $rawOU = $Global:rawOUs | Where-Object { $_.Name -eq $originalName } | Select-Object -First 1
                if ($rawOU) {
                    $rawOU.Description = $newDesc
                }
                
                Show-Modal "Success" "OU changes applied (demo mode)"
            }
            
            Debug-Log ("SUCCESS: OU changes applied (demo mode)") -Type "info"
            
            # Refresh the tree to show changes
            Refresh-Data -domain $Global:CurrentDomain
            
        } else {
            # Production mode - use Rename-ADObject or Set-ADOrganizationalUnit
            $isRename = $originalName -ne $newName
            
            if ($isRename) {
                # Get the actual AD OU object
                $adOU = Get-ADOrganizationalUnit -Filter "Name -eq '$originalName'" -ErrorAction Stop | Select-Object -First 1
                
                if ($adOU) {
                    # Rename the OU
                    Rename-ADObject -Identity $adOU.DistinguishedName -NewName $newName -ErrorAction Stop
                    
                    # Update description if provided
                    if ($newDesc) {
                        Set-ADOrganizationalUnit -Identity "OU=$newName,$($adOU.DistinguishedName -replace '^OU=[^,]+,')" -Description $newDesc -ErrorAction Stop
                    }
                    
                    Debug-Log ("SUCCESS: OU renamed in AD from $originalName to $newName") -Type "info"
                    Show-Modal "Success" "OU renamed successfully"
                } else {
                    throw "OU not found in AD"
                }
            } else {
                # Just update description
                $adOU = Get-ADOrganizationalUnit -Filter "Name -eq '$originalName'" -ErrorAction Stop | Select-Object -First 1
                if ($adOU) {
                    Set-ADOrganizationalUnit -Identity $adOU.DistinguishedName -Description $newDesc -ErrorAction Stop
                    Debug-Log ("SUCCESS: OU description updated in AD") -Type "info"
                    Show-Modal "Success" "OU changes applied"
                }
            }
            
            # Refresh from AD
            Refresh-Data -domain $Global:CurrentDomain
        }
        
        $script:ouChangesMade = $false
        return $true
        
    } catch {
        Debug-Log ("ERROR: Failed to apply OU changes: $($_.Exception.Message)") -Type "error"
        Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)"
        return $false
    }
}

## -------{ Applicaiton execution starts formaly here }-------
## -------------------------------{ TreeView }------------------------------
#$tree = [Terminal.Gui.TreeView]::new()
$Global:tree.X=0; $Global:tree.Y=1; $Global:tree.Width=80; $Global:tree.Height=[Terminal.Gui.Dim]::Fill()

## ------------------------- Load Domain or Demo Data FIRST ------------------------
Load-DomainData -domain $Global:CurrentDomain

# (Optional) Debug after loading
Debug-Log ("POST-LOAD DEBUG: Domains: $($Global:Domains.Count)") -Type "info"
Debug-Log  ("POST-LOAD DEBUG: Users: $($Global:Users.Count)") -Type "info"
Debug-Log ("POST-LOAD DEBUG: Computers: $($Global:Computers.Count)") -Type "info"
Debug-Log ("POST-LOAD DEBUG: DCs: $($Global:DCs.Count)") -Type "info"
Debug-Log ("POST-LOAD DEBUG: Objects: $($Global:ADObjects.Count)") -Type "info"

## ------------------------- Build initial tree ------------------------
Build-Tree -domain $Global:CurrentDomain

## ------------------------- Update status label ------------------------
Update-FilterStatusLabel -label $Global:FilterStatusLabel

# ------------------------- Tree Mouse Handler ------------------------
# Add this to tree setup after creating ${tree}:

# Add mouse event handler for right-click
$Global:tree.add_MouseClick({
    param($mouseEvent)
    
    Debug-Log ("DEBUG: Mouse event type: $($mouseEvent.GetType().Name)") -Type "info"
    Debug-Log ("DEBUG: Mouse event properties: $($mouseEvent | Get-Member -MemberType Property | Select-Object -ExpandProperty Name)") -Type "info"
    
    # Try different ways to access the mouse button
    $isRightClick = $false
    
    # Method 1: Direct Flags property
    if ($mouseEvent.PSObject.Properties['Flags']) {
        Debug-Log ("DEBUG: Flags = $($mouseEvent.Flags)") -Type "info"
        $isRightClick = $mouseEvent.Flags -band [Terminal.Gui.MouseFlags]::Button3Clicked
    }
    
    # Method 2: MouseEvent property
    if ($mouseEvent.PSObject.Properties['MouseEvent']) {
        Debug-Log ("DEBUG: MouseEvent.Flags = $($mouseEvent.MouseEvent.Flags)") -Type "info"
        $isRightClick = $mouseEvent.MouseEvent.Flags -band [Terminal.Gui.MouseFlags]::Button3Clicked
    }
    
    # Method 3: Check if it's Button3
    if ($mouseEvent.PSObject.Properties['Button']) {
        Debug-Log ("DEBUG: Button = $($mouseEvent.Button)") -Type "info"
        $isRightClick = $mouseEvent.Button -eq 3
    }
    
    Debug-Log ("DEBUG: Is right-click: $isRightClick") -Type "info"
    
    if ($isRightClick) {
        Debug-Log ("DEBUG: Right-click detected") -Type "info"
        
        if ($Global:tree.SelectedObject) {
            $selectedNode = $Global:tree.SelectedObject
            $nodeName = $selectedNode.Text
            
            Debug-Log ("DEBUG: Right-clicked on node: $nodeName") -Type "info"
            
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


## -------{ Application execution starts formally here }-------
## -------------------------------{ TreeView Setup }------------------------------
$Global:tree.X=0; $Global:tree.Y=1; $Global:tree.Width=80; $Global:tree.Height=[Terminal.Gui.Dim]::Fill()

## ------------------------- Load Domain or Demo Data FIRST ------------------------
Load-DomainData -domain $Global:CurrentDomain

# (Optional) Debug after loading
Debug-Log ("POST-LOAD DEBUG: Domains: $($Global:Domains.Count)") -Type "info"
Debug-Log ("POST-LOAD DEBUG: Users: $($Global:Users.Count)") -Type "info"
Debug-Log ("POST-LOAD DEBUG: Computers: $($Global:Computers.Count)") -Type "info"
Debug-Log ("POST-LOAD DEBUG: DCs: $($Global:DCs.Count)") -Type "info"
Debug-Log ("POST-LOAD DEBUG: Objects: $($Global:ADObjects.Count)") -Type "info"

## ------------------------- Build initial tree ------------------------
Build-Tree -domain $Global:CurrentDomain

## ------------------------- Update status label ------------------------
Update-FilterStatusLabel -label $Global:FilterStatusLabel

##############################################################################
## ----------------------------{ Start program }----------------------------
## ------------------------{ Initialize Terminal.Gui }----------------------
[Terminal.Gui.Application]::Init()
$top = [Terminal.Gui.Application]::Top

# Get the selected theme
Debug-Log ("Applying theme: $Theme") -Type "info"
$themeData = Get-Theme -mode $Theme

if ($null -eq $themeData) {
    Debug-Log ("WARNING: themeData is null, using default colors") -Type "warn"
}

# Apply theme to top level first
if ($themeData -and $themeData.Global) {
  $top.ColorScheme = $themeData.Global
  Debug-Log ("Applied Global color scheme to top") -Type "info"
} else {
  Debug-Log ("WARNING: No Global color scheme available") -Type "warn"
}

## ------------------------------{ Main Window }----------------------------
$win = [Terminal.Gui.Window]::new("$($Global:ProjectName) — Active Directory ${BuildVersion} ${Global:FruitName}")
$win.X=0; $win.Y=0; $win.Width=[Terminal.Gui.Dim]::Fill(); $win.Height=[Terminal.Gui.Dim]::Fill(1)
$top.Add($win)

$filterPanel = Create-FilterPanel
if ($null -eq $filterPanel) {
    Debug-Log ("ERROR: filterPanel is null!") -Type "error"
} else {
    $win.Add($filterPanel)
}

$Global:FilterStatusLabel = Create-FilterStatusLabel
if ($null -eq $Global:FilterStatusLabel) {
    Debug-Log ("ERROR: FilterStatusLabel is null!") -Type "error"
} else {
    $win.Add($Global:FilterStatusLabel)
    Debug-Log ("FilterStatusLabel created and added successfully") -Type "info"
}

$selectionPanel = Create-SelectionPanel
if ($null -eq $selectionPanel) {
    Debug-Log ("ERROR: selectionPanel is null!") -Type "error"
} else {
    $win.Add($selectionPanel)
}

## Update AFTER UI exists
[Terminal.Gui.Application]::Top.SetNeedsDisplay()
Update-FilterStatusLabel -label $Global:FilterStatusLabel

## ------------------------- Status Bar ------------------------
# Global status system - Use integer codes to avoid CLS compliance issues
# Key codes: 0=Null, 0x1000000A=F1, 0x10000012=F9, 0x10000013=F10, 0x10000015=F12

$initialStatus = "Forest: $($Global:ForestName) | Objects: $($Global:ADObjects.Count) | Ready"

$Global:StatusItem = [Terminal.Gui.StatusItem]::new(0, $initialStatus, $null)

$Global:statusSpinner = @('|', '/', '-', '\')
$Global:statusSpinnerIndex = 0
$Global:statusBaseMessage = ""

## Create status bar with all items using integer key codes
$Global:StatusBar = [Terminal.Gui.StatusBar]::new(@(
  [Terminal.Gui.StatusItem]::new(0x1000000A,"~F1~ Help",{ Show-Modal "Shortcuts" "F1 - Help`n`nF2 - Password GeneratorF3 - New`nF5 - Refresh`n`nF6 - Themes`nF7 - Search`nF7 - TODO`nF9 - Menus`nF10 - Quit`nF11 Full-Screen`nF12 Redraw" }),
  [Terminal.Gui.StatusItem]::new(0x1000000B,"~F3~ Password Generator",{  Generate-RandomPasswor }),
  [Terminal.Gui.StatusItem]::new(0x1000000C,"~F3~ New",{ Show-NewObjectWizard }),
  [Terminal.Gui.StatusItem]::new(0x1000000D,"~F5~ Refresh",{ Refresh-Data }),
  [Terminal.Gui.StatusItem]::new(0x1000000E,"~F6~ Themes",{ Show-ThemeSelector }),
  [Terminal.Gui.StatusItem]::new(0x1000000F,"~F7~ Search",{ Show-ADSearchDialog }),
  [Terminal.Gui.StatusItem]::new(0x10000012,"~F9~ Menus",{ }), ## Can't use F9 it shows dropdown menys
  [Terminal.Gui.StatusItem]::new(0x10000013,"~F10~ Quit",{ [Terminal.Gui.Application]::RequestStop() }),
  [Terminal.Gui.StatusItem]::new(0x10000014,"~F11~ Full-Screen",{ }),
  [Terminal.Gui.StatusItem]::new(0x10000015,"~F12~ Redraw",{ [Terminal.Gui.Application]::Refresh() }),
  $Global:StatusItem  # Dynamic status on the right
))
$top.Add($StatusBar)

## ------------------------- Menu ------------------------
$mFile = [Terminal.Gui.MenuItem]::new("_Exit","Exit application (F10)",[Action]{ [Terminal.Gui.Application]::RequestStop() })
$mNew = [Terminal.Gui.MenuItem]::new("New Object","Create a new object (F3)",[Action]{ Show-NewObjectWizard })
$mProps = [Terminal.Gui.MenuItem]::new("_Properties","Edit selected properties",[Action]{ Show-Properties }) ## No paramters!
    
$mUndo = [Terminal.Gui.MenuItem]::new("_Undo","Undo last action",[Action]{ Debug-Log ("DEBUG: Undo placeholder") -Type "info" }) 
$mChangeDomain = [Terminal.Gui.MenuItem]::new("Change _Domain","Select domain",[Action]{ Show-ChangeDomainDialog })
$mChangeDC = [Terminal.Gui.MenuItem]::new("Change _Domain Controller","Select DC",[Action]{ Show-ChangeDCDialog })
$mSearchAD = [Terminal.Gui.MenuItem]::new("_Search AD","Search Active Directory (F7)",[Action]{ Show-ADSearchDialog })
$mRefresh = [Terminal.Gui.MenuItem]::new("_Refresh","Refresh AD data (F5)",[Action]{
  Debug-Log ("DEBUG: Refresh menu clicked - scheduling refresh...") -Type "info"
    
  ## Schedule the actual refresh to happen AFTER this menu event completes
  [Terminal.Gui.Application]::MainLoop.AddTimeout([TimeSpan]::FromMilliseconds(100), {
    Debug-Log ("DEBUG: Timeout callback - starting refresh...") -Type "info"
    try {
      Update-Status "Refreshing..." -spinner
      $result = Refresh-Data -domain $Global:CurrentDomain -RebuildTree
            
      if ($result) {
        Update-Status "Refresh complete" -final
      } else {
        Update-Status "Refresh failed" -final
      }
            
      Debug-Log ("DEBUG: Refresh completed with result: $result") -Type "info"
            
    } catch {
      Debug-Log ("ERROR: Refresh crashed: $($_.Exception.Message)") -Type "info"
      Update-Status "Refresh error" -final
    }
        
    return $false  # Don't repeat the timeout
  })
    
  Debug-Log ("DEBUG: Refresh scheduled") -Type "info"
})

$mQuickFilter = [Terminal.Gui.MenuItem]::new("_Quick Filter","Apply quick filters",[Action]{ Show-QuickFilterDialog })
$mSelectionMode = [Terminal.Gui.MenuItem]::new("_Selection Mode (Ctrl+S)","Toggle selection mode",[Action]{ Toggle-SelectionMode })
$mSelectAll = [Terminal.Gui.MenuItem]::new("Select _All (Ctrl+A)","Select all objects",[Action]{ Select-AllObjects })
$mDeselectAll = [Terminal.Gui.MenuItem]::new("_Deselect All (Ctrl+D)","Deselect all objects",[Action]{ Deselect-AllObjects })
$mBulkAddGroup = [Terminal.Gui.MenuItem]::new("Add to _Group...","Add selected users to group",[Action]{ Invoke-BulkAddToGroup })
$mPasswordGenerator = [Terminal.Gui.MenuItem]::new("_Password Generator","Password Generator (F2)",[Action]{ Generate-RandomPassword })
$mADHealth = [Terminal.Gui.MenuItem]::new("_AD Health Status","AD Health And Replicaiton Status",[Action]{ Get-ADHealth })
$mShortcuts = [Terminal.Gui.MenuItem]::new("_Shortcuts","Keyboard shortcuts (F1)",[Action]{ Show-Modal "Shortcuts" "F1 - Help`nF2 - Password Generator`nF3 - New`nF5 - Refresh`n`nF6 - Themes`nF7 - Search`nF7 - TODO`nF9 - Menus`nF10 - Quit`nF11 Full-Screen`nF12 Redraw" })
$mAboutDSATUI = [Terminal.Gui.MenuItem]::new("_About","About $($Global:ProjectName)",[Action]{ Show-Modal "About" "$($Global:ProjectName)`n`nCodename: $($Global:FruitName)`nv$($Global:BuildVersion) STABLE`nGPL-3 Copyleft`nBy Knightmare2600 (https://github.com/knightmare2600)" })
$mWhyBlaabaer = [Terminal.Gui.MenuItem]::new("Why _Blaabaer?","Why the $($Global:FruitName) codename?",[Action]{ Show-BlaabaerInfo })
$mTheme = [Terminal.Gui.MenuItem]::new("_Theme","Change color theme (F6)",[Action]{ })

## ------------------------- Menu Bar -------------------------
$menu = [Terminal.Gui.MenuBar]::new(@(
  [Terminal.Gui.MenuBarItem]::new("_File", @($mRefresh, $mTheme, $mFile)),
  [Terminal.Gui.MenuBarItem]::new("_Action", @($mNew, $mProps, $mQuickFilter, $mUndo, $mChangeDomain, $mPasswordGenerator, $mChangeDC, $mSearchAD, $mADHealth)),
  [Terminal.Gui.MenuBarItem]::new("_Selection", @($mSelectionMode, $mSelectAll, $mDeselectAll, $mBulkAddGroup)),
  [Terminal.Gui.MenuBarItem]::new("_About", @($mShortcuts, $mAboutDSATUI, $mWhyBlaabaer))
))

## Apply full theme to all components
try {
    Apply-Theme -ThemeData $themeData -TopLevel $top -MainWindow $win -Menu $menu -Status $StatusBar
    Debug-Log ("Theme applied successfully: $Theme") -Type "info"
} catch {
    Debug-Log ("WARNING: Theme application failed: $($_.Exception.Message)") -Type "warn"
}

if ($null -eq $menu) {
    Debug-Log ("ERROR: Menu is null before adding to top!") -Type "error"
} else {
    $top.Add($menu)
    Debug-Log ("Menu added to top successfully") -Type "info"
}

## Add treeview to main window
if ($null -eq $Global:tree) {
    Debug-Log ("ERROR: tree is null!") -Type "error"
} else {
    $win.Add($Global:tree)
    Debug-Log ("Tree added to window successfully") -Type "info"
}

## ------------------------- Window Key Handler (for F-keys) ------------------------
$win.add_KeyPress({
    param($keyEvent)
    
    $key = $keyEvent.KeyEvent.Key
    
    switch ($key) {
        ([Terminal.Gui.Key]::F1) { Show-Modal "Shortcuts" "F1 - Help`nF2 - Password Generator`nF3 - New`nF5 - Refresh`n`nF6 - Themes`nF7 - Search`nF7 - TODO`nF9 - Menus`nF10 - Quit`nF11 Full-Screen`nF12 Redraw" ; $keyEvent.Handled = $true }
        ([Terminal.Gui.Key]::F2) { Generate-RandomPassword ; $keyEvent.Handled = $true }
        ([Terminal.Gui.Key]::F3) { Show-NewObjectWizard ; $keyEvent.Handled = $true }
        ([Terminal.Gui.Key]::F5) { Refresh-Data -domain $Global:CurrentDomain ; $keyEvent.Handled = $true }
        ([Terminal.Gui.Key]::F6) { Show-ThemeSelector ; $keyEvent.Handled = $true }
        ([Terminal.Gui.Key]::F7) { Show-ADSearchDialog ; $keyEvent.Handled = $true }
        ([Terminal.Gui.Key]::F8) { Show-Modal "Coming soon!" ; $keyEvent.Handled = $true }
        ## Can't use F9 that shows drop down menus
        ([Terminal.Gui.Key]::F10) { [Terminal.Gui.Application]::RequestStop() ; $keyEvent.Handled = $true }
        ([Terminal.Gui.Key]::F11) { $keyEvent.Handled = $true }
        ([Terminal.Gui.Key]::F12) { [Terminal.Gui.Application]::Refresh() ; $keyEvent.Handled = $true }
    }
})

## ------------------------- Run application ------------------------
Debug-Log ("DEBUG: Top is null? $($null -eq $top)") -Type "info"
Debug-Log ("DEBUG: Win is null? $($null -eq $win)") -Type "info"
Debug-Log ("DEBUG: Menu is null? $($null -eq $menu)") -Type "info"
Debug-Log ("DEBUG: StatusBar is null? $($null -eq $StatusBar)") -Type "info"
Debug-Log ("DEBUG: Tree is null? $($null -eq $Global:tree)") -Type "info"
Debug-Log ("DEBUG: CurrentDomain = $Global:CurrentDomain") -Type "info"
Debug-Log ("DEBUG: Users count = $($Global:Users.Count)") -Type "info"
Debug-Log ("DEBUG: ADObjects count = $($Global:ADObjects.Count)") -Type "info"

[Terminal.Gui.Application]::Run($top)
[Terminal.Gui.Application]::Shutdown()

## Cleanup
if ($Global:LogStream) {
  try {
    $Global:LogStream.Flush()
    $Global:LogStream.Close()
    $Global:LogStream.Dispose()
  } catch {
    # Ignore cleanup errors
  }
}
