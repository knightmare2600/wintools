<#

DSA-TUI Text Mode version of dsa.msc for powershell
Locked-in baseline: dynamic resize, menu, demo data mirrors prod format, Change Domain fixed, fixed DC selection, full production AD object detection, properties modal, AD search popup

Information one may find useful

https://nyheder.tv2.dk/lokalt/2021-10-19-er-det-her-postbuddets-vaerste-skraek-hvem-pokker-har-fundet-paa-det-her

Citations for the Active directory commands:

  Reference URLs for AD Forest & Get-ADForest usage:

  - Microsoft Official Documentation
    https://learn.microsoft.com/powershell/module/activedirectory/get-adforest
    https://learn.microsoft.com/windows-server/identity/ad-ds/plan/active-directory-forest-and-domain-guide

  - General AD Forest Structure Explanations
    https://learn.microsoft.com/windows-server/identity/ad-ds/plan/understanding-active-directory-domain-services
    https://learn.microsoft.com/windows-server/identity/ad-ds/plan/designing-the-logical-structure

  - Community / Confirmatory Sources
    https://ss64.com/ps/get-adforest.html
    https://4sysops.com/archives/get-adforest-cmdlet-what-it-does-and-how-to-use-it/
    https://social.technet.microsoft.com/wiki/contents/articles/22953.active-directory-forest.aspx

Notes:
  - (Get-ADForest).RootDomain     Example Output: example.com
  - (Get-ADForest).Domains        Example Output: example.com
                                                  example.net

  - (Get-ADForest).GlobalCatalogs Example Output: gc-de01.example.com
                                                  gc-dk01.example.com
                                                  gc-uk01.example.com
                                                  gc-de01.example.net
                                                  gc-dk01.example.net
                                                  gc-uk01.example.net

  - (Get-ADForest).$Script:Sites  Example Output: DE # Germany
                                                  DK # Denmark
                                                  UK # United Kingdom

$forest = Get-ADForest

Write-Host "FOREST:" $forest.Name
Write-Host "ROOT DOMAIN:" $forest.RootDomain

Write-Host "ALL DOMAINS IN FOREST:"
$forest.Domains | ForEach-Object { " - $_" }

## Group domains by trees (DNS namespace)
Write-Host "DOMAIN TREES:
$forest.Domains | Group-Object { ($_ -split '\.')[1..99] -join '.' } | ForEach-Object { Write-Host "Tree: $($_.Name)" $_.Group | ForEach-Object { Write-Host "   - $_" } }

Write-Host "TRUSTS:"
Get-ADTrust -Filter * | Format-Table Name, Source, Target, TrustType, TrustDirection

A note on phone numbers:

  - UK Phone Number Standards (Ofcom reserved ranges for fiction/testing):
    - Glasgow: 0141 496 0xxx + Mobile: 07700 900xxx
    - Edinburgh: 0131 496 0xxx + Mobile: 07700 900xxx
    - London: 020 7946 0xxx / 01632 96xxxx + Mobile: 07700 900xxx

  - Denmark Testing Numbers:
    - Copenhagen landline: +45 0000-xxxx
    - Denmark mobile: +45 2xxx xxxx

NB: I don't believe Denmark uses 0000 but this is not confirmed!

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
 - Fixed Password Generator dialog always showing success regardless of click.
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
  - Bring forward $Script code for a more streamlined approach

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

2.1.5.2  (Computers, refresh debug, new bands, statusbar++)
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

2.3.7.0 - (Logic, Order of operations, and Plumbing fixed)
  - Tab views are now so complex we'll use functions as 1.16 does not support multiple tab lines in
    window. This is scaffolding for implementing all (!!) dsa.msc properties for produciton.
  - Clean up and standardis Debug-Log calls. Add Debug with a spanner and STOP SHOUTING in messages
  - TWA Airlines Theme
  - Move menus to a fnction
  - Rebuild tabs to have two tab rows. This is not technically possible os call ti "faux" tab rows
  - Refactor and introduce Steps (phases) of code so the order of operations is natural
  - Add Debug-DumpViewTree to debug the application at runtime
  - Clean-up redundant code and merge fitlering boxes into two panes
  - Change domain text box reworked and inut box widened
  - Refresh Domian data no longer kills the script, and status bar messages advise on progress
  - Show-UserProperties renders demo and real world information properly
  - Show-LAPSSearchModal modal to look up LAPS passwords
  - Fixed the regression where themes apply just fine from Selector Modal but not from command line.
  - Fix both the tree and the top left modal to not show different background colours.
  - Improve the User and Computer properties with tabs
  - Changing domain updates syatus bar during each step, and refreshes the tree
  - Statusbar hotkeys don't work but code is there
  - Password entrpoy strength follows colour scheme now, rather than always being green
  - Use one statusbar inside a function, so updates and UX feedback to the user is consistent. also helps
    with feedback during startup
  - Allow copying, pasting and exporting of output of LDAP queries
  - Move add and remove groups into functions and rework it to be aware if called from group or user properties tab

2.3.7.1
  - Initialize-DirectoryEmoji based on date. Special days use different emojis, e.g. Dec 25th

TODO:
 - misisng AD module is non fatal BUT if it's not installed, a global needs to not let users do stupid stuff
 - the terminal icons module, likewise if it's there great, if not fall back
 - show locked users actually shows computers under maintenence, which is... nice but not what you're asking for - add both
 - Did the right click popup go away or is it broken...?

 BUGS:
  - Users don't get shown in real AD, one suspects it's too hardcoded on the heiraracy so needs ot actually detect it [X] Think this is fixed but confirm
  - also it gives appearance of "hanging" because it doesn't exist before trying to switch. Needs to acknowledge it, repaint status bar and hten show loading enumraintg stuff [x] possibly fixed too

===========================================================================================
#>

param(
  [switch]$DemoMode,
  [switch]$Logging,
  [string]$LogFile,
  [string]$Domain,  ## User can specify domain
  [ValidateSet("light","dark","matrix","british", "panam", "dsb", "gemstones", "class91", "scotrail","twa" )]
  [string]$Theme
)

## Define the build version, project and code names once only - up here to ease patching. The rest at main
$Script:ProjectName = "DSA-TUI pwsh dsa.msc TUI"
$Script:FruitName = "Blåbær"
$Script:BuildVersion = "2.3.7.1"

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

##################################################################################################################
## Any functions added in here, make sure to keep chronology when calling them from inside other functions...   ##
##################################################################################################################

## -------------------{ Test For Required Modules }--------------------
function Test-RequiredModule {
  param(
    [Parameter(Mandatory)]
    [string]$Name,
    [string]$MinimumVersion = $null,
    [switch]$Optional
  )

  $module = Get-Module -ListAvailable -Name $Name -ErrorAction SilentlyContinue

  if ($module) {
    ## Check version if specified
    if ($MinimumVersion -and ($module.Version -lt [version]$MinimumVersion)) {
      Debug-Log "Module '$Name' found but version is too old. Need $MinimumVersion or later." -Type "Warn"
      return $false
    }

    Debug-Log "Module '$Name' is installed." -Type "Success"

    ## Try to import it
    try {
      Import-Module $Name -ErrorAction SilentlyContinue
    } catch {
      Debug-Log "Failed to import module '$Name': $_" -Type "Warn"
    }

    return $true
  }
  else {
    if ($Optional) {
      Debug-Log "Optional module '$Name' is NOT installed." -Type "Warn"
    } else {
      Debug-Log "Module '$Name' is NOT installed. Please run: Install-Module -Name $Name" -Type "Error"
    }
    return $false
  }
}

function Debug-DumpViewTree {
  param(
    [Terminal.Gui.View]$View,
    [int]$Indent = 0
  )

  if (-not $View) { return }

  $pad = (' ' * ($Indent * 2))
  $type = $View.GetType().Name

  $x = if ($View.X) { $View.X.ToString() } else { "null" }
  $y = if ($View.Y) { $View.Y.ToString() } else { "null" }
  $w = if ($View.Width) { $View.Width.ToString() } else { "null" }
  $h = if ($View.Height) { $View.Height.ToString() } else { "null" }

  $vis = $View.Visible
  $scheme = if ($View.ColorScheme) { "HasTheme" } else { "DefaultTheme" }

  Debug-Log ("$pad$type | X=$x Y=$y W=$w H=$h Visible=$vis Theme=$scheme") -Type "Debug"

  if ($View.Subviews -and $View.Subviews.Count -gt 0) {
    foreach ($child in $View.Subviews) {
      Debug-DumpViewTree -View $child -Indent ($Indent + 1)
    }
  }
}

## --------------------------{ Debug Logging }-------------------------
function Debug-Log {
  param(
    [string]$Message,
    [ValidateSet('Info','Warn','Error','Success','Debug')]
    [string]$Type = 'Info'
  )

  $timestamp = Get-Date -Format 'HH:mm:ss'
  ## Emoji + colors for each type
  switch ($Type) {
    'Info'    { $emoji = 'ⓘ' ; $color = 'Cyan' }   # Circled i
    'Warn'    { $emoji = '▲' ; $color = 'Yellow' }  # Triangle
    'Error'   { $emoji = '✗' ; $color = 'Red' }    # X mark
    'Success' { $emoji = '✓' ; $color = 'Green' }  # Check mark
    'Debug'   { $emoji = '🔧'; $color = 'Cyan'}  # Spanner (Wrench for left pondians)
  }

  $line = "[$timestamp] $emoji $Type $Message"

  ## Show in console when Logging switch is enabled
  if ($Script:Logging) {
    ## Use PSWriteColor if available, otherwise fall back to Write-Host
    if ($Script:HasPSWriteColor) {
      try {
        Write-Color -Encoding UTF8 -Text $line -Color $color
      } catch {
        Write-Host $line -ForegroundColor $color  # Fallback
      }
    } else {
      Write-Host $line -ForegroundColor $color
    }
  }

  ## Write to log file if enabled
  if ($Script:Logging -and $Script:LogStream) {
    try {
      $Script:LogStream.WriteLine($line)
    } catch { }
  }
}

## ---------------------------{ Set Up Tree }--------------------------
function Initialize-Tree {
  if (-not $Script:tree) { $Script:tree = [Terminal.Gui.TreeView]::new() }
}

function Build-MainMenu {
  [CmdletBinding()]
  param()

  Debug-Log ": Building main menu..." -Type "Info"

  ## ------------------------- Menu Items -------------------------
  $mFile  = [Terminal.Gui.MenuItem]::new("_Exit","Exit application (F10)",[Action]{ [Terminal.Gui.Application]::RequestStop() })
  $mNew   = [Terminal.Gui.MenuItem]::new("New Object","Create a new object (F3)",[Action]{ Show-NewObjectWizard })
  $mProps = [Terminal.Gui.MenuItem]::new("_Properties","Edit selected properties",[Action]{ Show-Properties })

  $mUndo         = [Terminal.Gui.MenuItem]::new("_Undo","Undo last action",[Action]{ Debug-Log (": Undo placeholder") -Type "Info" })
  $mChangeDomain = [Terminal.Gui.MenuItem]::new("Change _Domain","Select domain",[Action]{ Show-ChangeDomainDialog })
  $mChangeDC     = [Terminal.Gui.MenuItem]::new("Change _Domain Controller","Select DC",[Action]{ Show-ChangeDCDialog })
  $mSearchAD     = [Terminal.Gui.MenuItem]::new("_Search AD","Search Active Directory (F7)",[Action]{ Show-ADSearchDialog })

  $mRefresh = [Terminal.Gui.MenuItem]::new("_Refresh","Refresh AD data (F5)",[Action]{
    Debug-Log (": Refresh menu clicked - scheduling refresh...") -Type "Info"
    [Terminal.Gui.Application]::MainLoop.AddTimeout([TimeSpan]::FromMilliseconds(100), {
    Debug-Log (": Timeout callback - starting refresh...") -Type "Info"
    try {
      Update-Status "Refreshing..." -spinner
      $result = Refresh-Data -domain $Script:CurrentDomain -RebuildTree
      if ($result) { Update-Status "Refresh complete" -final } else { Update-Status "Refresh failed" -final }
        Debug-Log (": Refresh completed with result: $result") -Type "Info"
      } catch {
        Debug-Log (": Refresh crashed: $($_.Exception.Message)") -Type "Info"
        Update-Status "Refresh error" -final
      }
        return $false
    })
      Debug-Log (": Refresh scheduled") -Type "Info"
  })

  $mQuickFilter       = [Terminal.Gui.MenuItem]::new("_Quick Filter","Apply quick filters",[Action]{ Show-QuickFilterDialog })
  $mSelectionMode      = [Terminal.Gui.MenuItem]::new("_Selection Mode (Ctrl+S)","Toggle selection mode",[Action]{ Toggle-SelectionMode })
  $mSelectAll         = [Terminal.Gui.MenuItem]::new("Select _All (Ctrl+A)","Select all objects",[Action]{ Select-AllObjects })
  $mDeselectAll       = [Terminal.Gui.MenuItem]::new("_Deselect All (Ctrl+D)","Deselect all objects",[Action]{ Deselect-AllObjects })
  $mBulkAddGroup      = [Terminal.Gui.MenuItem]::new("Add to _Group...","Add selected users to group",[Action]{ Invoke-BulkAddToGroup })
  $mPasswordGenerator = [Terminal.Gui.MenuItem]::new("_Password Generator","Password Generator (F2)",[Action]{ Generate-RandomPassword })
  $mLAPSPasswords     = [Terminal.Gui.MenuItem]::new("_LAPS Passwords","Lookup LAPS Creds (Fx)",[Action]{ Show-LAPSSearchModal })
  $mADHealth          = [Terminal.Gui.MenuItem]::new("_AD Health Status","AD Health And Replication Status",[Action]{ Get-ADHealth })
  $mShortcuts         = [Terminal.Gui.MenuItem]::new("_Shortcuts","Keyboard shortcuts (F1)",[Action]{ Show-Modal "Shortcuts" "F1 - Help`nF2 - Password Generator`nF3 - New`nF5 - Refresh`nF6 - Themes`nF7 - Search`nF10 - Quit" })
  $mAboutDSATUI       = [Terminal.Gui.MenuItem]::new("_About","About $($Script:ProjectName)",[Action]{ Show-Modal "About" "$($Script:ProjectName)`n`nCodename: $($Script:FruitName)`nv$($Script:BuildVersion) STABLE`nGPL-3 Copyleft`nBy Knightmare2600 (https://github.com/knightmare2600)" })
  $mWhyBlaabaer       = [Terminal.Gui.MenuItem]::new("Why _Blaabaer?","Why the $($Script:FruitName) codename?",[Action]{ Show-BlaabaerInfo })
  $mTheme             = [Terminal.Gui.MenuItem]::new("_Theme","Change color theme (F6)",[Action]{ Show-ThemeSelector })

  ## ------------------------- Menu Bar -------------------------
  $menu = [Terminal.Gui.MenuBar]::new(@(
    [Terminal.Gui.MenuBarItem]::new("_File", @($mRefresh, $mTheme, $mFile)),
    [Terminal.Gui.MenuBarItem]::new("_Action", @($mNew, $mProps, $mQuickFilter, $mUndo, $mChangeDomain, $mLAPSPasswords, $mPasswordGenerator, $mChangeDC, $mSearchAD, $mADHealth)),
    [Terminal.Gui.MenuBarItem]::new("_Selection", @($mSelectionMode, $mSelectAll, $mDeselectAll, $mBulkAddGroup)),
    [Terminal.Gui.MenuBarItem]::new("_About", @($mShortcuts, $mAboutDSATUI, $mWhyBlaabaer))
  ))

  Debug-Log ": Main menu created successfully" -Type "Info"
  return $menu
}

## ----------------------------{ Get Theme }---------------------------
function Get-Theme {
  param([string]$mode)

  ## Initialize color schemes and Ensure ColorSchemes are instantiated
  if (-not $ScriptCs)     { $ScriptCs     = [Terminal.Gui.ColorScheme]::new() }
  if (-not $mainWindowCs) { $mainWindowCs = [Terminal.Gui.ColorScheme]::new() }

  ## Normalize theme string: lowercase + ASCII
  $mode = $mode.Trim().ToLower()

  ####    $mode = $mode -replace "ae","ae"

  <# Leave me alone for I am documentation

  Adding Themes:

  Add an option above in the [ValidateSet]() then define a theme below:

  "faxekondi" {
    $ScriptCs.Normal     <-- Foreground borders and background colour for all modals
    $ScriptCs.Focus      <-- Foreground and background for menus
    $mainWindowCs.Normal <-- Main opening dialog and foreground text colour
    $mainWindowCs.Focus  <-- Main opening window focus colours foreground nad background
    }

  Valid colours: Black, Blue, Green, Cyan, Red, Magenta, Brown, Gray, DarkGray, BrightBlue,
                 BrightGreen, BrightCyan, BrightRed, BrightMagenta, BrightYellow, White

  Also leave me alone for I am also documentaiton #>

  switch ($mode) {
    "light" {
      $ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Cyan)
      $ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Cyan)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::Blue)
    }

    "dark" {
      $ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Black)
      $ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Green,[Terminal.Gui.Color]::White)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Green)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Red)
    }

    "matrix" {
      $ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Green,[Terminal.Gui.Color]::Black)
      $ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Green)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::Black)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::Gray)
    }

    "british" {
      $ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
      $ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::White)
    }

    "panam" {
      $ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightBlue,[Terminal.Gui.Color]::White)
      $ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightBlue)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightBlue)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightBlue,[Terminal.Gui.Color]::White)
    }

    "dsb" {
      $ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
      $ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::White)
    }

    "gemstones" {
      $ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Green,[Terminal.Gui.Color]::White)
      $ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::BrightMagenta)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightGreen)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Green,[Terminal.Gui.Color]::White)
    }

    "scotrail" {
      $ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::Blue)
      $ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Blue,[Terminal.Gui.Color]::White)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::BrightBlue)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
    }

    "class91" {
      $ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Black)
      $ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::White)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightRed,[Terminal.Gui.Color]::White)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightBlue)
    }

    "twa" {
      $ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
      $ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::DarkGray)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightRed,[Terminal.Gui.Color]::DarkGray)
    }

    "default" {
      # fallback to dark
      $ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Black)
      $ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Gray)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Black)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::DarkGray)
    }
  }

  ## Ensure HotNormal/HotFocus
  $ScriptCs.HotNormal     = $ScriptCs.Normal
  $ScriptCs.HotFocus      = $ScriptCs.Focus
  $mainWindowCs.HotNormal = $mainWindowCs.Normal
  $mainWindowCs.HotFocus  = $mainWindowCs.Focus

  return @{
    Global     = $ScriptCs
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
  ## --- Tree
  if ($Script:tree -and $Script:tree.PSObject.Properties.Name -contains 'ColorScheme') { $Script:tree.ColorScheme = $ThemeData.Global }
  ## --- Filter Panel
  if ($filterPanel -and $filterPanel.PSObject.Properties.Name -contains 'ColorScheme') { $filterPanel.ColorScheme = $ThemeData.Global }

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
  if ($null -eq $Scheme) { Debug-Log ("ColorScheme is null") -Type "Warn" ; return }
  Debug-Log ("Normal    : $($Scheme.Normal)") -Type "Info"
  Debug-Log ("Focus     : $($Scheme.Focus)") -Type "Info"
  Debug-Log ("HotNormal : $($Scheme.HotNormal)") -Type "Info"
  Debug-Log ("HotFocus  : $($Scheme.HotFocus)") -Type "Info"
  Debug-Log ("Disabled  : $($Scheme.Disabled)") -Type "Info"
}

## ------------------------{ Show progress bar }-----------------------
## Helper: Show a simple loading/progress dialog with spinner
function Show-LoadingDialog {
  param(
    [string]$Message = "Loading, please wait..."
  )

  ## --- Create the dialog ---
  $dlg = [Terminal.Gui.Dialog]::new("", 40, 7)
  $dlg.X = 2
  $dlg.Y = 2

  ## --- Message label ---
  $lbl = [Terminal.Gui.Label]::new($Message)
  $lbl.X = 2
  $lbl.Y = 2
  $dlg.Add($lbl)

  ## --- Spinner label ---
  $spinner = [Terminal.Gui.Label]::new("|")
  $spinner.X = [Terminal.Gui.Pos]::Right($lbl) + 1
  $spinner.Y = 2
  $dlg.Add($spinner)

  ## --- Spinner globals ---
  $Script:spinnerFrames = @("|","/","-","\\")
  $Script:spinnerFrameIndex = 0

  ## --- Timer callback: update spinner safely on UI thread ---
  $timer = [System.Threading.Timer]::new({
    [Terminal.Gui.Application]::MainLoop.Invoke({
      $Script:spinnerFrameIndex = ($Script:spinnerFrameIndex + 1) % $Script:spinnerFrames.Count
      $spinner.Text = $Script:spinnerFrames[$Script:spinnerFrameIndex]

      ## Optional demo mode slowdown
      if ($Script:DemoMode) { Start-Sleep -Milliseconds 350 }
    })
  }, $null, 0, 150)  # 150ms interval

  ## --- Show dialog in modal-safe, non-blocking way ---
  [Terminal.Gui.Application]::BeginModal($dlg)

  ## --- Return dialog and timer for clean-up ---
  return [PSCustomObject]@{
  Dialog = $dlg
  Timer  = $timer
  }
}

## ------------------------{ Close Progressbar }-----------------------
function Show-ThemeSelector {

  ## --- Theme list ---
  $themes = @("british","dark","dsb","light","matrix","panam","gemstones","class91","scotrail", "twa")

  ## Split into two columns
  $half = [math]::Ceiling($themes.Count / 2)
  $leftThemes  = $themes[0..($half-1)]
  $rightThemes = $themes[$half..($themes.Count-1)]

  ## --- Determine current theme (case-insensitive) ---
  $currentTheme = $Script:ThemeMode
  Debug-Log (": Global ThemeMode = ${Global:ThemeMode}") -Type "Debug"
  Debug-Log (": Current theme for selection = ${currentTheme}") -Type "Debug"

  $currentIndex = -1
  for ($i = 0; $i -lt $themes.Count; $i++) {
    if ($themes[$i].ToLower() -eq $currentTheme.ToLower()) {
      $currentIndex = $i
      break
    }
  }

  Debug-Log (": Index of current theme in $themes array = ${currentIndex}") -Type "Debug"

  ## Calculate which column gets the selection
  $leftSelected  = if ($currentIndex -ge 0 -and $currentIndex -lt $leftThemes.Count) { $currentIndex } else { -1 }
  $rightSelected = if ($currentIndex -ge $leftThemes.Count) { $currentIndex - $leftThemes.Count } else { -1 }

  Debug-Log (": LeftSelected = ${leftSelected}, RightSelected = ${rightSelected}") -Type "Debug"

  ## --- Create dialog ---
  $dlg = [Terminal.Gui.Dialog]::new("Select Theme", 60, 16)

  $lbl = [Terminal.Gui.Label]::new("Choose a color theme:")
  $lbl.X = 2; $lbl.Y = 1
  $dlg.Add($lbl)

  ## --- Left column RadioGroup ---
  $rdoLeft = [Terminal.Gui.RadioGroup]::new($leftThemes)
  $rdoLeft.X = 2
  $rdoLeft.Y = 3
  $rdoLeft.SelectedItem = $leftSelected
  $dlg.Add($rdoLeft)

  ## --- Right column RadioGroup ---
  $rdoRight = [Terminal.Gui.RadioGroup]::new($rightThemes)
  $rdoRight.X = 32
  $rdoRight.Y = 3
  $rdoRight.SelectedItem = $rightSelected
  $dlg.Add($rdoRight)

  Debug-Log (": rdoLeft.SelectedItem = ${rdoLeft.SelectedItem}, rdoRight.SelectedItem = ${rdoRight.SelectedItem}") -Type "Debug"

  ## --- Sync columns so only one can be selected ---
  $rdoLeft.add_SelectedItemChanged({
    if ($rdoLeft.SelectedItem -ge 0) { $rdoRight.SelectedItem = -1 }
  })
  $rdoRight.add_SelectedItemChanged({
    if ($rdoRight.SelectedItem -ge 0) { $rdoLeft.SelectedItem = -1 }
  })

  ## --- Apply Button ---
  $btnApply = [Terminal.Gui.Button]::new("Apply")
  $btnApply.add_Clicked({
    $sel = if ($rdoLeft.SelectedItem -ge 0) {
      $leftThemes[$rdoLeft.SelectedItem]
    } elseif ($rdoRight.SelectedItem -ge 0) {
      $rightThemes[$rdoRight.SelectedItem]
    } else {
      "dark"
    }

    Debug-Log ": Theme selected on Apply = ${sel}" -Type "Debug"
    Debug-Log "Switching to theme: ${sel}" -Type "Info"

    $Script:ThemeMode = $sel
    $newTheme = Get-Theme -mode $sel  # Get NEW theme

    ## Apply to ALL components including tree frame
    Apply-Theme -ThemeData $newTheme -TopLevel [Terminal.Gui.Application]::Top -MainWindow $win -Menu $menu -Status $Script:StatusBar

    ## Apply to tree frame specifically
    if ($treeFrame -and $newTheme.MainWindow) {
      $treeFrame.ColorScheme = $newTheme.MainWindow
    }

    ## Apply to filter panel
    if ($filterPanel -and $newTheme.MainWindow) {
      $filterPanel.ColorScheme = $newTheme.MainWindow
    }

    ## Apply to selected objects panel
    if ($selectedObjectsPanel -and $newTheme.MainWindow) {
      $selectedObjectsPanel.ColorScheme = $newTheme.MainWindow
    }

    ## Force redraw everything
    [Terminal.Gui.Application]::Top.SetNeedsDisplay()
    [Terminal.Gui.Application]::Refresh()

    Show-Modal "Theme Changed" "Theme changed to: ${sel}"
    [Terminal.Gui.Application]::RequestStop()
  })
  $dlg.AddButton($btnApply)

  ## --- Cancel Button ---
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $dlg.AddButton($btnCancel)

  ## --- Run the dialog ---
  [Terminal.Gui.Application]::Run($dlg)
}

function Layout-Tabs {

  Debug-Log "Reflowing tabs across rows" -Type "Debug"

  if (-not $Script:AllTabs -or $Script:AllTabs.Count -eq 0) { return }

  ## Clear all rows
  foreach ($row in $Script:TabRows) {
    while ($row.Tabs.Count -gt 0) { $row.RemoveTab($row.Tabs[0]) }
  }

  for ($i = 0; $i -lt $Script:AllTabs.Count; $i++) {

    $rowIndex = [Math]::Floor($i / 6)
    if ($rowIndex -ge $Script:TabRows.Count) { break }

    $row = $Script:TabRows[$rowIndex]
    $originalTab = $Script:AllTabs[$i]

    ## Create a fresh tab
    $tab = [Terminal.Gui.TabView+Tab]::new()

    ## Convert ASCII numbers to readable text (works for log case)
    $asciiString = "$($originalTab.Text)"
    $tab.Text = ($asciiString -split '\s+' | ForEach-Object { [char][int]$_ }) -join ''

    ## ✅ Attach a fresh View to avoid Visible exception
    $tab.View = [Terminal.Gui.View]::new()
    Debug-Log ("Layout-Tabs: rowIndex=$rowIndex i=$i tabType=$($tab.GetType().FullName) tabText='$($tab.Text)' hasView=$([bool]$tab.View)") -Type "Debug"

    ## Add tab using 2-argument overload (required in 1.16)
    $row.AddTab($tab, $false)
  }

  ## Select first tab
  if ($Script:AllTabs.Count -gt 0) { Select-TabGlobal $Script:AllTabs[0] }
}

## ------------------------{ Faux Tab Builder }-----------------------
function New-FauxTabContainer {
  <#
  .SYNOPSIS
    Creates a multi-row tab container for Terminal.Gui dialogs.

  .DESCRIPTION
    This function sets up a faux multi-row tab container with keyboard navigation (Ctrl+Arrow) and a content host where individual tab Views can be displayed.
    It returns a custom object containing the container view, the content host, and a helper function to select tabs programmatically.

  .PARAMETER Tabs
    An array of [Terminal.Gui.TabView+Tab] objects to be placed in the tab rows.

  .PARAMETER Rows
    Number of visible tab rows. Defaults to 2.

  .PARAMETER Top
    Y-offset for the container within the parent dialog. Defaults to 0.

  .PARAMETER Height
    Height of the container. Defaults to Fill minus row count.

  .OUTPUTS
    PSCustomObject with:
      - Container -> outer view to add to dialog
      - ContentHost -> view for tab contents
      - SelectTab(tab) -> scriptblock to switch tabs
  #>

  param(
    [Parameter(Mandatory)]
    [array]$Tabs,

    [int]$Rows = 2,
    [int]$Top = 0,
    [int]$Height = 0
  )

  ## Outer container
  $container = [Terminal.Gui.View]::new()
  $container.X = 0
  $container.Y = $Top
  $container.Width  = [Terminal.Gui.Dim]::Fill()
  $container.Height = if ($Height -gt 0) { $Height } else { [Terminal.Gui.Dim]::Fill($Rows + 1) }

  ## Create tab rows
  $tabRows = @()
  for ($i=0; $i -lt $Rows; $i++) {
    $row = [Terminal.Gui.TabView]::new()
    $row.X = 0
    $row.Y = $i
    $row.Width  = [Terminal.Gui.Dim]::Fill()
    $row.Height = 1
    $tabRows += $row

    ## Wire up tab selection to global handler
    $row.add_SelectedTabChanged({
      param($sender, $args)
      if ($Script:AllTabs -and $args.NewTab) {
        $Script:ActiveTab = $args.NewTab
        if ($container.ContentHost) {
          $container.ContentHost.RemoveAll()
          $container.ContentHost.Add($args.NewTab.View)
        }
      }
    })
  }

  $Script:TabRows = $tabRows
  $Script:AllTabs = $Tabs
  $Script:ActiveTab = $Tabs[0]

  ## Keyboard navigation: Ctrl+Arrows
  $container.Add_KeyPress({
    param($s, $e)
    if (-not $Script:ActiveTab) { return }
      $idx = $Script:AllTabs.IndexOf($Script:ActiveTab)

      switch ($e.KeyEvent.Key) {
        ([Terminal.Gui.Key]::CtrlMask -bor [Terminal.Gui.Key]::CursorRight) {
          if ($idx -lt ($Script:AllTabs.Count - 1)) {
            $container.SelectTab.Invoke($Script:AllTabs[$idx + 1])
            $e.Handled = $true
          }
        }
        ([Terminal.Gui.Key]::CtrlMask -bor [Terminal.Gui.Key]::CursorLeft) {
          if ($idx -gt 0) {
            $container.SelectTab.Invoke($Script:AllTabs[$idx - 1])
            $e.Handled = $true
          }
        }
        ([Terminal.Gui.Key]::CtrlMask -bor [Terminal.Gui.Key]::CursorDown) {
          if ($idx + $Rows -lt $Script:AllTabs.Count) {
            $container.SelectTab.Invoke($Script:AllTabs[$idx + $Rows])
            $e.Handled = $true
          }
        }
        ([Terminal.Gui.Key]::CtrlMask -bor [Terminal.Gui.Key]::CursorUp) {
          if ($idx - $Rows -ge 0) {
            $container.SelectTab.Invoke($Script:AllTabs[$idx - $Rows])
            $e.Handled = $true
          }
        }
      }
  })

  ## Visual separator below tabs
  $separator = [Terminal.Gui.LineView]::new()
  $separator.X = 0
  $separator.Y = $Rows
  $separator.Width = [Terminal.Gui.Dim]::Fill()
  $separator.Height = 1

  ## Content host
  $contentHost = [Terminal.Gui.View]::new()
  $contentHost.X = 0
  $contentHost.Y = $Rows + 1
  $contentHost.Width  = [Terminal.Gui.Dim]::Fill()
  $contentHost.Height = [Terminal.Gui.Dim]::Fill()
  $container.ContentHost = $contentHost

  ## Add everything to container
  foreach ($row in $tabRows) { $container.Add($row) }
    $container.Add($separator)
    $container.Add($contentHost)

    ## Initial tab display
    if ($Tabs.Count -gt 0) {
      $container.ContentHost.Add($Tabs[0].View)
    }

    ## Function to select a tab programmatically
    $selectTab = {
      param($tab)
      $Script:ActiveTab = $tab
      $container.ContentHost.RemoveAll()
      $container.ContentHost.Add($tab.View)
      foreach ($row in $tabRows) {
        if ($row.Tabs.Contains($tab)) { $row.SelectedTab = $tab }
      }
    }
    $container.SelectTab = $selectTab

    ## Return container object
    return [PSCustomObject]@{
      Container   = $container
      ContentHost = $contentHost
      SelectTab   = $selectTab
      TabRows     = $tabRows
    }
}

############################### AD FUNCTIONS BELOW HERE ##################################
function Invoke-AD {
  param(
    [scriptblock]$Script,
    [switch]$SuppressError
  )
  try {
    if (-not (Get-Module -Name ActiveDirectory)) { Import-Module ActiveDirectory -ErrorAction Stop }
      return & $Script
    } catch {
    Debug-Log ("AD call failed: $($_.Exception.Message)") -Type "Warn"

    ## NEVER call Show-Modal here - it crashes Terminal.Gui! Just log and return null
    ## Fallback to demo mode
    $Script:DemoMode = $true
    return $null
  }
}

## ------------------------{ Load Domain Data }------------------------
function Get-ADObjectsByType {
  param([string]$domain)

  ## Debug info
  Debug-Log (": Get-ADObjectsByType DemoMode=${Script:DemoMode} Domain=${domain}") -Type "Debug"

  $objTypes = @("user","computer","group","organizationalUnit","contact")
  $allObjects = @()
  foreach ($type in $objTypes) {
    try {
      $objs = if ($Script:DemoMode) {
      ## Demo objects already structured
      @()
      } else {
        Get-ADObject -Filter "ObjectClass -eq '$type'" -Server $domain -Properties Name,ObjectClass,DistinguishedName |
        ForEach-Object { @{ Name=$_.Name; Type=$_.ObjectClass; DN=$_.DistinguishedName } }
      }
      $allObjects += $objs
      } catch {
      ## minimal fix: string interpolation of exception object done via ToString()
      Debug-Log (": Failed to enumerate ${type}: $($_.ToString())") -Type "Debug"
    }
  }
  return $allObjects
}

function Get-CleanObjectInfo {
  param([string]$treeText)  # ← Remove $Script:

  Debug-Log (": Get-CleanObjectInfo called with: '$treeText'") -Type "Info"  # ← Use $treeText

  ## Determine type FIRST (before cleaning)
  $objectType = if ($treeText -like "(U)*") { "user" }           # ← Use $treeText
                elseif ($treeText -like "(DC)*") { "dc" }        # ← Use $treeText
                elseif ($treeText -like "(G)*") { "group" }      # ← Use $treeText
                elseif ($treeText -like "(PC)*") { "computer" }  # ← Use $treeText
                else { "unknown" }

  Debug-Log (": Detected type: $objectType") -Type "Info"

  ## Remove prefixes like "(U) " or "(DC) " or "(G) "
  $cleanName = $treeText -replace '^\([^)]+\)\s*', ''  # ← Use $treeText
  Debug-Log (": After removing prefix: $cleanName") -Type "Info"

  ## Remove any non-letter, non-space characters from the start (status icons)
  $cleanName = $cleanName -replace '^[^a-zA-Z]+\s*', ''
  Debug-Log (": After removing icons: $cleanName") -Type "Info"

  ## Extract just the name if it has [SITE] suffix (for DCs)
  if ($cleanName -match '^(.+?)\s+\[.+\]$') {
    $cleanName = $matches[1].Trim()
    Debug-Log (": Extracted name from [SITE] format: $cleanName") -Type "Info"
  }

  Debug-Log (": Final cleaned name: $cleanName") -Type "Info"

  return @{
    Type = $objectType
    Name = $cleanName
  }
}

## -----------------------{ Refresh Domain Data }----------------------
function Refresh-Data {
  param([string]$domain, [switch]$RebuildTree)

  ## If in demo mode, just reload demo data
  if ($Script:DemoMode) {
    try {
      Update-Status "Refreshing demo data..." -spinner
      $converted = Convert-DataToADObjects -Users $Script:rawUsers -DCs $Script:rawDCs -Groups $Script:rawDemoGroups -Computers $Script:rawComputers -Domain $domain

      if ($RebuildTree) {
        Update-Status "Rebuilding tree..." -spinner

        ## Rebuild tree safely in the UI thread
        [Terminal.Gui.Application]::MainLoop.Invoke({
          try {
            $Script:tree.ClearObjects()
            Build-Tree -domain $domain
            [Terminal.Gui.Application]::Refresh()
          } catch {
            Debug-Log ("Tree rebuild error: $_") -Type "Error"
          }
        })
      }

      Update-Status "Demo data refreshed" -final
      return $true

    } catch {
      Debug-Log ("Demo refresh error: $_") -Type "Warn"
      Update-Status "Refresh error" -final
      return $false
    }
  }

  ## Production mode - query AD
  try {
    Update-Status "Loading DCs..." -spinner
    $dcs = Invoke-AD { Get-ADDomainController -Filter * -DomainName $domain -ErrorAction Stop } -SuppressError
    Update-Status "Loading Users..." -spinner

    $users = Invoke-AD { Get-ADUser -Filter * -Server $domain -Properties DisplayName,EmailAddress,Title,Department,Enabled,LockedOut,DistinguishedName -ErrorAction Stop } -SuppressError
    Update-Status "Loading Groups..." -spinner

    $groups = Invoke-AD { Get-ADGroup -Filter * -Server $domain -Properties Description,GroupCategory,GroupScope,Members,DistinguishedName -ErrorAction Stop } -SuppressError
    Update-Status "Loading Computers..." -spinner

    $computers = Invoke-AD { Get-ADComputer -Filter * -Server $domain -Properties OperatingSystem,OperatingSystemVersion,Enabled,LastLogonDate,DistinguishedName -ErrorAction Stop } -SuppressError

    ## Check results
    if ($null -eq $dcs -or $null -eq $users -or $null -eq $groups -or $null -eq $computers) {
      Debug-Log ("Refresh failed - one or more queries returned null") -Type "Warn"
      Update-Status "Refresh failed - check logs" -final
      return $false
    }

    Update-Status "Converting data..." -spinner
    $converted = Convert-DataToADObjects -Users $users -DCs $dcs -Groups $groups -Computers $computers -Domain $domain

    if ($RebuildTree) {
      Update-Status "Rebuilding tree..." -spinner

      ## Rebuild tree safely in the UI thread
      [Terminal.Gui.Application]::MainLoop.Invoke({
      try {
        $Script:tree.ClearObjects()
        Build-Tree -domain $domain
        [Terminal.Gui.Application]::Refresh()
      } catch {
        Debug-Log ("Tree rebuild error: $_") -Type "Warn"
      }
    })
  }

  Update-Status "Refresh complete" -final
  return $true

  } catch {
    Debug-Log ("Refresh error: $_") -Type "Warn"
    Update-Status "Refresh error" -final
    return $false
  }
}

## -----------------------{ Apply Group Changes }----------------------
function Apply-GroupChanges {
  param($group, $fields)

  try {
    if ($Script:DemoMode) {
      ## Apply changes to demo object
      $group.Description = $fields.txtDesc.Text.ToString()
      $group.Mail        = $fields.txtEmail.Text.ToString()
      $group.ManagedBy   = $fields.txtManagedBy.Text.ToString()

      ## Update raw demo group
      $rawGroup = $Script:rawDemoGroups | Where-Object { $_.Name -eq $group.Name } | Select-Object -First 1

      if ($rawGroup) {
        $rawGroup.Description = $fields.txtDesc.Text.ToString()
        $rawGroup.Email       = $fields.txtEmail.Text.ToString()
        $rawGroup.ManagedBy   = $fields.txtManagedBy.Text.ToString()
      }

      Debug-Log ("SUCCESS: Group changes applied (demo mode)") -Type "Success"
      Show-Modal "Success" "Changes applied successfully (demo mode)"
      Refresh-Data -domain $Script:CurrentDomain
    }
    else {
      ## AD mode
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
      Debug-Log ("SUCCESS: Group changes applied to AD") -Type "Success"
      Show-Modal "Success" "Changes applied successfully"
      Refresh-Data -domain $Script:CurrentDomain
    }

    $Script:groupChangesMade = $false
    return $true
  }
  catch {
    Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)"
    return $false
  }
}

## Is it a special day
function Initialize-DirectoryEmoji {
    <#
        Determines which emoji to use for the Active Directory window title.
        Uses Unicode escapes for flags to avoid editor issues.
    #>

    param(
        [DateTime]$Date = (Get-Date)
    )

    $month = $Date.Month
    $day   = $Date.Day

    # Default: card index
    $emoji = "🗂️"

    switch ($true) {
      { $month -eq 4 -and $day -eq 9 } { $emoji = "`u{1F1E9}`u{1F1F0}" ; break }      ## 9th Apr Danmarks besættelse (liberation day)
      { $month -eq 5 -and $day -eq 4 } { $emoji = "🕯️" ; break }                     ## 4th May Candle for Besættelsen
      { $month -eq 5 -and $day -eq 5 } { $emoji = "`u{1F1E9}`u{1F1F0}" ; break }      ## 5th May Constitution day in Denmark
      { $month -eq 6 -and $day -eq 21 } { $emoji = "`u{1F1EC}`u{1F1F1}" ; break }     ## 21st May Grønland Day
      { $month -eq 7 -and $day -eq 29 } { $emoji = "`u{1F1EB}`u{1F1F4}" ; break }     ## 29th Jul Faroe Islands
      { $month -eq 9 -and $day -eq 9 } { $emoji = "`u{1F1E9}`u{1F1EA}" ; break }      ## 9th Nov Erich Honecker leck mich am Arsch!
      { $month -eq 12 -and ($day -eq 24 -or $day -eq 25) } { $emoji = "🎄" ; break } ## 24th/25th Dec Tree for Jul / Christmas
    }

    Debug-Log ("Today's emoji is: $emoji") -Type "info"
    $Script:DirectoryEmoji = $emoji
}

## --------------------------{ Danske Soda vand }--------------------------
## This is a theme now. Danish Fruit soda based fun. Method to the madness
function Show-BlaabaerInfo {
  $dlg = [Terminal.Gui.Dialog]::new("Why $($Script:FruitName)? 🫐", 60, 12)

  $message = @"
$($Script:ProjectName) is codenamed $Script:FruitName because:

- I was drinking blueberry soda when writing the code
- $($Script:FruitName) is Danish for blueberry
- Føtex sells a rather nice $($Script:FruitName) soda
- Every great project needs a forest-fruit mascot!
"@

  $label = [Terminal.Gui.Label]::new(1, 1, $message)
  $dlg.Add($label)

  Show-Modal "Why $($Script:FruitName)...? 🫐" $message
}

## ------------------------{ Load Domain Data }------------------------
function Load-DomainData {
  param([string]$domain)
  if ($Logging) { Debug-Log (": Loading domain data for: $domain") -Type "Info" }

  ## If DemoMode already enabled, use demo data
  if ($Script:DemoMode) {
    Debug-Log ("Starting $($Script:ProjectName) in DEMO mode...") -Type "Debug"

    ## ------------------ Define Demo Users ------------------
    $Script:rawUsers = @(
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
  $Script:rawDemoGroups = @(
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
  ## Updated Demo Domain Controllers - Matching Production AD Structure
  ## Complete Demo Domain Controllers - Production Structure

$Script:rawDCs = @(
  ## ===== UK DCs (example.com) =====
  @{ Name = 'EXAGLADC01' ; HostName = 'EXAGLADC01.example.com' ; DNSHostName = 'EXAGLADC01.example.com' ; Site = 'GLA' ; Location = 'Glasgow, Scotland' ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OperatingSystem = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.4.20' ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @('Schema Master', 'Domain Naming Master', 'PDC Emulator') ;  OperationMasterRoles = @('Schema Master', 'Domain Naming Master', 'PDC Emulator') ; LastReplication = (Get-Date).AddMinutes(-12) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-45) ; LastBootUpTime = (Get-Date).AddDays(-45) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '45 GB'; Free = '75 GB'; PercentFree = 62} ; 'SYSVOL' = @{Total = '50 GB'; Used = '8 GB'; Free = '42 GB'; PercentFree = 84} } ; ReplicationPartners = @('EXALNDDC01', 'EXAEDIDC01') },
  @{ Name = 'EXAEDIDC01' ; HostName = 'EXAEDIDC01.example.com' ; DNSHostName = 'EXAEDIDC01.example.com' ; Site = 'EDI' ; Location = 'Edinburgh, Scotland' ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OperatingSystem = 'Windows Server 2019 Standard' ; OperatingSystemVersion = '10.0 (17763)' ; IPv4Address = '192.168.3.20' ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @() ; OperationMasterRoles = @() ; LastReplication = (Get-Date).AddMinutes(-18) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-67) ; LastBootUpTime = (Get-Date).AddDays(-67) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '38 GB'; Free = '82 GB'; PercentFree = 68} ; 'SYSVOL' = @{Total = '50 GB'; Used = '6 GB'; Free = '44 GB'; PercentFree = 88} } ; ReplicationPartners = @('EXAGLADC01', 'EXALNDDC01','EXANEWDC01', 'EXALIVDC01') },
  @{ Name = 'EXALNDDC01' ; HostName = 'EXALNDDC01.example.com' ; DNSHostName = 'EXALNDDC01.example.com' ; Site = 'LND' ; Location = 'London, England' ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OperatingSystem = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.2.20' ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @('RID Master', 'Infrastructure Master') ; OperationMasterRoles = @('RID Master', 'Infrastructure Master') ; LastReplication = (Get-Date).AddMinutes(-8) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-23) ; LastBootUpTime = (Get-Date).AddDays(-23) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '52 GB'; Free = '68 GB'; PercentFree = 57} ; 'SYSVOL' = @{Total = '50 GB'; Used = '12 GB'; Free = '38 GB'; PercentFree = 76} } ; ReplicationPartners = @('EXAGLADC01', 'EXAEDIDC01', 'EXAKGEDC01', 'EXACPHDC01') },
  @{ Name = 'EXANEWDC01' ; HostName = 'EXANEWDC01.example.com' ; DNSHostName = 'EXANEWDC01.example.com' ; Site = 'NEW' ; Location = 'Newcastle, UK' ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OperatingSystem = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.91.20' ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @() ; OperationMasterRoles = @() ; LastReplication = (Get-Date).AddMinutes(-13) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-42) ; LastBootUpTime = (Get-Date).AddDays(-42) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '24 GB'; Free = '96 GB'; PercentFree = 80} ; 'SYSVOL' = @{Total = '50 GB'; Used = '8 GB'; Free = '42 GB'; PercentFree = 84} } ; ReplicationPartners = @('EXALIVDC01', 'EXAEDIDC01', 'EXAGLADC01', 'EXALNDDC01') },
  @{ Name = 'EXALIVDC01' ; HostName = 'EXALIVDC01.example.com' ; DNSHostName = 'EXALIVDC01.example.com' ; Site = 'LIV' ; Location = 'Liverpool, UK' ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OperatingSystem = 'Windows Server 2025 Standard' ; OperatingSystemVersion = '10.0 (26100)' ; IPv4Address = '192.168.51.20' ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @() ; OperationMasterRoles = @() ; LastReplication = (Get-Date).AddMinutes(-15) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-12) ; LastBootUpTime = (Get-Date).AddDays(-12) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '46 GB'; Free = '74 GB'; PercentFree = 62} ; 'SYSVOL' = @{Total = '50 GB'; Used = '8 GB'; Free = '42 GB'; PercentFree = 84}  } ; ReplicationPartners = @('EXANEWDC01', 'EXAEDIDC01') },

  # ===== Denmark DCs (example.com) =====
  @{ Name = 'EXACPHDC01' ; HostName = 'EXACPHDC01.example.com' ; DNSHostName = 'EXACPHDC01.example.com' ; Site = 'CPH' ; Location = 'Copenhagen, Denmark' ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OperatingSystem = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.6.20' ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @() ; OperationMasterRoles = @() ; LastReplication = (Get-Date).AddMinutes(-10) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-34) ; LastBootUpTime = (Get-Date).AddDays(-34) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '41 GB'; Free = '79 GB'; PercentFree = 66} ; 'SYSVOL' = @{Total = '50 GB'; Used = '9 GB'; Free = '41 GB'; PercentFree = 82} } ; ReplicationPartners = @('EXALNDDC01', 'EXAKGEDC01', 'EXAODEDC01') },
  @{ Name = 'EXAKGEDC01' ; HostName = 'EXAKGEDC01.example.com' ; DNSHostName = 'EXAKGEDC01.example.com' ; Site = 'KGE' ; Location = 'Køge, Denmark' ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OperatingSystem = 'Windows Server 2016 Standard' ; OperatingSystemVersion = '10.0 (14393)' ; IPv4Address = '192.168.5.20' ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $false ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @() ; OperationMasterRoles = @() ; LastReplication = (Get-Date).AddHours(-4).AddMinutes(-35) ; ReplicationHealth = 'Warning - Out of Sync' ; LastBoot = (Get-Date).AddDays(-156) ; LastBootUpTime = (Get-Date).AddDays(-156) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '100 GB'; Used = '78 GB'; Free = '22 GB'; PercentFree = 22} ; 'SYSVOL' = @{Total = '40 GB'; Used = '28 GB'; Free = '12 GB'; PercentFree = 30} } ; ReplicationPartners = @('EXALNDDC01', 'EXACPHDC01') },
  @{ Name = 'EXAODEDC01' ; HostName = 'EXAODEDC01.example.com' ; DNSHostName = 'EXAODEDC01.example.com' ; Site = 'ODE' ; Location = 'Odense, Denmark' ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OperatingSystem = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.7.20' ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @() ; OperationMasterRoles = @() ; LastReplication = (Get-Date).AddMinutes(-14) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-28) ; LastBootUpTime = (Get-Date).AddDays(-28) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '42 GB'; Free = '78 GB'; PercentFree = 65} ; 'SYSVOL' = @{Total = '50 GB'; Used = '7 GB'; Free = '43 GB'; PercentFree = 86} } ; ReplicationPartners = @('EXACPHDC01', 'EXAKGEDC01') },

  # ===== Germany DCs (example.net) =====
  @{ Name = 'EXABONDC01' ; HostName = 'EXABONDC01.example.net' ; DNSHostName = 'EXABONDC01.example.net' ; Site = 'BON' ; Location = 'Bonn, Germany' ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OperatingSystem = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.22.20' ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @('Schema Master', 'Domain Naming Master') ; OperationMasterRoles = @('Schema Master', 'Domain Naming Master') ; LastReplication = (Get-Date).AddMinutes(-11) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-38) ; LastBootUpTime = (Get-Date).AddDays(-38) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '48 GB'; Free = '72 GB'; PercentFree = 60} ; 'SYSVOL' = @{Total = '50 GB'; Used = '10 GB'; Free = '40 GB'; PercentFree = 80} } ; ReplicationPartners = @('EXABRLDC01') },
  @{ Name = 'EXABRLDC01' ; HostName = 'EXABRLDC01.example.net' ; DNSHostName = 'EXABRLDC01.example.net' ; Site = 'BRL' ; Location = 'West Berlin, Germany' ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OperatingSystem = 'Windows Server 2019 Standard' ; OperatingSystemVersion = '10.0 (17763)' ; IPv4Address = '192.168.30.20' ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @('PDC Emulator', 'RID Master', 'Infrastructure Master') ; OperationMasterRoles = @('PDC Emulator', 'RID Master', 'Infrastructure Master') ; LastReplication = (Get-Date).AddMinutes(-9) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-51) ; LastBootUpTime = (Get-Date).AddDays(-51) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '55 GB'; Free = '65 GB'; PercentFree = 54} ;'SYSVOL' = @{Total = '50 GB'; Used = '14 GB'; Free = '36 GB'; PercentFree = 72} } ; ReplicationPartners = @('EXABONDC01', 'EXAMUCDC01') },
  @{ Name = 'EXAMUCDC01' ; HostName = 'EXAMUCDC01.example.net' ; DNSHostName = 'EXAMUCDC01.example.net' ; Site = 'MUC' ; Location = 'Munich, Germany' ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OperatingSystem = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.89.20' ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @() ; OperationMasterRoles = @() ; LastReplication = (Get-Date).AddMinutes(-13) ; ReplicationHealth = 'Healthy' ; LastBoot = (Get-Date).AddDays(-42) ; LastBootUpTime = (Get-Date).AddDays(-42) ; Services = @{DNS = 'Running'; DFSR = 'Running'; Netlogon = 'Running'; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '44 GB'; Free = '76 GB'; PercentFree = 63} ; 'SYSVOL' = @{Total = '50 GB'; Used = '8 GB'; Free = '42 GB'; PercentFree = 84} } ; ReplicationPartners = @('EXABONDC01', 'EXABRLDC01') }
)

Debug-Log ": Loaded $($Script:rawDCs.Count) demo domain controllers" -Type "Info"

# Note: Add remaining DCs (EXACPHDC01, EXAKGEDC01, etc.) following the same structure
# Each DC should have all the production properties listed above

## --------{ Demo Computers (Workstations, Laptops, Printers) }--------
$Script:rawComputers = @(
  @{ Name = 'GLA-WKS-001' ; Type = 'Workstation' ; Location = 'Glasgow, Scotland' ; OU = @('Locations','UK','Scotland','Glasgow','Computers') ; OS = 'Windows 11 Pro' ; OSVersion = '10.0.22631' ; LastLogon = (Get-Date).AddHours(-2) ; IPv4Address = '192.168.4.50' ; Description = 'Hot desk workstation - Glasgow office'; Enabled = $true },
  @{ Name = 'GLA-WKS-002' ; Type = 'Workstation' ; Location = 'Glasgow, Scotland' ; OU = @('Locations','UK','Scotland','Glasgow','Computers') ; OS = 'Windows 11 Pro' ; OSVersion = '10.0.22631' ; LastLogon = (Get-Date).AddDays(-1) ; IPv4Address = '192.168.4.51' ; Description = 'Hot desk workstation - Glasgow office' ; Enabled = $true },
  @{ Name = 'GLA-LAP-001' ; Type = 'Laptop' ; Location = 'Glasgow, Scotland' ; OU = @('Locations','UK','Scotland','Glasgow','Computers') ; OS = 'Windows 11 Pro' ; OSVersion = '10.0.22631' ; LastLogon = (Get-Date).AddHours(-5) ; IPv4Address = '192.168.4.80' ; Description = 'Dell Latitude laptop - Pool device' ; Enabled = $true },
  @{ Name = 'GLA-PRN-001' ; Type = 'Printer' ; Location = 'Glasgow, Scotland' ; OU = @('Locations','UK','Scotland','Glasgow','Computers') ; OS = 'Printer' ; OSVersion = 'N/A' ; LastLogon = (Get-Date).AddMinutes(-30) ; IPv4Address = '192.168.4.100' ; Description = 'HP LaserJet Pro - Main floor' ; Enabled = $true },
  @{ Name = 'LND-WKS-001' ; Type = 'Workstation' ; Location = 'London, England' ; OU = @('Locations','UK','England','London','Computers') ; OS = 'Windows 11 Pro' ; OSVersion = '10.0.22631' ; LastLogon = (Get-Date).AddHours(-1) ; IPv4Address = '192.168.2.50' ; Description = 'Hot desk workstation - London office' ; Enabled = $true },
  @{ Name = 'LND-PRN-001' ; Type = 'Printer' ; Location = 'London, England' ; OU = @('Locations','UK','England','London','Computers') ; OS = 'Printer' ; OSVersion = 'N/A' ; LastLogon = (Get-Date).AddMinutes(-15) ; IPv4Address = '192.168.2.100' ; Description = 'Xerox WorkCentre - Reception' ; Enabled = $true },
  @{ Name = 'BON-WKS-001' ; Type = 'Workstation' ; Location = 'Bonn, Germany' ; OU = @('Locations','Germany','Bonn','Computers') ; OS = 'Windows 11 Pro' ; OSVersion = '10.0.22631' ; LastLogon = (Get-Date).AddHours(-3) ; IPv4Address = '192.168.22.50' ; Description = 'Hot desk workstation - Bonn office' ; Enabled = $true },
  @{ Name = 'BON-LAP-001' ; Type = 'Laptop' ; Location = 'Bonn, Germany' ; OU = @('Locations','Germany','Bonn','Computers') ; OS = 'Windows 11 Pro' ; OSVersion = '10.0.22631' ; LastLogon = (Get-Date).AddDays(-2) ; IPv4Address = '192.168.22.75' ; Description = 'Lenovo ThinkPad - Pool device' ; Enabled = $false  } ## Disabled for maintenance
  @{ Name = 'LIV-MBP-001' ; Type = 'Macbook' ; Location = 'Liverpool, UK' ; OU = @('Locations','UK','Liverpool','Computers') ; OS = 'MacOS Tahoe' ; OSVersion = '26.0.1' ; LastLogon = (Get-Date).AddDays(-14) ; IPv4Address = '192.168.51.75' ; Description = 'Macbook Pro 2024 - Pool device' ; Enabled = $true  }
  @{ Name = 'LIV-MAC-001' ; Type = 'iMac' ; Location = 'Newcastle, UK' ; OU = @('Locations','UK','Newcastle','Computers') ; OS = 'MacOS Tahoe' ; OSVersion = '26.0.1' ; LastLogon = (Get-Date).AddDays(-40) ; IPv4Address = '192.168.91.75' ; Description = 'Macbook Pro 2024 - Pool device' ; Enabled = $false  } ## Disabled for maintenance
)

  Debug-Log ": Loaded $($Script:rawComputers.Count) demo computers and printers" -Type "Info"

  ## Load demo data
  $converted = Convert-DataToADObjects -Users $Script:rawUsers -DCs $Script:rawDCs -Groups $Script:rawDemoGroups -Computers $Script:rawComputers -Domain $Script:CurrentDomain
  return $converted
}

## Production - get real AD data
try {
  Debug-Log ("Attempting to load production AD data for domain: $domain") -Type "Info"

  ## Get Domain Controllers
  Update-Status "Enumerating DCs..." -spinner
  $dcs = Invoke-AD { Get-ADDomainController -Filter * -Server $domain -ErrorAction Stop } -SuppressError

  ## Get Users
  Update-Status "Enumerating Users..." -spinner
  $users = Invoke-AD { Get-ADUser -Filter * -Server $domain -Properties DisplayName,EmailAddress,Title,Department,Enabled,LockedOut,DistinguishedName -ErrorAction Stop } -SuppressError

  ## Get Groups
  Update-Status "Enumerating Groups..." -spinner
  $groups = Invoke-AD { Get-ADGroup -Filter * -Server $domain -Properties Description,GroupCategory,GroupScope,Members,DistinguishedName -ErrorAction Stop } -SuppressError

  ## Get Computers
  Update-Status "Enumerating Computers..." -spinner
  $computers = Invoke-AD { Get-ADComputer -Filter * -Server $domain -Properties OperatingSystem,OperatingSystemVersion,Enabled,LastLogonDate,DistinguishedName -ErrorAction Stop } -SuppressError

  ## Check if we got any data
  if ($null -eq $dcs -or $null -eq $users -or $null -eq $groups -or $null -eq $computers) {
    Debug-Log ("One or more AD queries returned null - falling back to DEMO mode") -Type "Warn"
      $Script:DemoMode = $true
      $converted = Convert-DataToADObjects -Users $Script:rawUsers -DCs $Script:rawDCs -Groups $Script:rawDemoGroups -Computers $Script:rawComputers -Domain $domain
      return $converted
    }

    Debug-Log ("Successfully retrieved AD data: $($dcs.Count) DCs, $($users.Count) Users, $($groups.Count) Groups, $($computers.Count) Computers") -Type "Info"

    ## In PRODUCTION, just set the Script variables directly - no conversion needed!
    Update-Status "Setting AD data..." -spinner
    $Script:Users = $users
    $Script:Groups = $groups
    $Script:DCs = $dcs
    $Script:Computers = $computers
    $Script:ADObjects = $users + $dcs + $groups + $computers

    Debug-Log ("Script variables set: Users=$($Script:Users.Count), Groups=$($Script:Groups.Count), DCs=$($Script:DCs.Count), Computers=$($Script:Computers.Count)") -Type "Info"

    ## Return hashtable for consistency
    return @{
      Users     = $users
      DCs       = $dcs
      Groups    = $groups
      Computers = $computers
    }

  } catch {
    Debug-Log ("Unexpected error loading production data: $_") -Type "Warn"
    Write-Warning "Failed to load production Active Directory data. Falling back to DEMO mode."
    $Script:DemoMode = $true

    ## Load demo data
    $converted = Convert-DataToADObjects -Users $Script:rawUsers -DCs $Script:rawDCs -Groups $Script:rawDemoGroups -Computers $Script:rawComputers -Domain $domain
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

  Debug-Log ": Converting demo data to AD-like objects..." -Type "Debug"

  ## Helper functions
  function New-FakeGuid { [guid]::NewGuid().ToString() }
  function New-FakeSid {
    $rid = Get-Random -Minimum 1000 -Maximum 65535
    "S-1-5-21-{0}-{1}-{2}-{3}" -f (Get-Random -Max 999999999), (Get-Random -Max 999999999), (Get-Random -Max 999999999), $rid
  }

  ## ========================= Users =========================
  $convertedUsers = @()
  foreach ($user in $Users) {
    ## SamAccountName derived from name
    $sam = ($user.Name -replace '\s+', '.').ToLower()
    ## UPN: prefer email if present, else generate
    $upn = if ($user.Email) { $user.Email } else { "$sam@$Domain" }

    ## Determine domain from email or default
    $userDomain = if ($user.Email -and $user.Email -match '@(.+)$') { $matches[1] } else { $Domain }

    ## Build DN properly: reverse OU array if present
    if ($user.OU) {
      $ouChain = $user.OU | ForEach-Object { "OU=$_" }
      [array]::Reverse($ouChain)   # reverse the array in-place
      $dn = "CN=$($user.Name)," + ($ouChain -join ',') + ",$BaseDN"
    } else {
      $dn = "CN=$($user.Name),$BaseDN"
    }

    $adUser = [PSCustomObject]@{
      ObjectClass                   = 'user'
      Name                          = $user.Name
      Domain                        = $userDomain
      SamAccountName                = $sam
      UserPrincipalName             = $upn
      DisplayName                   = $user.Name
      GivenName                     = ($user.Name -split '\s+')[0]
      Surname                       = ($user.Name -split '\s+')[-1]
      DistinguishedName             = $dn
      ObjectGUID                    = New-FakeGuid
      SID                           = New-FakeSid
      Enabled                       = (-not $user.Disabled)
      Disabled                      = [bool]$user.Disabled
      LockedOut                     = [bool]$user.Locked
      Locked                        = [bool]$user.Locked
      PasswordExpired               = [bool]$user.MustChangePassword
      PasswordLastSet               = (Get-Date).AddDays(-30)
      PasswordNeverExpires          = $false
      CannotChangePassword          = $false
      PasswordNotRequired           = $false
      pwdLastSet                    = if ($user.MustChangePassword) { 0 } else { 134091109521968821 }
      LastLogonDate                 = (Get-Date).AddHours(-(Get-Random -Minimum 1 -Maximum 72))
      LastBadPasswordAttempt        = if ($user.Locked) { (Get-Date).AddMinutes(-5) } else { $null }
      BadLogonCount                 = if ($user.Locked) { (Get-Random -Minimum 3 -Maximum 10) } else { 0 }
      badPwdCount                   = if ($user.Locked) { (Get-Random -Minimum 3 -Maximum 10) } else { 0 }
      LogonCount                    = Get-Random -Minimum 100 -Maximum 10000
      AccountExpirationDate         = $null
      accountExpires                = 9223372036854775807
      Title                         = $user.Title
      Department                    = $user.Department
      Company                       = $user.Company
      Manager                       = $user.Manager
      EmailAddress                  = $user.Email
      mail                          = $user.Email
      OfficePhone                   = $user.Phone
      MobilePhone                   = $user.MobilePhone
      Office                        = $user.Office
      StreetAddress                 = $user.Street
      City                          = $user.City
      PostalCode                    = $user.PostalCode
      Country                       = $user.Country
      ProfilePath                   = ""
      ScriptPath                    = ""
      HomeDirectory                 = ""
      HomeDrive                     = ""
      HomedirRequired               = $false
      TerminalServicesProfilePath   = ""
      TerminalServicesHomeDirectory = ""
      TerminalServicesHomeDrive     = ""
      Description                   = $user.Description
      MemberOf                      = $user.Groups
      CanonicalName                 = ($user.OU -join '/') + "/$($user.Name)"
      whenCreated                   = (Get-Date).AddDays(-90)
      whenChanged                   = (Get-Date).AddDays(-5)
      Created                       = (Get-Date).AddDays(-90)
      Modified                      = (Get-Date).AddDays(-5)
      OU                            = $user.OU
      Groups                        = $user.Groups
    }

    $adUser.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADUser')
    $convertedUsers += $adUser
  }

  ## ========================= DCs =========================
  $convertedDCs = @()
  foreach ($dc in $DCs) {
    $dn = "CN=$($dc.Name),OU=Domain Controllers,$BaseDN"
    $dcDomain = if ($dc.Location -match 'Germany') { 'example.net' } else { $Domain }

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

  ## ========================= Groups =========================
  $convertedGroups = @()
  foreach ($group in $Groups) {
    $sam = ($group.Name -replace '\s+', '.').ToLower()
    $dn  = "CN=$($group.Name),OU=Groups,$BaseDN"
    $groupDomain = if ($group.Email -and $group.Email -match '@(.+)$') { $matches[1] } else { $Domain }

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

    $adGroup.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADGroup')
    $convertedGroups += $adGroup
  }

  ## ========================= Computers =========================
  $convertedComputers = @()
  foreach ($computer in $Computers) {
    $dn = if ($computer.OU) {
      $ouChain = $computer.OU | ForEach-Object { "OU=$_" }
      "CN=$($computer.Name)," + ($ouChain[-1..0] -join ',') + ",$BaseDN"
    } else {
      "CN=$($computer.Name),CN=Computers,$BaseDN"
    }

    $compDomain = if ($computer.Location -match 'Germany') { 'example.net' } else { $Domain }

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
      ComputerType           = $computer.Type
      whenCreated            = (Get-Date).AddDays(-120)
      OU                     = $computer.OU
    }

    $adComputer.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADComputer')
    $convertedComputers += $adComputer
  }

  ## ========================= Set global variables =========================
  $Script:Users     = $convertedUsers
  $Script:Computers = $convertedComputers
  $Script:DCs       = $convertedDCs
  $Script:Groups    = $convertedGroups
  $Script:ADObjects = $convertedUsers + $convertedDCs + $convertedGroups + $convertedComputers

  Debug-Log ": Converted $($convertedUsers.Count) users, $($convertedDCs.Count) DCs, $($convertedComputers.Count) computers, and $($convertedGroups.Count) groups to AD-like objects" -Type "Debug"

  ## Return hashtable
  return @{
    Users     = $convertedUsers
    DCs       = $convertedDCs
    Groups    = $convertedGroups
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

## Helper Function: Parse Distinguished Name
function Get-OUPathFromDN {
  <#
  .SYNOPSIS
  Extracts OU hierarchy from a Distinguished Name

  .DESCRIPTION
  Parses an AD Distinguished Name and returns the OU path as an array,
  ordered from top-level OU down to the immediate parent OU.

  .PARAMETER DistinguishedName
  The full Distinguished Name from Active Directory

  .EXAMPLE
  Get-OUPathFromDN "CN=John Doe,OU=UK,OU=Users,OU=here.corp,DC=here,DC=corp"
  Returns: @("here.corp", "Users", "UK")

  .EXAMPLE
  Get-OUPathFromDN "CN=agadmin,CN=Users,DC=here,DC=corp"
  Returns: @("Users")  # CN=Users is treated as an OU
  #>

  param(
    [string]$DistinguishedName
  )

  if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return @() }

  ## Split DN by comma (handling escaped commas in quotes)
  $parts = $DistinguishedName -split ',(?=(?:[^"]*"[^"]*")*[^"]*$)'
  $ouList = @()

  foreach ($part in $parts) {
    $part = $part.Trim()

    ## Match OU= or CN= (CN=Users is a container, treat it like an OU)
    if ($part -match '^(OU|CN)=(.+)$') {
      $type = $matches[1]
      $value = $matches[2]

      ## Skip the actual user/group CN (first one)
      if ($type -eq 'CN' -and $ouList.Count -eq 0) { continue }

      ## Add to list (we'll reverse it later)
      $ouList += $value
    }

    ## Stop at DC= (domain components)
    elseif ($part -match '^DC=') { break }
  }

  ## Reverse to get top-down hierarchy
  [array]::Reverse($ouList)
  return $ouList
}

# Helper Function: Get Domain from DN
function Get-DomainFromDN {
  <#
  .SYNOPSIS
  Extracts domain name from Distinguished Name

  .PARAMETER DistinguishedName
  The full Distinguished Name from Active Directory

  .EXAMPLE
  Get-DomainFromDN "CN=John Doe,OU=UK,OU=Users,DC=here,DC=corp"
  Returns: "here.corp"
  #>

  param(
    [string]$DistinguishedName
  )

  if ([string]::IsNullOrWhiteSpace($DistinguishedName)) {
    return $null
  }

  ## Extract all DC= parts
  $dcParts = [regex]::Matches($DistinguishedName, 'DC=([^,]+)') |
  ForEach-Object { $_.Groups[1].Value }

  if ($dcParts.Count -gt 0) {
    return $dcParts -join '.'
  }

  return $null
}

function Build-DomainContent {
  param(
    [Terminal.Gui.Trees.TreeNode]$domainNode,
    [string]$domain
  )

  Debug-Log (": Building content for domain: $domain") -Type "Info"

  ## Apply filters - filter users by domain
  $nameFilter = $Script:FilterOptions.NameFilter.Trim()

  ## Filter users for this specific domain
  $domainUsers = $Script:Users | Where-Object {  $_.Domain -eq $domain -or (Get-DomainFromDN $_.DistinguishedName) -eq $domain }
  $filteredUsers = $domainUsers | Where-Object {  ($_.Disabled -and $Script:FilterOptions.ShowDisabledUsers) -or   (-not $_.Disabled -and $Script:FilterOptions.ShowEnabledUsers) }

  if ($nameFilter) {
    $filteredUsers = $filteredUsers | Where-Object {  $_.Name -like "*$nameFilter*" -or  $_.EmailAddress -like "*$nameFilter*" -or $_.mail -like "*$nameFilter*" -or $_.Title -like "*$nameFilter*" }
  }

  Debug-Log (": Filtered to $($filteredUsers.Count) users for domain $domain") -Type "Info"

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
      if ($null -eq $Parent.Children) { $Parent | Add-Member -Force -MemberType NoteProperty -Name Children -Value (New-Object 'System.Collections.ObjectModel.Collection[Terminal.Gui.Trees.ITreeNode]') }
        $Parent.Children.Add($newNode)
      } catch {
        Debug-Log (": Could not add node $Name to parent: $_") -Type "Warn"
      }

      $nodeCache[$FullPath] = $newNode
      return $newNode
  }

  ## Helper function to parse Distinguished Name and extract OU path
  function Get-OUPathFromDN {
    param([string]$DistinguishedName)

    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return @() }

    ## Split DN by comma (handling escaped commas in quotes)
        $parts = $DistinguishedName -split ',(?=(?:[^"]*"[^"]*")*[^"]*$)'

        $ouList = @()
        $skipFirst = $true

        foreach ($part in $parts) {
            $part = $part.Trim()

            # Match OU= or CN= (CN=Users is a container, treat it like an OU)
            if ($part -match '^(OU|CN)=(.+)$') {
                $type = $matches[1]
                $value = $matches[2]

                # Skip the actual user/group CN (first one only)
                if ($skipFirst) {
                    $skipFirst = $false
                    continue
                }

                # Add to list (we'll reverse it later)
                $ouList += $value
            }
            # Stop at DC= (domain components)
            elseif ($part -match '^DC=') {
                break
            }
        }

        # Reverse to get top-down hierarchy
        [array]::Reverse($ouList)

        return $ouList
    }

    ## Helper function to get domain from DN
    function Get-DomainFromDN {
        param([string]$DistinguishedName)

        if ([string]::IsNullOrWhiteSpace($DistinguishedName)) {
            return $null
        }

        # Extract all DC= parts
        $dcParts = [regex]::Matches($DistinguishedName, 'DC=([^,]+)') |
            ForEach-Object { $_.Groups[1].Value }

        if ($dcParts.Count -gt 0) {
            return $dcParts -join '.'
        }

        return $null
    }

    ## Build OU hierarchy for this domain
    $userCount = 0
    foreach ($user in $filteredUsers) {
        $userCount++

        # Update every 25 users to avoid spam
        if ($userCount % 25 -eq 0) {
            Update-Status "Building $domain - $userCount/$($filteredUsers.Count) users" -spinner
        }

        # Parse OU path from DistinguishedName
        $ouPath = Get-OUPathFromDN $user.DistinguishedName

        if ($ouPath.Count -eq 0) {
            Debug-Log (": User $($user.Name) has no OU path, skipping tree placement") -Type "Warn"
            continue
        }

        $currentNode = $domainNode
        $pathSoFar = ""

        # Build the OU hierarchy
        foreach ($ouLevel in $ouPath) {
            $pathSoFar = if ($pathSoFar) { "$pathSoFar/$ouLevel" } else { $ouLevel }
            $currentNode = Get-OrCreateChildNode -Parent $currentNode -Name $ouLevel -FullPath "$domain/$pathSoFar"
        }

        # Determine status icon
        $statusIcon = if ($user.LockedOut -or $user.Locked) {
            "🔒"
        } elseif (-not $user.Enabled -or $user.Disabled) {
            "⊗"
        } else {
            "○"
        }

        # Create user node
        $userNode = [Terminal.Gui.Trees.TreeNode]::new("(U) $statusIcon $($user.Name)")
        $userNode.Tag = $user

        if ($null -eq $currentNode.Children) {
            $currentNode | Add-Member -Force -MemberType NoteProperty -Name Children -Value (New-Object 'System.Collections.ObjectModel.Collection[Terminal.Gui.Trees.ITreeNode]')
        }
        $currentNode.Children.Add($userNode)
    }

    ## Add Groups for this domain
    if ($Script:FilterOptions.ShowGroups) {
        Update-Status "Building $domain - adding groups" -spinner

        # Filter groups for this domain
        $domainGroups = $Script:Groups | Where-Object {
            $_.Domain -eq $domain -or
            (Get-DomainFromDN $_.DistinguishedName) -eq $domain
        }

        if ($domainGroups.Count -gt 0) {
            $groupsNode = Get-OrCreateChildNode -Parent $domainNode -Name "Groups" -FullPath "$domain/_Groups"

            foreach ($group in ($domainGroups | Sort-Object -Property Name)) {
                $groupNode = [Terminal.Gui.Trees.TreeNode]::new("(G) $($group.Name)")
                $groupNode.Tag = $group

                # Find members of this group
                $members = $filteredUsers | Where-Object {
                    $_.Groups -contains $group.Name -or
                    $_.MemberOf -contains $group.DistinguishedName
                } | Sort-Object -Property Name

                foreach ($member in $members) {
                    $statusIcon = if ($member.LockedOut -or $member.Locked) {
                        "🔒"
                    } elseif (-not $member.Enabled -or $member.Disabled) {
                        "⊗"
                    } else {
                        "○"
                    }

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
    if ($Script:FilterOptions.ShowDCs) {
        Update-Status "Building $domain - adding DCs" -spinner

        # Filter DCs for this domain
        $domainDCs = $Script:DCs | Where-Object {
            $_.Domain -eq $domain -or
            (Get-DomainFromDN $_.DistinguishedName) -eq $domain
        }

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
    if ($Script:FilterOptions.ShowComputers) {
        Update-Status "Building $domain - adding computers" -spinner

        # Filter computers for this domain
        $domainComputers = $Script:Computers | Where-Object {
            $_.Domain -eq $domain -or
            (Get-DomainFromDN $_.DistinguishedName) -eq $domain
        }

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
    Debug-Log (": Finished building content for domain $domain - nodes in cache: $($nodeCache.Count)") -Type "Success"
}

##---------------------------{ Build The Tree }--------------------------
function Build-Tree {
    param([string]$domain)

    if (-not $domain) {
        $domain = $Script:CurrentDomain
    }

    Debug-Log ": Building tree..." -Type "Info"

    # TreeView MUST already exist in 1.16
    if ($null -eq $Script:tree) {
        throw "Build-Tree failed: TreeView does not exist"
    }

    # Clear existing objects safely
    try {
        $Script:tree.ClearObjects()
    } catch {
        Debug-Log ": WARNING - ClearObjects failed: $_" -Type "Warn"
    }

    $root = $null

    ## ================= Multi-domain forest =================
    if ($Script:Domains.Count -gt 1) {
        Debug-Log ": Creating multi-domain forest tree with root: $($Script:ForestName)" -Type "Info"

        $root = [Terminal.Gui.Trees.TreeNode]::new($Script:ForestName)

        foreach ($dom in $Script:Domains) {
            Debug-Log ": Adding domain node: $($dom)" -Type "Info"

            $domainNode = [Terminal.Gui.Trees.TreeNode]::new($dom)
            $root.Children.Add($domainNode)

            Build-DomainContent -domainNode $domainNode -domain $dom
        }

        if (-not $Script:CurrentDomain) {
            $Script:CurrentDomain = $Script:Domains[0]
            Debug-Log ": CurrentDomain was empty. Auto-selected: $($Script:CurrentDomain)" -Type "Info"
        }
    }

    ## ================= Single-domain =================
    else {
        Debug-Log ": Creating single-domain tree: $($Script:Domains[0])" -Type "Info"

        $root = [Terminal.Gui.Trees.TreeNode]::new($Script:Domains[0])
        Build-DomainContent -domainNode $root -domain $Script:Domains[0]

        if (-not $Script:CurrentDomain) {
            $Script:CurrentDomain = $Script:Domains[0]
            Debug-Log ": CurrentDomain set (single-domain): $($Script:CurrentDomain)" -Type "Info"
        }
    }

    if ($null -eq $root) {
        throw "Build-Tree failed: root node is null"
    }

    ## ================= Attach root to TreeView =================
    try {
        $Script:tree.AddObject($root)
        $Script:tree.SelectedObject = $root
        Debug-Log ": Root node added to TreeView" -Type "Success"
    } catch {
        throw "Build-Tree failed while attaching root to TreeView: $_"
    }

    Debug-Log ": Build-Tree completed successfully" -Type "Success"
}


## -----------------------{ Update Filter Label }----------------------
function Create-FilterStatusLabel {
    Debug-Log (": Entered Create-FilterStatusLabel") -Type "Info"
    try {
        $lblStatus = [Terminal.Gui.Label]::new("")
        Debug-Log (": Created label object: $lblStatus") -Type "Info"
        $lblStatus.X = 80
        $lblStatus.Y = 6
        $lblStatus.Width = 40
        return $lblStatus
    }
    catch {
        Debug-Log ("ERROR in Create-FilterStatusLabel: $($_.Exception.Message)") -Type "Warn"
        return $null
    }
}

## -----------------------{ Update Filter Label }----------------------
function Update-FilterStatusLabel {
  param($label)

  if (-not $label) {
    Debug-Log (": label parameter is null in Update-FilterStatusLabel") -Type "Warn"
    return
  }

  $activeFilters = @()
  if (-not $Script:FilterOptions.ShowEnabledUsers) { $activeFilters += "No Enabled" }
  if (-not $Script:FilterOptions.ShowDisabledUsers) { $activeFilters += "No Disabled" }
  if (-not $Script:FilterOptions.ShowGroups) { $activeFilters += "No Groups" }
  if (-not $Script:FilterOptions.ShowDCs) { $activeFilters += "No DCs" }
  if ($Script:FilterOptions.NameFilter) { $activeFilters += "Name:$($Script:FilterOptions.NameFilter)" }
  if ($activeFilters.Count -gt 0) {
    $label.Text = "Active Filters: " + ($activeFilters -join ", ")
  } else {
### TODO: This one needs adding into the filter modal not main window
    $label.Text = "No filters active (showing all)"
  }
}

## ------------------------{ Update Status Bar }-----------------------
function Start-Spinner {
  param([string]$baseMessage)

  $Script:SpinnerActive = $true
  $Script:statusBaseMessage = $baseMessage
  $Script:statusSpinnerIndex = 0

  ## Cancel existing timer if any
  if ($Script:SpinnerTimer) {
    Stop-Spinner
  }

  ## Create a timer that updates every 200ms
  $Script:SpinnerTimer = [System.Timers.Timer]::new(200)
  $Script:SpinnerTimer.AutoReset = $true

  ## Timer callback - updates spinner character
  $Script:SpinnerTimer.Add_Elapsed({
    if ($Script:SpinnerActive -and $Script:StatusItem) {
      $Script:statusSpinnerIndex = ($Script:statusSpinnerIndex + 1) % 4
      $spinChar = $Script:statusSpinner[$Script:statusSpinnerIndex]
      $staticPrefix = "Forest: $($Script:ForestName) | Objs: $($Script:ADObjects.Count)"
      $Script:StatusItem.Title = "$staticPrefix | $spinChar $($Script:statusBaseMessage)"

      try {
        if ($Script:statusBar) {
          $Script:statusBar.SetNeedsDisplay()
        }
      } catch {
        ## Ignore display errors
        }
      }
  })

  $Script:SpinnerTimer.Start()
}

function Stop-Spinner {
  $Script:SpinnerActive = $false

  if ($Script:SpinnerTimer) {
    $Script:SpinnerTimer.Stop()
    $Script:SpinnerTimer.Dispose()
    $Script:SpinnerTimer = $null
  }
}

function Update-Status {
  param(
    [string]$message,
    [switch]$spinner,
    [switch]$final
  )

  Debug-Log (": Update-Status called - message='$message', spinner=$spinner, final=$final") -Type "Info"

  if (-not $Script:StatusItem) {
    Debug-Log (": StatusItem not available") -Type "Info"
    return
  }

  $staticPrefix = "Forest: $($Script:ForestName) | Objs: $($Script:ADObjects.Count)"

  if ($final) {
    Stop-Spinner
    $displayText = "$staticPrefix | ✓ $message"
    Debug-Log (": Setting final status to: '$displayText'") -Type "Success"
    $Script:StatusItem.Title = $displayText
    Start-Sleep -milliseconds 50
  } elseif ($spinner) {
    ## Start the spinner with this message
    Start-Spinner -baseMessage $message
    Debug-Log (": $message") -Type "Info"
    Start-Sleep -milliseconds 50
  } else {
    Stop-Spinner
    Debug-Log (": Setting static status to: '$message'") -Type "Info"
    $Script:StatusItem.Title = "$staticPrefix | $message"
    Start-Sleep -milliseconds 50
  }

  try {
    if ($Script:statusBar) {
      $Script:statusBar.SetNeedsDisplay()
    }
  } catch {
    ## Ignore
  }
}

# ------------------------- Filter Panel (Add to main window) ------------------------
function Create-FilterPanel {

  ## Guard rail so the script doesn't blow up if this is not defined
  if (-not $Script:FilterOptions) {
    Debug-Log "FilterOptions not initialised — initialising defaults" -Type "Warn"
    $Script:FilterOptions = @{
      ShowDisabledUsers = $true
      ShowEnabledUsers  = $true
      ShowLockedUsers   = $true
      ShowGroups        = $true
      ShowDCs           = $true
      ShowComputers     = $true
      ShowOUs           = $true
      NameFilter        = ""
      SortBy            = "Name"
      SortDescending    = $false
    }
  }

  ## Create a frame for filters
  $filterFrame = [Terminal.Gui.FrameView]::new("Filters")
  $filterFrame.X = 32  # Right of the tree
  $filterFrame.Y = 1
  $filterFrame.Width = 40
  $filterFrame.Height = 20  # Increased to fit status label
  $y = 0

  ## Name filter
  $lblNameFilter = [Terminal.Gui.Label]::new("Name contains:"); $lblNameFilter.X=1; $lblNameFilter.Y=$y; $filterFrame.Add($lblNameFilter)
  $txtNameFilter = [Terminal.Gui.TextField]::new($Script:FilterOptions.NameFilter)
  $txtNameFilter.X=1; $txtNameFilter.Y=$y+1; $txtNameFilter.Width=35
  $txtNameFilter.add_TextChanged({ $Script:FilterOptions.NameFilter = $txtNameFilter.Text.ToString() })
  $filterFrame.Add($txtNameFilter)
  $y+=3

  ## Show/Hide checkboxes
  $chkEnabled = [Terminal.Gui.CheckBox]::new("Show Enabled Users")
  $chkEnabled.X=1; $chkEnabled.Y=$y; $chkEnabled.Checked=$Script:FilterOptions.ShowEnabledUsers
  $chkEnabled.add_Toggled({ $Script:FilterOptions.ShowEnabledUsers = $chkEnabled.Checked })
  $filterFrame.Add($chkEnabled)
  $y+=1

  $chkLocked = [Terminal.Gui.CheckBox]::new("Show Locked Users")
  $chkLocked.X=1; $chkLocked.Y=$y; $chkLocked.Checked=$Script:FilterOptions.ShowLockedUsers
  $chkLocked.add_Toggled({ $Script:FilterOptions.ShowLockedUsers = $chkLocked.Checked })
  $filterFrame.Add($chkLocked)
  $y+=1

  $chkDisabled = [Terminal.Gui.CheckBox]::new("Show Disabled Users")
  $chkDisabled.X=1; $chkDisabled.Y=$y; $chkDisabled.Checked=$Script:FilterOptions.ShowDisabledUsers
  $chkDisabled.add_Toggled({ $Script:FilterOptions.ShowDisabledUsers = $chkDisabled.Checked })
  $filterFrame.Add($chkDisabled)
  $y+=1

  $chkGroups = [Terminal.Gui.CheckBox]::new("Show Groups")
  $chkGroups.X=1; $chkGroups.Y=$y; $chkGroups.Checked=$Script:FilterOptions.ShowGroups
  $chkGroups.add_Toggled({ $Script:FilterOptions.ShowGroups = $chkGroups.Checked })
  $filterFrame.Add($chkGroups)
  $y+=1

  $chkDCs = [Terminal.Gui.CheckBox]::new("Show Domain Controllers")
  $chkDCs.X=1; $chkDCs.Y=$y; $chkDCs.Checked=$Script:FilterOptions.ShowDCs
  $chkDCs.add_Toggled({ $Script:FilterOptions.ShowDCs = $chkDCs.Checked })
  $filterFrame.Add($chkDCs)
  $y+=2

  ## Sort options
  $lblSort = [Terminal.Gui.Label]::new("Sort by:"); $lblSort.X=1; $lblSort.Y=$y; $filterFrame.Add($lblSort)
  $y+=1

  $rdoSort = [Terminal.Gui.RadioGroup]::new(@("Name", "Type", "OU"))
  $rdoSort.X=1; $rdoSort.Y=$y; $rdoSort.SelectedItem=0
  $rdoSort.add_SelectedItemChanged({
    switch ($rdoSort.SelectedItem) {
      0 { $Script:FilterOptions.SortBy = "Name" }
      1 { $Script:FilterOptions.SortBy = "Type" }
      2 { $Script:FilterOptions.SortBy = "OU" }
    }
  })
  $filterFrame.Add($rdoSort)
  $y+=4
  Debug-Log "Y value is: $y" -Type "Debug"

  ## Apply/Reset buttons
  $btnApplyFilter = [Terminal.Gui.Button]::new("Apply Filter")
  $btnApplyFilter.X=1; $btnApplyFilter.Y=$y
  $btnApplyFilter.add_Clicked({
    Debug-Log "Applying filters..." -Type "Info"

    # Rebuild tree with filters
    $rootNode = Build-Tree -domain $Script:CurrentDomain
    if ($rootNode) {
      $Script:tree.ClearObjects()
      $Script:tree.AddObject($rootNode)
      [Terminal.Gui.Application]::Refresh()
    }

    # Update status label
    Update-FilterStatusLabel -label $Script:FilterStatusLabel
  })
  $filterFrame.Add($btnApplyFilter)

  $btnResetFilter = [Terminal.Gui.Button]::new("Reset")
  $btnResetFilter.X=21; $btnResetFilter.Y=$y
  $btnResetFilter.add_Clicked({
    Debug-Log "Resetting filters..." -Type "Info"
    $Script:FilterOptions.ShowDisabledUsers = $true
    $Script:FilterOptions.ShowEnabledUsers = $true
    $Script:FilterOptions.ShowLockedUsers = $true
    $Script:FilterOptions.ShowGroups = $true
    $Script:FilterOptions.ShowDCs = $true
    $Script:FilterOptions.ShowComputers = $true
    $Script:FilterOptions.ShowOUs = $true
    $Script:FilterOptions.NameFilter = ""
    $Script:FilterOptions.SortBy = "Name"
    $Script:FilterOptions.SortDescending = $false

    ## Reset UI controls
    $chkEnabled.Checked = $true
    $chkDisabled.Checked = $true
    $chkGroups.Checked = $true
    $chkDCs.Checked = $true
    $txtNameFilter.Text = ""
    $rdoSort.SelectedItem = 0

    ## Rebuild tree
    $rootNode = Build-Tree -domain $Script:CurrentDomain
    if ($rootNode) {
      $Script:tree.ClearObjects()
      $Script:tree.AddObject($rootNode)
      [Terminal.Gui.Application]::Refresh()
    }

    ## Update status label
    Update-FilterStatusLabel -label $Script:FilterStatusLabel
  })
  $filterFrame.Add($btnResetFilter)
  $y+=2

  ## -------------------- Filter Status Label (inside panel) --------------------
  $Script:FilterStatusLabel = [Terminal.Gui.Label]::new("")
  $Script:FilterStatusLabel.X = 1
  $Script:FilterStatusLabel.Y = $y
  $Script:FilterStatusLabel.Width = [Terminal.Gui.Dim]::Fill(1)
  $filterFrame.Add($Script:FilterStatusLabel)

  Debug-Log "FilterPanel created with embedded status label" -Type "Info"

  ## Apply theme correctly
  if ($Script:themeData -and $Script:themeData.MainWindow) {
    $filterFrame.ColorScheme = $Script:themeData.MainWindow
  }

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
    Debug-Log (": Applying quick filter: $selected") -Type "Info"

    switch ($selected) {
      "Show All" {
        $Script:FilterOptions.ShowLockedUsers = $true
        $Script:FilterOptions.ShowDisabledUsers = $true
        $Script:FilterOptions.ShowEnabledUsers = $true
        $Script:FilterOptions.ShowGroups = $true
        $Script:FilterOptions.ShowDCs = $true
        $Script:FilterOptions.NameFilter = ""
        }
      "Locked Users Only" {
        $Script:FilterOptions.ShowLockedUsers = $true
        $Script:FilterOptions.ShowEnabledUsers = $false
        $Script:FilterOptions.ShowDisabledUsers = $false
        }
      "Disabled Users Only" {
        $Script:FilterOptions.ShowDisabledUsers = $true
        $Script:FilterOptions.ShowEnabledUsers = $false
        }
      "Enabled Users Only" {
        $Script:FilterOptions.ShowDisabledUsers = $false
        $Script:FilterOptions.ShowEnabledUsers = $true
        }
      "Users Never Logged In" {
        ## This would require additional logic to track last logon
        Show-Modal "Filter" "Filter applied: $selected"
        }
      "Users with No Manager" {
        ## Filter users with no manager
        $Script:FilterOptions.NameFilter = ""
        }
      "Empty Groups" {
        ## Show only groups with no members
        Show-Modal "Filter" "Filter applied: $selected"
        }
      "Domain Admins Only" {
        $Script:FilterOptions.NameFilter = ""
        }
      }

      Build-Tree -domain $Script:CurrentDomain
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
## | - Uses $Script:CurrentDomain for the domain to query                   |
## | - auto-refresh Domain if it changes                                    |
## | - No F-keys, no Ctrl+Q/Ctrl+D; uses Alt+... + Esc as needed            |
## '------------------------------------------------------------------------'
function Get-ADHealth {

  ## --------------------------------{ Layout }-------------------------------
  $width  = 118
  $height = 36

  ## Ensure cache / state exists
  if (-not $Script:ADHealth) { $Script:ADHealth = @{} }
  if (-not $Script:ADHealth.Cache) { $Script:ADHealth.Cache = @{} }

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
    if (-not $Script:CurrentDomain -or [string]::IsNullOrWhiteSpace($Script:CurrentDomain)) {
      throw "Global variable `$Script:CurrentDomain is not set. Set `$Script:CurrentDomain before calling Get-ADHealth."
    }
    $domain = $Script:CurrentDomain

    ## If domain changed since last run, clear cached data
    if ($Script:ADHealth.Cache.Domain -ne $domain) {
      $Script:ADHealth.Cache = @{
        Domain = $domain
        Timestamp = (Get-Date).ToString("s")
      }
    }

  ## Prepare blank data containers for fresh run
  $Script:ADHealth.Data = @{
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
    $Script:ADHealth.Tools = @{
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
      $dcs = Get-ADDomainController -Filter * -Server $domain -ErrorAction Stop
    } catch {
      $Script:ADHealth.Data.DCStatus.Details = "Failed to enumerate DCs: $($_.Exception.Message)"
      $Script:ADHealth.Data.DCStatus.Summary = @(": Could not enumerate domain controllers.")
      $Script:ADHealth.Data.DCStatus.Health = "FAIL"
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
      $Script:ADHealth.Data.DCStatus.Summary = $summary
      $Script:ADHealth.Data.DCStatus.Details = $sb.ToString()
      $Script:ADHealth.Data.DCStatus.Health = $health
    }

  ## Replication
  function Run-ReplicationCheck {
    param([string]$Domain)
    $sb = New-Object System.Text.StringBuilder
    $summary = @()

    if ($Script:ADHealth.Tools.Repadmin) {
      $r = Invoke-External -Exe "repadmin.exe" -Args "/replsummary"
      $sb.AppendLine($r.StdOut)
      if ($r.ExitCode -eq 0) {
        $lines = ($r.StdOut -split "`r?`n") | Select-Object -First 12
        $summary += "repadmin /replsummary (top lines):"
        $summary += $lines
        ## Simple health: if the output contains "failed" or "error" mark WARN
        if ($r.StdOut -imatch "(fail|error|last attempt)") { $health = "FAIL" }
      } else {
        $summary += "repadmin returned : $($r.StdErr)"
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

      $Script:ADHealth.Data.Replication.Summary = $summary
      $Script:ADHealth.Data.Replication.Details = $sb.ToString()
      $Script:ADHealth.Data.Replication.Health = $health
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
      $dcs = Get-ADDomainController -Filter * -Server $domain -ErrorAction Stop
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
    $Script:ADHealth.Data.DNS.Summary = $summary
    $Script:ADHealth.Data.DNS.Details = $sb.ToString()
    $Script:ADHealth.Data.DNS.Health = $health
  }

  ## SYSVOL / DFSR
  function Run-SYSVOLCheck {
    param([string]$Domain)
    $summary = @()
      $sb = New-Object System.Text.StringBuilder

      try {
        $dcs = Get-ADDomainController -Discover -Filter * -DomainName $domain -ErrorAction Stop
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

          if ($Script:ADHealth.Tools.DfsrDiag) {
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

      $Script:ADHealth.Data.SYSVOL.Summary = $summary
      $Script:ADHealth.Data.SYSVOL.Details = $sb.ToString()
      $Script:ADHealth.Data.SYSVOL.Health = $health
  }

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

      ## Health: check if any holder resolves and is reachable
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

    $Script:ADHealth.Data.FSMO.Summary = $summary
    $Script:ADHealth.Data.FSMO.Details = $sb.ToString()
    $Script:ADHealth.Data.FSMO.Health = $health
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

      ## Quick version compare for first 20 GPOs
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

      $Script:ADHealth.Data.GPO.Summary = $summary
      $Script:ADHealth.Data.GPO.Details = $sb.ToString()
      $Script:ADHealth.Data.GPO.Health = $health
    }

    # ----------------------------
    # Run all checks orchestrator
    # ----------------------------
    function Run-AllChecks {
      param([string]$Domain)
      Initialize-ADHealth
      try { Run-DCStatusCheck -Domain $Domain } catch { $Script:ADHealth.Data.DCStatus.Details = "DCStatus failed: $($_.Exception.Message)"; $Script:ADHealth.Data.DCStatus.Health="WARN" }
      try { Run-ReplicationCheck -Domain $Domain } catch { $Script:ADHealth.Data.Replication.Details = "Replication failed: $($_.Exception.Message)"; $Script:ADHealth.Data.Replication.Health="WARN" }
      try { Run-DNSCheck -Domain $Domain } catch { $Script:ADHealth.Data.DNS.Details = "DNS failed: $($_.Exception.Message)"; $Script:ADHealth.Data.DNS.Health="WARN" }
      try { Run-SYSVOLCheck -Domain $Domain } catch { $Script:ADHealth.Data.SYSVOL.Details = "SYSVOL failed: $($_.Exception.Message)"; $Script:ADHealth.Data.SYSVOL.Health="WARN" }
      try { Run-FSMOCheck -Domain $Domain } catch { $Script:ADHealth.Data.FSMO.Details = "FSMO failed: $($_.Exception.Message)"; $Script:ADHealth.Data.FSMO.Health="WARN" }
      try { Run-GPOCheck -Domain $Domain } catch { $Script:ADHealth.Data.GPO.Details = "GPO failed: $($_.Exception.Message)"; $Script:ADHealth.Data.GPO.Health="WARN" }

      $Script:ADHealth.Cache.Timestamp = (Get-Date)
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
      ## Show up to 10 lines and place Details button bottom-right
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

      ## Details button text uses an ampersand accelerator for Alt+E: we place & before 'e' in "Details"
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

function Initialize-Globals {
  $Script:tree = [Terminal.Gui.TreeView]::new()
  $Script:TabRows = @()
  $Script:AllTabs = @()
}


## ----------------------------
## UI: Build modal & controls
## ----------------------------
$modal = [Terminal.Gui.Window]::new("Active Directory Health Check - $($Script:CurrentDomain)")
$modal.X = [Terminal.Gui.Pos]::Center(); $modal.Y = [Terminal.Gui.Pos]::Center()
$modal.Width = $width; $modal.Height = $height; $modal.Modal = $true

## Tools row - single-line compact
$toolsView = [Terminal.Gui.View]::new()
$toolsView.X = 1; $toolsView.Y = 1; $toolsView.Width = $width - 2; $toolsView.Height = 1
$modal.Add($toolsView)

## Tab header view (below tools)
$tabHeader = [Terminal.Gui.View]::new()
$tabHeader.X = 1; $tabHeader.Y = 2; $tabHeader.Width = $width - 2; $tabHeader.Height = 1
$modal.Add($tabHeader)

## This could be the soruce of the mystery window
## Content frame
$contentView = [Terminal.Gui.FrameView]::new()
$contentView.X = 1; $contentView.Y = 4; $contentView.Width = $width - 2; $contentView.Height = $height - 10
$contentView.Border = "Ascii"
$modal.Add($contentView)

## Help line (shortcuts)
$helpLine = [Terminal.Gui.Label]::new("Shortcuts: Alt+R=Refresh All | Alt+S=Refresh Tab | Alt+←/→=Switch Tabs | Alt+E=Details | Esc=Close")
$helpLine.X = 1; $helpLine.Y = $height - 5
$modal.Add($helpLine)

## Buttons: use ampersand to set Alt-accelerators, avoiding reserved combos
$btnRefreshAll = [Terminal.Gui.Button]::new(2, $height - 4, "&Refresh All")   # Alt+R
$btnRefreshTab = [Terminal.Gui.Button]::new(20, $height - 4, "Re&scan Tab")    # Alt+S
$btnExport = [Terminal.Gui.Button]::new(40, $height - 4, "&Export")           # Alt+E for Export (convenient)
$btnClose = [Terminal.Gui.Button]::new(60, $height - 4, "Close (Esc)")
$modal.Add($btnRefreshAll); $modal.Add($btnRefreshTab); $modal.Add($btnExport); $modal.Add($btnClose)

## Tab definitions
$tabs = @("DC Status","Replication","DNS","SYSVOL","FSMO Roles","GPO Health")
$Script:ActiveHealthTab = 0

## Function to compute simple health icon string
function Health-Icon {
  param([string]$State)
    switch ($State) {
      "OK"   { return "(✓)" }
      "WARN" { return "(!)" }
      "FAIL" { return "(✗)" }
      "default { return "(?)" }
    }
  }

   ## Render tools row (certUI-style)
function Render-ToolsRow {
  $toolsView.RemoveAll()
  $offset = 0
  $tools = $Script:ADHealth.Tools
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

    ## Add click handler that shows tooltip-like message
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

## Render tab header with health icons
function Render-ADHealthTabs {
  $tabHeader.RemoveAll()
  $offset = 0
  for ($i = 0; $i -lt $tabs.Count; $i++) {
    $name = $tabs[$i]
    $health = $Script:ADHealth.Data.$((($name -split ' ')[0])) 2>$null
    ## We stored health per named key (e.g., DCStatus, Replication, DNS...). fallback if not matching:
    switch ($i) {
      0 { $h = $Script:ADHealth.Data.DCStatus.Health }
      1 { $h = $Script:ADHealth.Data.Replication.Health }
      2 { $h = $Script:ADHealth.Data.DNS.Health }
      3 { $h = $Script:ADHealth.Data.SYSVOL.Health }
      4 { $h = $Script:ADHealth.Data.FSMO.Health }
      5 { $h = $Script:ADHealth.Data.GPO.Health }
      default { $h = "UNKNOWN" }
    }

    $icon = Health-Icon -State $h
    $labelText = if ($i -eq $Script:ActiveHealthTab) { "[ $name $icon ]" } else { "  $name $icon  " }
    $lbl = [Terminal.Gui.Label]::new($labelText)
    $lbl.X = $offset; $lbl.Y = 0
    $idx = $i
    $lbl.add_MouseClick({

    param($args)
      $Script:ActiveHealthTab = $idx
      Render-ADHealthTabs
      Render-ADHealthContent -Tab $idx
    })
    $tabHeader.Add($lbl)
    $offset += ($labelText.Length + 1)
  }
}

## Render content for each tab
function Render-ADHealthContent {
  param([int]$Tab)
    $contentView.RemoveAll()
    switch ($Tab) {
      0 {
        $sumObjs = $Script:ADHealth.Data.DCStatus.Summary
        if (-not $sumObjs -or $sumObjs.Count -eq 0) {
          $lbl = [Terminal.Gui.Label]::new("No DC status data. Press Alt+R to run checks.")
          $lbl.X = 1; $lbl.Y = 1; $contentView.Add($lbl)
        } else {
          $lines = @()
          $lines += "Name`tIP`tSite`tReach`tUptime`tNTDS`tDNS"
          foreach ($o in $sumObjs) {
            $lines += ("{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}" -f $o.Name, $o.IP, $o.Site, $o.Reachable, $o.UptimeDays, $o.NTDS, $o.DNS)
          }
          Render-SummaryWithDetails -Parent $contentView -SummaryLines $lines -DetailsText $Script:ADHealth.Data.DCStatus.Details
        }
      }
      1 {
         $lines = $Script:ADHealth.Data.Replication.Summary
         if (-not $lines) { $lines = @("No replication data. Press Alt+R.") }
         Render-SummaryWithDetails -Parent $contentView -SummaryLines $lines -DetailsText $Script:ADHealth.Data.Replication.Details
      }
      2 {
        $lines = $Script:ADHealth.Data.DNS.Summary
        if (-not $lines) { $lines = @("No DNS data. Press Alt+R.") }
        Render-SummaryWithDetails -Parent $contentView -SummaryLines $lines -DetailsText $Script:ADHealth.Data.DNS.Details
      }
      3 {
        $lines = $Script:ADHealth.Data.SYSVOL.Summary
        if (-not $lines) { $lines = @("No SYSVOL data. Press Alt+R.") }
        Render-SummaryWithDetails -Parent $contentView -SummaryLines $lines -DetailsText $Script:ADHealth.Data.SYSVOL.Details
      }
      4 {
        $lines = $Script:ADHealth.Data.FSMO.Summary
        if (-not $lines) { $lines = @("No FSMO data. Press Alt+R.") }
        Render-SummaryWithDetails -Parent $contentView -SummaryLines $lines -DetailsText $Script:ADHealth.Data.FSMO.Details
      }
      5 {
        $lines = $Script:ADHealth.Data.GPO.Summary
        if (-not $lines) { $lines = @("No GPO data. Press Alt+R.") }
        Render-SummaryWithDetails -Parent $contentView -SummaryLines $lines -DetailsText $Script:ADHealth.Data.GPO.Details
      }
    }
  }

  ## ----------------------------
  ## Button actions
  ## ----------------------------
  $btnRefreshAll.add_Click({
  ## Show a tiny busy modal while checks run
  $busy = [Terminal.Gui.Window]::new("Refreshing...")
  $busy.X = [Terminal.Gui.Pos]::Center(); $busy.Y = [Terminal.Gui.Pos]::Center()
  $busy.Width = 36; $busy.Height = 5
  $busy.Add([Terminal.Gui.Label]::new("Running AD health checks, please wait..."))
  [Terminal.Gui.Application]::Run($busy)
  try { Run-AllChecks -Domain $Script:CurrentDomain } finally { [Terminal.Gui.Application]::RequestStop() }
    Render-ADHealthTabs; Render-ADHealthContent -Tab $Script:ActiveHealthTab
  })

  $btnRefreshTab.add_Click({
  ## Run only the check for the active tab
  switch ($Script:ActiveHealthTab) {
    0 { Run-DCStatusCheck -Domain $Script:CurrentDomain }
    1 { Run-ReplicationCheck -Domain $Script:CurrentDomain }
    2 { Run-DNSCheck -Domain $Script:CurrentDomain }
    3 { Run-SYSVOLCheck -Domain $Script:CurrentDomain }
    4 { Run-FSMOCheck -Domain $Script:CurrentDomain }
    5 { Run-GPOCheck -Domain $Script:CurrentDomain }
  }
  Render-ADHealthTabs; Render-ADHealthContent -Tab $Script:ActiveHealthTab
  })

  ## Export (plain-text)
  $btnExport.add_Click({
  $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
  $fname = "ADHealth_${($Script:CurrentDomain -replace '[^a-zA-Z0-9\.-]','_')}_$timestamp.txt"
  $full = Join-Path -Path (Get-Location) -ChildPath $fname
  $sb = New-Object System.Text.StringBuilder
  $sb.AppendLine("AD Health Report for $($Script:CurrentDomain) - $timestamp") | Out-Null
  foreach ($tabName in $tabs) {
    $sb.AppendLine("---- $tabName ----") | Out-Null
    switch ($tabName) {
      "DC Status"   { $sb.AppendLine($Script:ADHealth.Data.DCStatus.Details) | Out-Null }
      "Replication" { $sb.AppendLine($Script:ADHealth.Data.Replication.Details) | Out-Null }
      "DNS"         { $sb.AppendLine($Script:ADHealth.Data.DNS.Details) | Out-Null }
      "SYSVOL"      { $sb.AppendLine($Script:ADHealth.Data.SYSVOL.Details) | Out-Null }
      "FSMO Roles"  { $sb.AppendLine($Script:ADHealth.Data.FSMO.Details) | Out-Null }
      "GPO Health"  { $sb.AppendLine($Script:ADHealth.Data.GPO.Details) | Out-Null }
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

  ## Add Escape key handler to close the modal (conventional)
  $modal.add_KeyDown({
    param($args)
      try {
        if ($args.KeyEvent.Key -eq [Terminal.Gui.Key]::Escape) {
          [Terminal.Gui.Application]::RequestStop()
          $args.Handled = $true
        }
        ## Alt+Left / Alt+Right support: move tabs
        if ($args.KeyEvent.IsAlt && $args.KeyEvent.Key -eq [Terminal.Gui.Key]::LeftArrow) {
          $Script:ActiveHealthTab = [Math]::Max(0, $Script:ActiveHealthTab - 1)
          Render-ADHealthTabs; Render-ADHealthContent -Tab $Script:ActiveHealthTab
          $args.Handled = $true
        } elseif ($args.KeyEvent.IsAlt && $args.KeyEvent.Key -eq [Terminal.Gui.Key]::RightArrow) {
          $Script:ActiveHealthTab = [Math]::Min($tabs.Count - 1, $Script:ActiveHealthTab + 1)
          Render-ADHealthTabs; Render-ADHealthContent -Tab $Script:ActiveHealthTab
          $args.Handled = $true
        }
      } catch { }
    })

    ## Also allow Alt+E to activate Details by simulating a click on the Details button in the content area
    ## Note: Details button is created per content render; we'll also respond to Alt+E globally to open Details for current tab.
    $modal.add_KeyDown({
      param($args)
        try {
          if ($args.KeyEvent.IsAlt -and $args.KeyEvent.Key -eq [Terminal.Gui.Key]::E) {
            ## Determine details for current tab and show modal if present
            switch ($Script:ActiveHealthTab) {
              0 { $d = $Script:ADHealth.Data.DCStatus.Details }
              1 { $d = $Script:ADHealth.Data.Replication.Details }
              2 { $d = $Script:ADHealth.Data.DNS.Details }
              3 { $d = $Script:ADHealth.Data.SYSVOL.Details }
              4 { $d = $Script:ADHealth.Data.FSMO.Details }
              5 { $d = $Script:ADHealth.Data.GPO.Details }
            }
            if ($d -and $d.Length -gt 0) { Show-DetailsModal -Title "Details" -Content $d } else { Show-Modal "Details" "No details available for this tab." }
              $args.Handled = $true
          }
        } catch { }
    })

    ## ----------------------------
    ## Initial run & render
    ## ----------------------------
    Run-AllChecks -Domain $Script:CurrentDomain
    Render-ToolsRow
    Render-ADHealthTabs
    Render-ADHealthContent -Tab 0

    ## Run the modal
    [Terminal.Gui.Application]::Run($modal)
}

## Generate Random Password modal
function Generate-RandomPassword {
  ## Helper function
  function Get-PasswordEntropy {
    param([int]$poolSize, [int]$length)
    if ($poolSize -le 0 -or $length -le 0) { return 0 }
    return [Math]::Log($poolSize, 2) * $length
  }

  $UpperCase = @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z')
  $LowerCase = @('a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z')
  $Numbers   = @('1','2','3','4','5','6','7','8','9','0')
  $Symbols   = @('!','@','$','?','<','>','*','&')
  $Script:actualPassword = ""

  ## --- Generate-RandomPassword UI ---
  $dlg = [Terminal.Gui.Dialog]::new("Generate Random Password", 66, 14)

  ## 2x2 checkbox layout
  $chkUpper = [Terminal.Gui.CheckBox]::new(2,1,"Include Uppercase (A-Z)", $true)
  $chkLower = [Terminal.Gui.CheckBox]::new(30,1,"Include Lowercase (a-z)", $true)
  $chkNums  = [Terminal.Gui.CheckBox]::new(2,3,"Include Numbers (0-9)", $true)
  $chkSyms  = [Terminal.Gui.CheckBox]::new(30,3,"Include Symbols (!,@,$)", $true)
  $dlg.Add($chkUpper, $chkLower, $chkNums, $chkSyms)

  ## Length input
  $dlg.Add([Terminal.Gui.Label]::new(2,5,"Length (1-127):"))
  $txtLen = New-Object Terminal.Gui.TextField
  $txtLen.X = 18; $txtLen.Y = 5; $txtLen.Width = 6
  $txtLen.Text = "12"
  $dlg.Add($txtLen)

  ## Password display box
  $dlg.Add([Terminal.Gui.Label]::new(2,7,"Generated Password:"))
  $txtPwd = New-Object Terminal.Gui.TextField
  $txtPwd.X = 2; $txtPwd.Y = 8; $txtPwd.Width = 25
  $txtPwd.Text = ""
  $dlg.Add($txtPwd)

  ## --- ENTROPY DISPLAY ---
  $lblStrength = [Terminal.Gui.Label]::new(30,7,"Strength: Not generated")
  $lblStrength.Width = 30
  $dlg.Add($lblStrength)
  $progEntropy = [Terminal.Gui.ProgressBar]::new()
  $progEntropy.X = 30; $progEntropy.Y = 8; $progEntropy.Width = 25
  $progEntropy.Fraction = 0.0
  $progEntropy = [Terminal.Gui.ProgressBar]::new()
  $progEntropy.X = 30; $progEntropy.Y = 8; $progEntropy.Width = 25
  $progEntropy.Fraction = 0.0

  ## Apply theme colors to progress bar
  if ($Script:ThemeMode) {
    $themeData = Get-Theme -mode $Script:ThemeMode
    if ($themeData -and $themeData.MainWindow) {
      ## Create a custom ColorScheme for the progress bar
      $progColorScheme = [Terminal.Gui.ColorScheme]::new()

      ## Use MainWindow colors for the background (matches dialog background)
      $progColorScheme.Normal = [Terminal.Gui.Attribute]::Make(
        $themeData.MainWindow.Normal.Foreground,
        $themeData.MainWindow.Normal.Background
      )

      $progColorScheme.Focus = [Terminal.Gui.Attribute]::Make(
        $themeData.MainWindow.Normal.Foreground,
        $themeData.MainWindow.Normal.Background
      )

      ## The filled bar - use MainWindow Focus colors to stand out
      $progColorScheme.HotNormal = [Terminal.Gui.Attribute]::Make(
        $themeData.MainWindow.Focus.Foreground,
        $themeData.MainWindow.Focus.Background
      )

      $progColorScheme.HotFocus = [Terminal.Gui.Attribute]::Make(
        $themeData.MainWindow.Focus.Foreground,
        $themeData.MainWindow.Focus.Background
      )

      $progEntropy.ColorScheme = $progColorScheme
    }
  }

  $dlg.Add($progEntropy)

  ## Show Password checkbox
  $chkShowPwd = [Terminal.Gui.CheckBox]::new(30,5,"Show Password",$false)
  $dlg.Add($chkShowPwd)

  ## Buttons
  $btnGenerate = [Terminal.Gui.Button]::new("Generate"); $btnGenerate.X=2; $btnGenerate.Y=10
  $btnCopy     = [Terminal.Gui.Button]::new("Copy");     $btnCopy.X=15; $btnCopy.Y=10
  $btnClose    = [Terminal.Gui.Button]::new("Close");    $btnClose.X=28; $btnClose.Y=10
  $dlg.Add($btnGenerate, $btnCopy, $btnClose)

  ## --- Generate Logic ---
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

      $Script:actualPassword = -join (1..$len | ForEach-Object { $pool | Get-Random })

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

        if ($chkShowPwd.Checked) {
            $txtPwd.Text = $Script:actualPassword
        } else {
            $txtPwd.Text = ('*' * $Script:actualPassword.Length)
        }
    })

    # --- Show Password toggle ---
    $chkShowPwd.add_Toggled({
        if ($chkShowPwd.Checked) {
            $txtPwd.Text = $Script:actualPassword
        } else {
            $txtPwd.Text = ('*' * $Script:actualPassword.Length)
        }
    })

    # --- Copy to Clipboard ---
    $btnCopy.add_Clicked({
        if (-not $Script:actualPassword) { return }
        if ($IsWindows) { Set-Clipboard -Value $Script:actualPassword }
        elseif ($IsMacOS) { $Script:actualPassword | pbcopy }
        else { $Script:actualPassword | xsel --clipboard --input }
        Show-Modal "Copied" "Password copied to clipboard."
    })

    # --- Close ---
    $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })

    [Terminal.Gui.Application]::Run($dlg)
    return $Script:actualPassword
}

## Show LAPS Passwords
function Show-LAPSSearchModal {
    <#
        Modal LAPS lookup UI
        - Browse all LAPS-enabled computers
        - Search by computer name
        - Masked password with press-to-reveal
        - Copy-to-clipboard (timeout currently DISABLED)
        - Expiry warning highlighting
        - DemoMode support

        NOTE ON CLIPBOARD TIMEOUT:
        --------------------------
        Clipboard auto-clear is intentionally DISABLED for now.
        To re-enable later, search for the section marked:
            ### CLIPBOARD TIMEOUT (DISABLED) ###
        and uncomment the timer block.
    #>

    # ---------------- Safety: AD availability ----------------
    if (-not $Global:DemoMode) {
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            Show-Modal -Title "LAPS Error" -Message "ActiveDirectory module not available." -Buttons @("OK")
            return
        }
        Import-Module ActiveDirectory -ErrorAction Stop
    }

    # ---------------- Modal ----------------
    $dialog = [Terminal.Gui.Dialog]::new("LAPS Password Lookup", 90, 26)

    $lblSearch = [Terminal.Gui.Label]::new(1,1,"Computer name (blank = all):")
    $dialog.Add($lblSearch)

    $txtSearch = [Terminal.Gui.TextField]::new("")
    $txtSearch.X = 1
    $txtSearch.Y = 2
    $txtSearch.Width = 40
    $dialog.Add($txtSearch)

    $lstComputers = [Terminal.Gui.ListView]::new()
    $lstComputers.X = 1
    $lstComputers.Y = 4
    $lstComputers.Width = 86
    $lstComputers.Height = 13
    $dialog.Add($lstComputers)

    $computers = @()

    # ---------------- Data loader ----------------
    $loadComputers = {
        param($filter)

        try {
            if ($Global:DemoMode) {
                # ---- Demo stub ----
                $computers = @(
                    [pscustomobject]@{
                        Name = 'DEMO-PC-01'
                        'msLAPS-AccountName' = 'Administrator'
                        'msLAPS-Password' = 'DemoPassword!123'
                        'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(5).ToFileTimeUtc()
                    }
                )
            }
            else {
                if ([string]::IsNullOrWhiteSpace($filter)) {
                    $computers = Get-ADComputer -Filter * -Properties msLAPS-Password,msLAPS-AccountName,msLAPS-PasswordExpirationTime |
                        Where-Object { $_.'msLAPS-Password' }
                }
                else {
                    $computers = Get-ADComputer -Filter "Name -like '*$filter*'" -Properties msLAPS-Password,msLAPS-AccountName,msLAPS-PasswordExpirationTime |
                        Where-Object { $_.'msLAPS-Password' }
                }
            }

            $lstComputers.SetSource($computers.Name)
        }
        catch {
            Show-Modal -Title "Error" -Message $_.Exception.Message -Buttons @("OK")
        }
    }

    & $loadComputers ""

    # ---------------- Buttons ----------------
    $btnSearch = [Terminal.Gui.Button]::new("Search")
    $btnSearch.X = 45
    $btnSearch.Y = 2
    $btnSearch.Add_Clicked({ & $loadComputers $txtSearch.Text.ToString() })
    $dialog.Add($btnSearch)

    $btnView = [Terminal.Gui.Button]::new("View")
    $btnView.X = 60
    $btnView.Y = 18
    $dialog.Add($btnView)

    $btnClose = [Terminal.Gui.Button]::new("Close")
    $btnClose.X = 72
    $btnClose.Y = 18
    $btnClose.Add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dialog.Add($btnClose)

    # ---------------- View handler ----------------
    $btnView.Add_Clicked({
        if ($lstComputers.SelectedItem -lt 0) { return }

        $computer = $computers[$lstComputers.SelectedItem]
        $expires = [DateTime]::FromFileTimeUtc($computer.'msLAPS-PasswordExpirationTime')
        $daysLeft = ($expires - (Get-Date)).Days

        # ---- Mask password ----
        $masked = ('*' * ($computer.'msLAPS-Password'.Length))

        $warn = if ($daysLeft -le 3) { "WARNING: Password expires in ${daysLeft} day(s)" } else { "Expires in ${daysLeft} days" }

        $msg = @"
Computer : ${computer.Name}

LAPS User:
${computer.'msLAPS-AccountName'}

LAPS Password (masked):
${masked}

${warn}

[Press Reveal to show password]
"@

        $detailDlg = [Terminal.Gui.Dialog]::new("LAPS Details", 70, 20)

        $lbl = [Terminal.Gui.Label]::new(1,1,$msg)
        $detailDlg.Add($lbl)

        $btnReveal = [Terminal.Gui.Button]::new("Reveal")
        $btnReveal.X = 1
        $btnReveal.Y = 14
        $detailDlg.Add($btnReveal)

        $btnCopy = [Terminal.Gui.Button]::new("Copy Password")
        $btnCopy.X = 12
        $btnCopy.Y = 14
        $detailDlg.Add($btnCopy)

        $btnOk = [Terminal.Gui.Button]::new("OK")
        $btnOk.X = 30
        $btnOk.Y = 14
        $btnOk.Add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
        $detailDlg.Add($btnOk)

        $btnReveal.Add_Clicked({
            $lbl.Text = $lbl.Text + "\n\nUNMASKED PASSWORD:\n${computer.'msLAPS-Password'}"
        })

        $btnCopy.Add_Clicked({
            Set-Clipboard -Value $computer.'msLAPS-Password'

            ### CLIPBOARD TIMEOUT (DISABLED) ###
            # To re-enable later:
            # Start-Job { Start-Sleep 15; Set-Clipboard -Value "" }

            Show-Modal -Title "Clipboard" -Message "Password copied to clipboard." -Buttons @("OK")
        })

        [Terminal.Gui.Application]::Run($detailDlg)
    })

    [Terminal.Gui.Application]::Run($dialog)
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

## First, create the Apply-UserChanges function (put this before Show-UserPropertiesDialog):

function Apply-UserChanges {
    param($user, $fields)

    Debug-Log (": Applying changes for user: $($user.Name)") -Type "Info"

    try {
        if ($Script:DemoMode) {
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

            Debug-Log ("SUCCESS: User changes applied (demo mode)") -Type "Info"
            Show-Modal "Success" "Changes applied successfully (demo mode)"

            # Rebuild tree to reflect changes
            [Terminal.Gui.Application]::MainLoop.Invoke({
                Build-Tree -domain $Script:CurrentDomain
                if ($filterStatusLabel) {
                    Update-FilterStatusLabel -label $Script:FilterStatusLabel
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

            Debug-Log ("SUCCESS: User changes applied to AD") -Type "Info"
            Show-Modal "Success" "Changes applied successfully"

            # Reload AD data
Refresh-Data -domain $Script:CurrentDomain
            if ($filterStatusLabel) {
                Update-FilterStatusLabel -label $Script:FilterStatusLabel
            }
        }

        $Script:changesMade = $false
        return $true

    } catch {
        Debug-Log (": Failed to apply changes: $($_.Exception.Message)") -Type "Warn"
        Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)"
        return $false
    }
}

# =====================================================
# Global Helper Functions for Multi-Row Tab System
# =====================================================

# ---------------------- Select-TabGlobal Function ----------------------
function Select-TabGlobal {
    param(
        $tab,
        $contentHost
    )

    try {
        if ($Script:IsInitializing) {
            Debug-Log (": Skipping tab selection - still initializing") -Type "Info"
            return
        }

        $tabText = if ($tab -and $tab.Text) { [System.Text.Encoding]::UTF8.GetString($tab.Text) } else { "null" }
        Debug-Log (": Select-TabGlobal called for tab: $tabText") -Type "Info"

        if (-not $tab) {
            Debug-Log (": Tab is null, aborting") -Type "Warn"
            return
        }

        if (-not $contentHost) {
            Debug-Log (": ContentHost is null, aborting") -Type "Warn"
            return
        }

        # Hide all tab views
        if ($Script:AllTabs) {
            foreach ($t in $Script:AllTabs) {
                if ($t -and $t.View) {
                    $t.View.Visible = $false
                }
            }
        }

        # Show selected tab view
        if ($tab.View) {
            $tab.View.Visible = $true
            $Script:ActiveTab = $tab
            Debug-Log (": Tab view set to visible for: $tabText") -Type "Info"
        } else {
            Debug-Log (": WARNING - Tab has no View property") -Type "Warn"
            return
        }

        # Update visual selection in both rows
        if ($Script:TabRows) {
            foreach ($row in $Script:TabRows) {
                if ($row) {
                    $row.SelectedTab = $tab
                }
            }
            Debug-Log (": Updated row selections") -Type "Info"
        }

        # Clear and add the selected tab's view to content host
        $contentHost.RemoveAll()
        $contentHost.Add($tab.View)
        $contentHost.SetNeedsDisplay()
        Debug-Log (": Added tab view to content host for: $tabText") -Type "Success"
    } catch {
        Debug-Log (": ERROR in Select-TabGlobal: $($_.Exception.Message)") -Type "Error"
        throw
    }
}

# ---------------------- Register-Tab Function ----------------------
function Register-Tab {
    param(
        [Terminal.Gui.TabView+Tab]$Tab,
        [int]$Row = 0,
        $contentHost
    )

    try {
        # Convert tab text to string properly
        $tabText = if ($Tab.Text -is [byte[]]) {
            [System.Text.Encoding]::UTF8.GetString($Tab.Text)
        } elseif ($Tab.Text) {
            $Tab.Text.ToString()
        } else {
            "Unknown"
        }

        Debug-Log (": Register-Tab starting for tab: $tabText, Row: $Row") -Type "Info"

        # Pick which row to add the tab to
        $targetRow = $Script:TabRows[$Row]

        if ($null -eq $targetRow) {
            throw "Target row $Row is null"
        }

        Debug-Log (": Target row found, adding tab") -Type "Info"

        # Use AddTab method instead of Tabs.Add
        $targetRow.AddTab($Tab, $false)  # false = don't select automatically

        Debug-Log (": Tab added to row, updating global list") -Type "Info"

        # Add to global tab list for navigation
        $Script:AllTabs += $Tab

        Debug-Log (": Tab added to global list, checking if first tab") -Type "Info"

        # Don't auto-select during initialization - we'll do it manually after all tabs are registered
        if (-not $Script:IsInitializing -and -not $Script:ActiveTab) {
            Debug-Log (": First tab after initialization, selecting it") -Type "Info"
            Select-TabGlobal -tab $Tab -contentHost $contentHost
        }

        Debug-Log (": Register-Tab completed for tab: $tabText") -Type "Success"
    } catch {
        Debug-Log (": ERROR in Register-Tab: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)") -Type "Error"
        throw
    }
}

# =====================================================
# Main Dialog Function
# =====================================================

function Show-UserPropertiesDialog {
    param($user)

    # ---------------------- Safety Checks ----------------------
    if (-not $user) {
        Debug-Log ": User object is null" -Type "Warn"
        return
    }
    Debug-Log ": Show-UserPropertiesDialog starting for: $($user.Name)" -Type "Info"

    try {
        # ---------------------- Buttons ----------------------
        $btnOK = [Terminal.Gui.Button]::new("OK")
        $btnCancel = [Terminal.Gui.Button]::new("Cancel")
        $btnApply = [Terminal.Gui.Button]::new("Apply")

        # Button handlers
        $btnOK.add_Clicked({
            Debug-Log ": OK clicked" -Type "Info"
            [Terminal.Gui.Application]::RequestStop()
        })

        $btnCancel.add_Clicked({
            Debug-Log ": Cancel clicked" -Type "Info"
            [Terminal.Gui.Application]::RequestStop()
        })

        $btnApply.add_Clicked({
            Debug-Log ": Apply clicked - changes would be saved here" -Type "Info"
        })

        # ---------------------- Dialog ----------------------
        $dlg = [Terminal.Gui.Dialog]::new("User Properties - $($user.Name)", 100, 40, $btnOK, $btnCancel, $btnApply)

        # ---------------------- Standard TabView ----------------------
        $tabView = [Terminal.Gui.TabView]::new()
        $tabView.X = 0
        $tabView.Y = 0
        $tabView.Width = [Terminal.Gui.Dim]::Fill()
        $tabView.Height = [Terminal.Gui.Dim]::Fill(1)  # Leave room for buttons

        # ==================== General Tab ====================
        Debug-Log ": Creating General tab" -Type "Info"
        $generalTab = [Terminal.Gui.TabView+Tab]::new()
        $generalTab.Text = "General"
        $generalView = [Terminal.Gui.View]::new()
        $generalView.X = 0; $generalView.Y = 0
        $generalView.Width = [Terminal.Gui.Dim]::Fill()
        $generalView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        # Basic Info
        $lbl = [Terminal.Gui.Label]::new("User Information"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Name:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $txtName = [Terminal.Gui.Label]::new($user.Name ?? ""); $txtName.X=20; $txtName.Y=$y
        $generalView.Add($txtName); $y+=1

        $lbl = [Terminal.Gui.Label]::new("Display Name:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $txtDisplayName = [Terminal.Gui.TextField]::new($user.DisplayName ?? ""); $txtDisplayName.X=20; $txtDisplayName.Y=$y; $txtDisplayName.Width=60
        $generalView.Add($txtDisplayName); $y+=1

        $lbl = [Terminal.Gui.Label]::new("Email:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $emailAddr = if ($user.EmailAddress) { $user.EmailAddress } elseif ($user.mail) { $user.mail } else { "" }
        $txtEmail = [Terminal.Gui.TextField]::new($emailAddr); $txtEmail.X=20; $txtEmail.Y=$y; $txtEmail.Width=60
        $generalView.Add($txtEmail); $y+=1

        $lbl = [Terminal.Gui.Label]::new("Description:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $txtDescription = [Terminal.Gui.TextField]::new($user.Description ?? ""); $txtDescription.X=20; $txtDescription.Y=$y; $txtDescription.Width=60
        $generalView.Add($txtDescription); $y+=2

        # Phone numbers
        $lbl = [Terminal.Gui.Label]::new("Contact Information"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Office:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $txtOfficePhone = [Terminal.Gui.TextField]::new($user.OfficePhone ?? ""); $txtOfficePhone.X=20; $txtOfficePhone.Y=$y; $txtOfficePhone.Width=30
        $generalView.Add($txtOfficePhone); $y+=1

        $lbl = [Terminal.Gui.Label]::new("Mobile:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $txtMobilePhone = [Terminal.Gui.TextField]::new($user.MobilePhone ?? ""); $txtMobilePhone.X=20; $txtMobilePhone.Y=$y; $txtMobilePhone.Width=30
        $generalView.Add($txtMobilePhone); $y+=1

        $generalTab.View = $generalView
        $tabView.AddTab($generalTab, $false)

        # ==================== Account Tab ====================
        Debug-Log ": Creating Account tab" -Type "Info"
        $accountTab = [Terminal.Gui.TabView+Tab]::new()
        $accountTab.Text = "Account"
        $accountView = [Terminal.Gui.View]::new()
        $accountView.X = 0; $accountView.Y = 0
        $accountView.Width = [Terminal.Gui.Dim]::Fill()
        $accountView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        # Account Status
        $lbl = [Terminal.Gui.Label]::new("Account Status"); $lbl.X=2; $lbl.Y=$y; $accountView.Add($lbl); $y+=2

        $isEnabled = if ($user.PSObject.Properties['Enabled']) { $user.Enabled } else { -not $user.Disabled }
        $chkEnabled = [Terminal.Gui.CheckBox]::new("Account Enabled"); $chkEnabled.X=4; $chkEnabled.Y=$y; $chkEnabled.Checked=$isEnabled
        $accountView.Add($chkEnabled); $y+=1

        $isLocked = if ($user.PSObject.Properties['LockedOut']) { $user.LockedOut } else { $user.Locked ?? $false }
        $chkLocked = [Terminal.Gui.CheckBox]::new("Account Locked"); $chkLocked.X=4; $chkLocked.Y=$y; $chkLocked.Checked=$isLocked; $chkLocked.Enabled=$false
        $accountView.Add($chkLocked); $y+=2

        # Password Settings
        $lbl = [Terminal.Gui.Label]::new("Password Settings"); $lbl.X=2; $lbl.Y=$y; $accountView.Add($lbl); $y+=2
        $chkPasswordExpired = [Terminal.Gui.CheckBox]::new("Password Expired"); $chkPasswordExpired.X=4; $chkPasswordExpired.Y=$y; $chkPasswordExpired.Checked=($user.PasswordExpired??$false); $chkPasswordExpired.Enabled=$false
        $accountView.Add($chkPasswordExpired); $y+=1

        $chkMustChangePassword = [Terminal.Gui.CheckBox]::new("User must change password at next logon"); $chkMustChangePassword.X=4; $chkMustChangePassword.Y=$y
        $chkMustChangePassword.Checked = if ($user.PasswordNeverExpires){$false}else{ if ($user.PSObject.Properties['pwdLastSet']){ $user.pwdLastSet -eq 0 } else { $false } }
        $accountView.Add($chkMustChangePassword); $y+=1

        $chkCannotChangePassword = [Terminal.Gui.CheckBox]::new("User cannot change password"); $chkCannotChangePassword.X=4; $chkCannotChangePassword.Y=$y; $chkCannotChangePassword.Checked=($user.CannotChangePassword??$false)
        $accountView.Add($chkCannotChangePassword); $y+=1

        $chkPasswordNeverExpires = [Terminal.Gui.CheckBox]::new("Password never expires"); $chkPasswordNeverExpires.X=4; $chkPasswordNeverExpires.Y=$y; $chkPasswordNeverExpires.Checked=($user.PasswordNeverExpires??$false)
        $accountView.Add($chkPasswordNeverExpires); $y+=2

        # Logon Information
        $lbl = [Terminal.Gui.Label]::new("Logon Information"); $lbl.X=2; $lbl.Y=$y; $accountView.Add($lbl); $y+=2
        $lbl = [Terminal.Gui.Label]::new("Last logon: "+($user.LastLogonDate?.ToString('yyyy-MM-dd HH:mm') ?? 'Never')); $lbl.X=4; $lbl.Y=$y
        $accountView.Add($lbl); $y+=1

        if ($user.PSObject.Properties['PasswordLastSet'] -and $user.PasswordLastSet) {
            $lbl = [Terminal.Gui.Label]::new("Password last set: "+$user.PasswordLastSet.ToString('yyyy-MM-dd HH:mm')); $lbl.X=4; $lbl.Y=$y
            $accountView.Add($lbl); $y+=1
        }

        if ($user.PSObject.Properties['LogonCount'] -or $user.PSObject.Properties['logonCount']) {
            $logonCount = if ($user.LogonCount) { $user.LogonCount } else { $user.logonCount }
            $lbl = [Terminal.Gui.Label]::new("Logon count: $logonCount"); $lbl.X=4; $lbl.Y=$y
            $accountView.Add($lbl); $y+=1
        }

        $accountTab.View = $accountView
        $tabView.AddTab($accountTab, $false)

        # ==================== Address Tab ====================
        Debug-Log ": Creating Address tab" -Type "Info"
        $addressTab = [Terminal.Gui.TabView+Tab]::new()
        $addressTab.Text = "Address"
        $addressView = [Terminal.Gui.View]::new()
        $addressView.X = 0; $addressView.Y = 0
        $addressView.Width = [Terminal.Gui.Dim]::Fill()
        $addressView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        $lbl = [Terminal.Gui.Label]::new("Street:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
        $txtStreet = [Terminal.Gui.TextField]::new($user.StreetAddress ?? ""); $txtStreet.X=20; $txtStreet.Y=$y; $txtStreet.Width=70
        $addressView.Add($txtStreet); $y+=2

        $lbl = [Terminal.Gui.Label]::new("City:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
        $txtCity = [Terminal.Gui.TextField]::new($user.City ?? ""); $txtCity.X=20; $txtCity.Y=$y; $txtCity.Width=70
        $addressView.Add($txtCity); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Postal Code:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
        $txtPostal = [Terminal.Gui.TextField]::new($user.PostalCode ?? ""); $txtPostal.X=20; $txtPostal.Y=$y; $txtPostal.Width=20
        $addressView.Add($txtPostal); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Country:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
        $txtCountry = [Terminal.Gui.TextField]::new($user.Country ?? ""); $txtCountry.X=20; $txtCountry.Y=$y; $txtCountry.Width=70
        $addressView.Add($txtCountry); $y+=1

        $addressTab.View = $addressView
        $tabView.AddTab($addressTab, $false)

        # ==================== Profile Tab ====================
        Debug-Log ": Creating Profile tab" -Type "Info"
        $profileTab = [Terminal.Gui.TabView+Tab]::new()
        $profileTab.Text = "Profile"
        $profileView = [Terminal.Gui.View]::new()
        $profileView.X = 0; $profileView.Y = 0
        $profileView.Width = [Terminal.Gui.Dim]::Fill()
        $profileView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        $lbl = [Terminal.Gui.Label]::new("User Profile"); $lbl.X=2; $lbl.Y=$y; $profileView.Add($lbl); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Profile path:"); $lbl.X=4; $lbl.Y=$y; $profileView.Add($lbl)
        $txtProfilePath = [Terminal.Gui.TextField]::new($user.ProfilePath ?? ""); $txtProfilePath.X=20; $txtProfilePath.Y=$y; $txtProfilePath.Width=70
        $profileView.Add($txtProfilePath); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Logon script:"); $lbl.X=4; $lbl.Y=$y; $profileView.Add($lbl)
        $txtLogonScript = [Terminal.Gui.TextField]::new($user.ScriptPath ?? ""); $txtLogonScript.X=20; $txtLogonScript.Y=$y; $txtLogonScript.Width=70
        $profileView.Add($txtLogonScript); $y+=3

        # Home Folder
        $lbl = [Terminal.Gui.Label]::new("Home Folder"); $lbl.X=2; $lbl.Y=$y; $profileView.Add($lbl); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Home directory:"); $lbl.X=4; $lbl.Y=$y; $profileView.Add($lbl)
        $txtHomeDirectory = [Terminal.Gui.TextField]::new($user.HomeDirectory ?? ""); $txtHomeDirectory.X=20; $txtHomeDirectory.Y=$y; $txtHomeDirectory.Width=70
        $profileView.Add($txtHomeDirectory); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Home drive:"); $lbl.X=4; $lbl.Y=$y; $profileView.Add($lbl)
        $txtHomeDrive = [Terminal.Gui.TextField]::new($user.HomeDrive ?? ""); $txtHomeDrive.X=20; $txtHomeDrive.Y=$y; $txtHomeDrive.Width=5
        $profileView.Add($txtHomeDrive); $y+=1

        $profileTab.View = $profileView
        $tabView.AddTab($profileTab, $false)

        # ==================== Organization Tab ====================
        Debug-Log ": Creating Organization tab" -Type "Info"
        $orgTab = [Terminal.Gui.TabView+Tab]::new()
        $orgTab.Text = "Organization"
        $orgView = [Terminal.Gui.View]::new()
        $orgView.X = 0; $orgView.Y = 0
        $orgView.Width = [Terminal.Gui.Dim]::Fill()
        $orgView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        $lbl = [Terminal.Gui.Label]::new("Title:"); $lbl.X=2; $lbl.Y=$y; $orgView.Add($lbl)
        $txtTitle = [Terminal.Gui.TextField]::new($user.Title ?? ""); $txtTitle.X=20; $txtTitle.Y=$y; $txtTitle.Width=70
        $orgView.Add($txtTitle); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Department:"); $lbl.X=2; $lbl.Y=$y; $orgView.Add($lbl)
        $txtDept = [Terminal.Gui.TextField]::new($user.Department ?? ""); $txtDept.X=20; $txtDept.Y=$y; $txtDept.Width=70
        $orgView.Add($txtDept); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Company:"); $lbl.X=2; $lbl.Y=$y; $orgView.Add($lbl)
        $txtCompany = [Terminal.Gui.TextField]::new($user.Company ?? ""); $txtCompany.X=20; $txtCompany.Y=$y; $txtCompany.Width=70
        $orgView.Add($txtCompany); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Manager:"); $lbl.X=2; $lbl.Y=$y; $orgView.Add($lbl)
        $txtManager = [Terminal.Gui.TextField]::new($user.Manager ?? ""); $txtManager.X=20; $txtManager.Y=$y; $txtManager.Width=70
        $orgView.Add($txtManager); $y+=1

        $orgTab.View = $orgView
        $tabView.AddTab($orgTab, $false)

# ==================== Member Of Tab ====================
Debug-Log ": Creating Member Of tab" -Type "Info"
$memberTab = [Terminal.Gui.TabView+Tab]::new()
$memberTab.Text = "Member Of"
$memberView = [Terminal.Gui.View]::new()
$memberView.X = 0; $memberView.Y = 0
$memberView.Width = [Terminal.Gui.Dim]::Fill()
$memberView.Height = [Terminal.Gui.Dim]::Fill(3)

$y = 1
$lbl = [Terminal.Gui.Label]::new("Group Memberships:"); $lbl.X=2; $lbl.Y=$y; $memberView.Add($lbl); $y+=2

# Create ListView for groups
$lstGroups = [Terminal.Gui.ListView]::new()
$lstGroups.X = 2
$lstGroups.Y = $y
$lstGroups.Width = [Terminal.Gui.Dim]::Fill(2)
$lstGroups.Height = 25  # Fixed height

# Get group list
$groupList = @()
if ($user.Groups) {
    $groupList = $user.Groups
} elseif ($user.MemberOf) {
    # Extract group names from DNs
    $groupList = $user.MemberOf | ForEach-Object {
        if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
    }
}

if ($groupList.Count -gt 0) {
    $lstGroups.SetSource($groupList)
} else {
    $lstGroups.SetSource(@("(No group memberships)"))
}

$memberView.Add($lstGroups)

# Add to Group button
$btnAdd = [Terminal.Gui.Button]::new("Add to Group...")
$btnAdd.X = 2
$btnAdd.Y = 28  # y(1) + label(1) + list(25) + gap(1)
$btnAdd.add_Clicked({
    Show-EditGroupMembershipDialog -User $user -OnUpdate {
        Debug-Log ": Refreshing group list after add" -Type "Info"

        # Re-extract the group list from the updated user object
        $refreshedGroups = @()
        if ($user.Groups) {
            $refreshedGroups = $user.Groups
        } elseif ($user.MemberOf) {
            # Extract group names from DNs
            $refreshedGroups = $user.MemberOf | ForEach-Object {
                if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
            }
        }

        # Update the ListView
        if ($refreshedGroups.Count -gt 0) {
            $lstGroups.SetSource($refreshedGroups)
        } else {
            $lstGroups.SetSource(@("(No group memberships)"))
        }

        # Also update the $groupList variable used by Remove button
        $groupList = $refreshedGroups
        Debug-Log ": Group list refreshed, now showing $($refreshedGroups.Count) groups" -Type "Success"
    }
})
$memberView.Add($btnAdd)

# Remove from Group button
$btnRemove = [Terminal.Gui.Button]::new("Remove from Group")
$btnRemove.X = 22
$btnRemove.Y = 28
$btnRemove.add_Clicked({
    $selectedIndex = $lstGroups.SelectedItem

    # Get current group list
    $currentGroups = @()
    if ($user.Groups) {
        $currentGroups = $user.Groups
    } elseif ($user.MemberOf) {
        $currentGroups = $user.MemberOf | ForEach-Object {
            if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
        }
    }

    if ($selectedIndex -ge 0 -and $selectedIndex -lt $currentGroups.Count) {
        $selectedGroup = $currentGroups[$selectedIndex]

        # Don't allow removal from "(No group memberships)" placeholder
        if ($selectedGroup -eq "(No group memberships)") {
            Show-Modal "Info" "No group selected"
            return
        }

        ## Confirm removal
        $confirmDlg = [Terminal.Gui.MessageBox]::Query(60, 10, "Confirm Removal", "Remove $($user.Name) from group '$selectedGroup'?",  "Yes", "No")

        if ($confirmDlg -eq 0) {  # Yes clicked
            try {
                if ($Script:DemoMode) {
                    # Demo mode - just update the list
                    $user.Groups = $user.Groups | Where-Object { $_ -ne $selectedGroup }
                    Debug-Log ": Removed $($user.Name) from group $selectedGroup (demo mode)" -Type "Success"
                } else {
                    # Production mode
                    Remove-ADGroupMember -Identity $selectedGroup -Members $user.SamAccountName -Confirm:$false
                    Debug-Log ": Removed $($user.Name) from group $selectedGroup" -Type "Success"
                }

                # Refresh the list
                $updatedGroups = @()
                if ($user.Groups) {
                    $updatedGroups = $user.Groups
                } elseif ($user.MemberOf) {
                    $updatedGroups = $user.MemberOf | Where-Object { $_ -ne $selectedGroup } | ForEach-Object {
                        if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
                    }
                }

                if ($updatedGroups.Count -gt 0) {
                    $lstGroups.SetSource($updatedGroups)
                } else {
                    $lstGroups.SetSource(@("(No group memberships)"))
                }

                $groupList = $updatedGroups

                Show-Modal "Success" "Successfully removed $($user.Name) from group '$selectedGroup'"

            } catch {
                Show-Modal "Error" "Failed to remove from group:`n$($_.Exception.Message)"
                Debug-Log ": Failed to remove from group: $($_.Exception.Message)" -Type "Error"
            }
        }
    } else {
        Show-Modal "Info" "Please select a group to remove"
    }
})
$memberView.Add($btnRemove)

$memberTab.View = $memberView
$tabView.AddTab($memberTab, $false)

# Add TabView to dialog
$dlg.Add($tabView)

Debug-Log ": All tabs added, running dialog" -Type "Success"

# Run the dialog
[Terminal.Gui.Application]::Run($dlg)
Debug-Log ": Dialog closed normally" -Type "Info"

} catch {
    Debug-Log ": Exception in Show-UserPropertiesDialog: $($_.Exception.Message)" -Type "Error"
    Debug-Log ": Stack trace: $($_.ScriptStackTrace)" -Type "Error"
    Show-Modal "Error" "Failed to display user properties:`n$($_.Exception.Message)"
}
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
    $ouList = if ($Script:DemoMode) {
        $Script:Users | Select-Object -ExpandProperty OU -Unique | Sort-Object
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
if ($Script:DemoMode) {
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
            $Script:rawUsers += $newUser

            # Reconvert to update $Script:Users with AD-like objects
            $converted = Convert-DataToADObjects -Users $Script:rawUsers -DCs $Script:rawDCs -Groups $Script:rawDemoGroups -Domain $Script:CurrentDomain
            $Script:Users = $converted.Users

            Debug-Log (": Created user $name in demo mode") -Type "Info"
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
            $Script:rawDemoGroups += $newGroup

            # Reconvert to update $Script:Groups with AD-like objects
            $converted = Convert-DataToADObjects -Users $Script:rawUsers -DCs $Script:rawDCs -Groups $Script:rawDemoGroups -Domain $Script:CurrentDomain
            $Script:Groups = $converted.Groups

            Debug-Log (": Created group $name in demo mode") -Type "Info"
        }

        "OrganizationalUnit" {
            # For OUs, we need to track them in a structure
            # OUs are built from the OU arrays in users, so we could either:
            # 1. Add a $Script:rawOUs array (cleaner)
            # 2. Just ensure the OU path exists when we rebuild the tree

            # For now, let's just ensure it's tracked
            if (-not $Script:rawOUs) {
                $Script:rawOUs = @()
            }

            $newOU = @{
                Name = $name
                Path = $ou
                Description = $displayName
            }

            $Script:rawOUs += $newOU

            Debug-Log (": Created OU $name in demo mode") -Type "Info"
        }

        "Computer" {
            # Similar to users, add to a computers array
            if (-not $Script:rawComputers) {
                $Script:rawComputers = @()
            }

            $newComputer = @{
                Name = $name
                OU = $ou
                Description = $displayName
            }

            $Script:rawComputers += $newComputer

            Debug-Log (": Created computer $name in demo mode") -Type "Info"
        }
    }

    Show-Modal "Success" "$objType '$name' created successfully (demo mode)"
    Build-Tree -domain $Script:CurrentDomain
    Update-FilterStatusLabel -label $Script:FilterStatusLabel
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
                            UserPrincipalName = "$sam@$($Script:CurrentDomain)"
                            AccountPassword = $secPwd
                            Enabled = $true
                            Path = $ou
                            ChangePasswordAtLogon = $true
                        }

                        if ($displayName) { $params['DisplayName'] = $displayName }
                        if ($email) { $params['EmailAddress'] = $email }

                        New-ADUser @params -ErrorAction Stop
                        Debug-Log (": Created user $name in AD") -Type "Info"
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
                        Debug-Log (": Created group $name in AD") -Type "Info"
                    }
                    "OrganizationalUnit" {
                        $params = @{
                            Name = $name
                            Path = $ou
                        }

                        if ($displayName) { $params['Description'] = $displayName }

                        New-ADOrganizationalUnit @params -ErrorAction Stop
                        Debug-Log (": Created OU $name in AD") -Type "Info"
                    }
                    "Computer" {
                        $params = @{
                            Name = $name
                            Path = $ou
                        }

                        New-ADComputer @params -ErrorAction Stop
                        Debug-Log (": Created computer $name in AD") -Type "Info"
                    }
                    "Contact" {
                        $params = @{
                            Name = $name
                            Type = "Contact"
                            Path = $ou
                        }

                        if ($displayName) { $params['DisplayName'] = $displayName }

                        New-ADObject @params -ErrorAction Stop
                        Debug-Log (": Created contact $name in AD") -Type "Info"
                    }
                }

                Show-Modal "Success" "$objType '$name' created successfully"

                # Refresh data
                Refresh-Data -domain $Script:CurrentDomain
                Update-FilterStatusLabel -label $Script:FilterStatusLabel
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

Debug-Log (": After removing prefix: '$cleanName'") -Type "Info"

    # Extra confirmation for destructive action
    $result = [Terminal.Gui.MessageBox]::Query(70, 11, "DELETE CONFIRMATION",
        "⚠️ WARNING: You are about to DELETE:`n`n  Type: $objectType`n  Name: $cleanName`n`nThis action CANNOT be undone!`n`nAre you absolutely sure?",
        "Yes, DELETE", "No, Cancel")

    if ($result -eq 0) {
        try {
            if ($Script:DemoMode) {
                # Demo mode - remove from in-memory structures
                switch ($objectType.ToLower()) {
                    "user" {
                        $Script:Users = $Script:Users | Where-Object { $_.Name -ne $cleanName }
                        Debug-Log (": Deleted user $cleanName (demo mode)") -Type "Info"
                    }
                    "group" {
                        # Remove group from all users
                        foreach ($u in $Script:Users) {
                            $u.Groups = $u.Groups | Where-Object { $_ -ne $cleanName }
                        }
                        Debug-Log (": Deleted group $cleanName (demo mode)") -Type "Info"
                    }
                    default {
                        Debug-Log (": Deleted $objectType $cleanName (demo mode)") -Type "Info"
                    }
                }

                Show-Modal "Deleted" "$objectType '$cleanName' deleted (demo mode)"
                Build-Tree -domain $Script:CurrentDomain
                Update-FilterStatusLabel -label $Script:FilterStatusLabel

            } else {
                # Production mode - delete from AD
                switch ($objectType.ToLower()) {
                    "user" {
                        Remove-ADUser -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log (": Deleted user $cleanName from AD") -Type "Info"
                    }
                    "group" {
                        Remove-ADGroup -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log (": Deleted group $cleanName from AD") -Type "Info"
                    }
                    "ou" {
                        Remove-ADOrganizationalUnit -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log (": Deleted OU $cleanName from AD") -Type "Info"
                    }
                    "computer" {
                        Remove-ADComputer -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log (": Deleted computer $cleanName from AD") -Type "Info"
                    }
                    default {
                        Remove-ADObject -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Debug-Log (": Deleted $objectType $cleanName from AD") -Type "Info"
                    }
                }

                Show-Modal "Deleted" "$objectType '$cleanName' deleted successfully"

                # Refresh data
                Refresh-Data -domain $Script:CurrentDomain
                Update-FilterStatusLabel -label $Script:FilterStatusLabel
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
        $user = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
        if ($user) { $currentOU = $user.OU }
    }

    $lblCurrentOU = [Terminal.Gui.Label]::new($currentOU); $lblCurrentOU.X=20; $lblCurrentOU.Y=1; $dlg.Add($lblCurrentOU)

    $lblTarget = [Terminal.Gui.Label]::new("Move to OU:"); $lblTarget.X=2; $lblTarget.Y=3; $dlg.Add($lblTarget)

    # Get list of OUs
    $ouList = if ($Script:DemoMode) {
        $Script:Users | Select-Object -ExpandProperty OU -Unique | Sort-Object
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
                if ($Script:DemoMode) {
                    # Demo mode - update in-memory
                    if ($objectType.ToLower() -eq "user") {
                        $user = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                        if ($user) {
                            $user.OU = $targetOU
                            Debug-Log (": Moved user $cleanName to $targetOU (demo mode)") -Type "Info"
                        }
                    }

                    Show-Modal "Success" "Object moved successfully (demo mode)"
                    Build-Tree -domain $Script:CurrentDomain
                    Update-FilterStatusLabel -label $Script:FilterStatusLabel
                    [Terminal.Gui.Application]::RequestStop()

                } else {
                    # Production mode - move in AD
                    $adObject = Get-ADObject -Filter "Name -eq '$cleanName'" -ErrorAction Stop
                    Move-ADObject -Identity $adObject.DistinguishedName -TargetPath $targetOU -ErrorAction Stop

                    Debug-Log (": Moved $cleanName to $targetOU in AD") -Type "Info"
                   Show-Modal "Success" "Object moved successfully"

                    # Refresh data
                    Refresh-Data -domain $Script:CurrentDomain
                    Update-FilterStatusLabel -label $Script:FilterStatusLabel
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
    $dlg = [Terminal.Gui.Dialog]::new("Change Domain", 50, 8)

    $lbl = [Terminal.Gui.Label]::new("Domain Name:")
    $lbl.X = 2
    $lbl.Y = 1
    $dlg.Add($lbl)

    $txtDomain = [Terminal.Gui.TextField]::new($Script:CurrentDomain)
    $txtDomain.X = 15
    $txtDomain.Y = 1
    $txtDomain.Width = 30
    $dlg.Add($txtDomain)

    $okBtn = [Terminal.Gui.Button]::new("OK")
    $okBtn.X = 10
    $okBtn.Y = 4
    $okBtn.add_Clicked({
        $domainString = -join ($txtDomain.Text | ForEach-Object { [char]$_ })
        Debug-Log (": OK pressed, Domain = $domainString") -Type "Info"

        # Close the dialog first
        [Terminal.Gui.Application]::RequestStop()

        # Schedule the domain change after dialog closes
        [Terminal.Gui.Application]::MainLoop.Invoke({
            try {
                Update-Status "Changing domain to $domainString..." -spinner
                Start-Sleep -Milliseconds 100  # Brief pause so user sees the message

                # Update the current domain
                $Script:CurrentDomain = $domainString
                Debug-Log (": CurrentDomain set to: $Script:CurrentDomain") -Type "Info"

                # Refresh data with tree rebuild
                Update-Status "Loading domain data..." -spinner
                $result = Refresh-Data -domain $Script:CurrentDomain -RebuildTree

                if ($result) {
                    Debug-Log (": Domain change successful") -Type "Success"
                    Update-Status "Domain changed to $domainString" -final

                    # Update filter status if it exists
                    if ($Script:FilterStatusLabel) {
                        Update-FilterStatusLabel -label $Script:FilterStatusLabel
                    }
                } else {
                    Debug-Log (": Domain change failed - Refresh-Data returned false") -Type "Error"
                    Update-Status "Failed to load domain $domainString" -final
                    Show-Modal "Error" "Failed to load domain '$domainString'`n`nCheck logs for details"
                }
            } catch {
                Debug-Log (": Domain change error: $($_.Exception.Message)") -Type "Error"
                Update-Status "Domain change error" -final
                Show-Modal "Error" "Error changing domain:`n`n$($_.Exception.Message)"
            }
        })
    })
    $dlg.Add($okBtn)

    $cancelBtn = [Terminal.Gui.Button]::new("Cancel")
    $cancelBtn.X = 25
    $cancelBtn.Y = 4
    $cancelBtn.add_Clicked({
        Debug-Log (": Cancel pressed") -Type "Info"
        [Terminal.Gui.Application]::RequestStop()
    })
    $dlg.Add($cancelBtn)

    [Terminal.Gui.Application]::Run($dlg)
}

# ------------------------- Change DC Dialog ------------------------
function Show-ChangeDCDialog {
    $dlg = [Terminal.Gui.Dialog]::new("Change Domain Controller",50,12)
    $dlg.Add([Terminal.Gui.Label]::new("Select Domain Controller:")) | Out-Null
    $dcNames = $Script:DCs | ForEach-Object { $_.Name }
    $listView = [Terminal.Gui.ListView]::new($dcNames); $listView.X=0; $listView.Y=1; $listView.Width=48; $listView.Height=6
    $dlg.Add($listView)
    $okBtn = [Terminal.Gui.Button]::new("OK"); $okBtn.X=10; $okBtn.Y=8
    $okBtn.add_Clicked({
        if ($listView.SelectedItem -ge 0) { $Script:CurrentDC = $dcNames[$listView.SelectedItem]; $status.Items[1].Title = "DC: $Script:CurrentDC" }
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
if ($null -ne $Script:tree) {
  ## backup will remove for prod code
  ##  $tree.add_KeyPress({ param($sender,$keyArgs) if ($keyArgs.KeyEvent.Key -eq [Terminal.Gui.Key]::Enter -and $tree.SelectedObject) { $tree.SelectedObject.Expanded = -not $tree.SelectedObject.Expanded; $tree.SetNeedsDisplay(); $keyArgs.Handled = $true } })
  $Script:tree.Add_KeyPress({
    param($sender, $keyArgs)

    if ($keyArgs.KeyEvent.Key -ne [Terminal.Gui.Key]::Enter) { return }

    $node = $Script:tree.SelectedNode
    if (-not $node) {
        Debug-Log ("Enter pressed but no selected node — ignoring safely.") -Type "Info"
        return
    }

    # Toggle expand/collapse properly
    if ($node.IsExpanded) {
        $Script:tree.CollapseNode($node)
    } else {
        $Script:tree.ExpandNode($node)
    }

    $Script:tree.SetNeedsDisplay()
    $keyArgs.Handled = $true
  })
} else {
    Debug-Log "Attempted to attach KeyPress handler but $Script:tree is null" -Type "Warn"
}


# ------------------------- AD Search Dialog ------------------------
# DSA-TUI Advanced Search Module v1.0
# Features: LDAP filters, saved searches, export results

# Global for saved searches
if (-not $Script:SavedSearches) {
    $Script:SavedSearches = @(
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

function Invoke-ADSearch {
    param(
        [Parameter(Mandatory=$true)] [Terminal.Gui.TextField] $UserField,
        [Parameter(Mandatory=$true)] [Terminal.Gui.TextField] $DomainField,
        [Parameter(Mandatory=$true)] [Terminal.Gui.ComboBox] $ObjType,
        [Parameter(Mandatory=$true)] [Terminal.Gui.TabView] $TabView,
        [Parameter(Mandatory=$true)] [Terminal.Gui.TextView] $TxtOutput,
        [Parameter(Mandatory=$true)] [Terminal.Gui.CheckBox] $ChkDisabledOnly,
        [Parameter(Mandatory=$true)] [Terminal.Gui.TextView] $LdapFilter,
        [Parameter(Mandatory=$true)] [Terminal.Gui.TabView+Tab] $AdvTab,
        # New parameters for clipboard buttons
        [Parameter(Mandatory=$false)] [Terminal.Gui.Button] $BtnCopyQuery,
        [Parameter(Mandatory=$false)] [Terminal.Gui.Button] $BtnPasteQuery,
        [Parameter(Mandatory=$false)] [Terminal.Gui.Button] $BtnCopyResults
    )

    $searchName = $UserField.Text.ToString().Trim()
    $domain = $DomainField.Text.ToString().Trim()
    $objType = $ObjType.Text.ToString()
    $currentTab = $TabView.SelectedTab

    try {
        $objs = @()

        if ($currentTab -eq $AdvTab) {
            # LDAP filter search
            $filter = $LdapFilter.Text.ToString().Trim()
            if (-not $filter) { $TxtOutput.Text="Please enter an LDAP filter."; return }

            if ($Script:DemoMode) {
                $TxtOutput.Text="LDAP search not supported in demo mode. Use Basic search."
                return
            } else {
                $loading = Show-LoadingDialog -Message "Executing LDAP query..."
                try {
                    $objs = Get-ADObject -LDAPFilter $filter -Properties Name,ObjectClass,DistinguishedName -ErrorAction Stop |
                        Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={$_.ObjectClass}}, @{Name='DN';Expression={$_.DistinguishedName}}
                    $Script:lastSearchType = "LDAP"
                } finally { Close-LoadingDialog $loading }
            }
        } else {
            # Basic search
            if (-not $searchName) { $TxtOutput.Text="Please enter a name."; return }

            if ($Script:DemoMode) {
                switch ($objType) {
                    "User" { $objs = $Script:Users | Where-Object { $_.Name -like "*$searchName*" } | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"user"}} }
                    "Group" {
                        $matchedGroups = @(); foreach ($u in $Script:Users) { foreach ($g in $u.Groups) { if ($g -like "*$searchName*") { $matchedGroups += $g } } }
                        $objs = ($matchedGroups | Sort-Object -Unique) | ForEach-Object { [PSCustomObject]@{ Name=$_; Type="group" } }
                    }
                    "OU" {
                        $ouNames = ($Script:Users | Select-Object -ExpandProperty OU -Unique)
                        $objs = ($ouNames | Where-Object { $_ -like "*$searchName*" }) | ForEach-Object { [PSCustomObject]@{ Name=$_; Type="organizationalUnit" } }
                    }
                    "Computer" { $objs = @() }
                    "Contact" { $objs = @() }
                }
                $Script:lastSearchType = "Basic ($objType)"
            } else {
                $loading = Show-LoadingDialog -Message "Searching AD for $objType '$searchName'..."
                try {
                    $filterStr = "Name -like '*$searchName*'"
                    if ($ChkDisabledOnly.Checked -and $objType -eq "User") {
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
                    $Script:lastSearchType = "Basic ($objType)"
                } finally { Close-LoadingDialog $loading }
            }
        }

        # Store results for export
        $Script:lastSearchResults = $objs

        if (-not $objs -or $objs.Count -eq 0) { $TxtOutput.Text = "No results found"; return }

        # Display results
        $resultText = "Found $($objs.Count) object(s):`n`n"
        $resultText += ($objs | ForEach-Object { "$($_.Name) [$($_.Type)]" }) -join "`n"
        $TxtOutput.Text = $resultText

    } catch {
        $TxtOutput.Text = "Error: $($_.Exception.Message)"
    }
}

# =====================================================
# Clipboard Helper Functions
# =====================================================

function Copy-LDAPQueryToClipboard {
    param([Terminal.Gui.TextView] $LdapFilter)

    try {
        $query = $LdapFilter.Text.ToString().Trim()

        if (-not $query) {
            Show-Modal "Info" "No LDAP query to copy"
            return
        }

        Set-Clipboard -Value $query
        Debug-Log ": Copied LDAP query to clipboard" -Type "Success"
        Show-Modal "Success" "LDAP query copied to clipboard"

    } catch {
        Debug-Log ": Failed to copy to clipboard: $($_.Exception.Message)" -Type "Error"
        Show-Modal "Error" "Failed to copy to clipboard:`n$($_.Exception.Message)"
    }
}

function Paste-LDAPQueryFromClipboard {
    param([Terminal.Gui.TextView] $LdapFilter)

    try {
        $clipboardText = Get-Clipboard -Raw

        if (-not $clipboardText) {
            Show-Modal "Info" "Clipboard is empty"
            return
        }

        $LdapFilter.Text = $clipboardText
        Debug-Log ": Pasted LDAP query from clipboard" -Type "Success"
        [Terminal.Gui.Application]::Refresh()

    } catch {
        Debug-Log ": Failed to paste from clipboard: $($_.Exception.Message)" -Type "Error"
        Show-Modal "Error" "Failed to paste from clipboard:`n$($_.Exception.Message)"
    }
}

function Copy-SearchResultsToClipboard {
    param([Terminal.Gui.TextView] $TxtOutput)

    try {
        if (-not $Script:lastSearchResults -or $Script:lastSearchResults.Count -eq 0) {
            Show-Modal "Info" "No search results to copy"
            return
        }

        # Format results for clipboard
        $clipboardText = "# AD Search Results - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
        $clipboardText += "# Search Type: $($Script:lastSearchType)`n"
        $clipboardText += "# Total Results: $($Script:lastSearchResults.Count)`n"
        $clipboardText += "`n"

        # Add results in multiple formats for flexibility

        # Format 1: Simple list
        $clipboardText += "=== SIMPLE LIST ===`n"
        foreach ($obj in $Script:lastSearchResults) {
            $clipboardText += "$($obj.Name) [$($obj.Type)]`n"
        }

        $clipboardText += "`n=== CSV FORMAT ===`n"
        $clipboardText += "Name,Type"
        if ($Script:lastSearchResults[0].PSObject.Properties['DN']) {
            $clipboardText += ",DistinguishedName"
        }
        if ($Script:lastSearchResults[0].PSObject.Properties['Enabled']) {
            $clipboardText += ",Enabled"
        }
        $clipboardText += "`n"

        # Format 2: CSV
        foreach ($obj in $Script:lastSearchResults) {
            $clipboardText += "`"$($obj.Name)`",`"$($obj.Type)`""
            if ($obj.PSObject.Properties['DN']) {
                $clipboardText += ",`"$($obj.DN)`""
            }
            if ($obj.PSObject.Properties['Enabled']) {
                $clipboardText += ",`"$($obj.Enabled)`""
            }
            $clipboardText += "`n"
        }

        # Format 3: PowerShell array (if useful)
        $clipboardText += "`n=== POWERSHELL NAMES ===`n"
        $clipboardText += "@(`n"
        $names = $Script:lastSearchResults | ForEach-Object { "    `"$($_.Name)`"" }
        $clipboardText += ($names -join ",`n")
        $clipboardText += "`n)`n"

        Set-Clipboard -Value $clipboardText
        Debug-Log ": Copied $($Script:lastSearchResults.Count) search results to clipboard" -Type "Success"
        Show-Modal "Success" "Copied $($Script:lastSearchResults.Count) results to clipboard`n`nFormats included:`n• Simple list`n• CSV`n• PowerShell array"

    } catch {
        Debug-Log ": Failed to copy results to clipboard: $($_.Exception.Message)" -Type "Error"
        Show-Modal "Error" "Failed to copy results to clipboard:`n$($_.Exception.Message)"
    }
}

## TODO: Implement this:
# =====================================================
# Example: How to Add Clipboard Buttons to Your UI
# =====================================================

<# In your Search/Lookup tab creation, add these buttons:

# LDAP Filter section
$lbl = [Terminal.Gui.Label]::new("LDAP Filter:"); $lbl.X=2; $lbl.Y=$y; $advView.Add($lbl); $y+=1
$txtLdapFilter = [Terminal.Gui.TextView]::new(); $txtLdapFilter.X=2; $txtLdapFilter.Y=$y; $txtLdapFilter.Width=[Terminal.Gui.Dim]::Fill(2); $txtLdapFilter.Height=4
$advView.Add($txtLdapFilter); $y+=5

# Clipboard buttons for LDAP query
$btnCopyQuery = [Terminal.Gui.Button]::new("Copy Query"); $btnCopyQuery.X=2; $btnCopyQuery.Y=$y
$btnCopyQuery.add_Clicked({ Copy-LDAPQueryToClipboard -LdapFilter $txtLdapFilter })
$advView.Add($btnCopyQuery)

$btnPasteQuery = [Terminal.Gui.Button]::new("Paste Query"); $btnPasteQuery.X=18; $btnPasteQuery.Y=$y
$btnPasteQuery.add_Clicked({ Paste-LDAPQueryFromClipboard -LdapFilter $txtLdapFilter })
$advView.Add($btnPasteQuery)
$y+=2

# Search button
$btnSearch = [Terminal.Gui.Button]::new("Search"); $btnSearch.X=2; $btnSearch.Y=$y
$btnSearch.add_Clicked({
    Invoke-ADSearch -UserField $txtSearchName -DomainField $txtSearchDomain -ObjType $cmbSearchType `
                    -TabView $searchTabView -TxtOutput $txtSearchOutput -ChkDisabledOnly $chkDisabledOnly `
                    -LdapFilter $txtLdapFilter -AdvTab $advTab
})
$advView.Add($btnSearch); $y+=3

# Results section
$lbl = [Terminal.Gui.Label]::new("Results:"); $lbl.X=2; $lbl.Y=$y; $advView.Add($lbl); $y+=1
$txtSearchOutput = [Terminal.Gui.TextView]::new(); $txtSearchOutput.X=2; $txtSearchOutput.Y=$y
$txtSearchOutput.Width=[Terminal.Gui.Dim]::Fill(2); $txtSearchOutput.Height=[Terminal.Gui.Dim]::Fill(4); $txtSearchOutput.ReadOnly=$true
$advView.Add($txtSearchOutput); $y+=[Terminal.Gui.Dim]::Fill(3)

# Copy results button
$btnCopyResults = [Terminal.Gui.Button]::new("Copy Results to Clipboard"); $btnCopyResults.X=2; $btnCopyResults.Y=[Terminal.Gui.Pos]::Bottom($txtSearchOutput)+1
$btnCopyResults.add_Clicked({ Copy-SearchResultsToClipboard -TxtOutput $txtSearchOutput })
$advView.Add($btnCopyResults)
#>

# DSA-TUI Context Menu & Refresh Module v1.0

function Show-ADSearchDialog {
    <#
    .SYNOPSIS
    Displays an AD Search dialog with Basic and Advanced LDAP search capabilities

    .DESCRIPTION
    Creates a modal dialog for searching Active Directory with:
    - Basic search (by name and object type)
    - Advanced LDAP filter search
    - Clipboard support for queries and results
    - Export functionality
    #>

    Debug-Log ": Opening AD Search Dialog" -Type "Info"

    try {
        # ---------------------- Buttons ----------------------
        $btnClose = [Terminal.Gui.Button]::new("Close")
        $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })

        # ---------------------- Dialog ----------------------
        $dlg = [Terminal.Gui.Dialog]::new("Active Directory Search", 120, 40, $btnClose)

        # ---------------------- TabView for Basic/Advanced ----------------------
        $searchTabView = [Terminal.Gui.TabView]::new()
        $searchTabView.X = 0
        $searchTabView.Y = 0
        $searchTabView.Width = [Terminal.Gui.Dim]::Fill()
        $searchTabView.Height = [Terminal.Gui.Dim]::Fill(1)

        # ==================== Basic Search Tab ====================
        $basicTab = [Terminal.Gui.TabView+Tab]::new()
        $basicTab.Text = "Basic Search"
        $basicView = [Terminal.Gui.View]::new()
        $basicView.X = 0; $basicView.Y = 0
        $basicView.Width = [Terminal.Gui.Dim]::Fill()
        $basicView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1

        # Domain
        $lbl = [Terminal.Gui.Label]::new("Domain:"); $lbl.X=2; $lbl.Y=$y; $basicView.Add($lbl)
        $txtDomain = [Terminal.Gui.TextField]::new($Script:CurrentDomain ?? ""); $txtDomain.X=20; $txtDomain.Y=$y; $txtDomain.Width=40
        $basicView.Add($txtDomain); $y+=2

        # Search Name
        $lbl = [Terminal.Gui.Label]::new("Name:"); $lbl.X=2; $lbl.Y=$y; $basicView.Add($lbl)
        $txtSearchName = [Terminal.Gui.TextField]::new(""); $txtSearchName.X=20; $txtSearchName.Y=$y; $txtSearchName.Width=40
        $basicView.Add($txtSearchName); $y+=2

        # Object Type
        $lbl = [Terminal.Gui.Label]::new("Object Type:"); $lbl.X=2; $lbl.Y=$y; $basicView.Add($lbl)
        $cmbObjectType = [Terminal.Gui.ComboBox]::new(); $cmbObjectType.X=20; $cmbObjectType.Y=$y; $cmbObjectType.Width=20
        $cmbObjectType.SetSource(@("User", "Group", "Computer", "OU", "Contact"))
        $cmbObjectType.SelectedItem = 0
        $basicView.Add($cmbObjectType); $y+=2

        # Disabled only checkbox (for users)
        $chkDisabledOnly = [Terminal.Gui.CheckBox]::new("Disabled accounts only"); $chkDisabledOnly.X=20; $chkDisabledOnly.Y=$y
        $basicView.Add($chkDisabledOnly); $y+=2

        # Search button
        $btnSearch = [Terminal.Gui.Button]::new("Search"); $btnSearch.X=20; $btnSearch.Y=$y
        $basicView.Add($btnSearch); $y+=3

        # Results
        $lbl = [Terminal.Gui.Label]::new("Results:"); $lbl.X=2; $lbl.Y=$y; $basicView.Add($lbl); $y+=1
        $txtResults = [Terminal.Gui.TextView]::new()
        $txtResults.X = 2
        $txtResults.Y = $y
        $txtResults.Width = [Terminal.Gui.Dim]::Fill(2)
        $txtResults.Height = [Terminal.Gui.Dim]::Fill(3)
        $txtResults.ReadOnly = $true
        $basicView.Add($txtResults)

        # Copy Results button
        $btnCopyResults = [Terminal.Gui.Button]::new("Copy Results to Clipboard")
        $btnCopyResults.X = 2
        $btnCopyResults.Y = [Terminal.Gui.Pos]::Bottom($txtResults) + 1
        $btnCopyResults.add_Clicked({ Copy-SearchResultsToClipboard -TxtOutput $txtResults })
        $basicView.Add($btnCopyResults)

        $basicTab.View = $basicView
        $searchTabView.AddTab($basicTab, $false)

        # ==================== Advanced LDAP Tab ====================
        $advTab = [Terminal.Gui.TabView+Tab]::new()
        $advTab.Text = "Advanced (LDAP)"
        $advView = [Terminal.Gui.View]::new()
        $advView.X = 0; $advView.Y = 0
        $advView.Width = [Terminal.Gui.Dim]::Fill()
        $advView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1

        # Domain
        $lbl = [Terminal.Gui.Label]::new("Domain:"); $lbl.X=2; $lbl.Y=$y; $advView.Add($lbl)
        $txtDomainAdv = [Terminal.Gui.TextField]::new($Script:CurrentDomain ?? ""); $txtDomainAdv.X=20; $txtDomainAdv.Y=$y; $txtDomainAdv.Width=40
        $advView.Add($txtDomainAdv); $y+=2

        # LDAP Filter
        $lbl = [Terminal.Gui.Label]::new("LDAP Filter:"); $lbl.X=2; $lbl.Y=$y; $advView.Add($lbl); $y+=1

        # Example filters label
        $lblExamples = [Terminal.Gui.Label]::new("Examples: (&(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))");
        $lblExamples.X=2; $lblExamples.Y=$y; $lblExamples.ColorScheme = [Terminal.Gui.Colors]::TopLevel
        $advView.Add($lblExamples); $y+=1

        $txtLdapFilter = [Terminal.Gui.TextView]::new()
        $txtLdapFilter.X = 2
        $txtLdapFilter.Y = $y
        $txtLdapFilter.Width = [Terminal.Gui.Dim]::Fill(2)
        $txtLdapFilter.Height = 5
        $advView.Add($txtLdapFilter); $y+=6

        # Clipboard buttons for LDAP query
        $btnCopyQuery = [Terminal.Gui.Button]::new("Copy Query"); $btnCopyQuery.X=2; $btnCopyQuery.Y=$y
        $btnCopyQuery.add_Clicked({ Copy-LDAPQueryToClipboard -LdapFilter $txtLdapFilter })
        $advView.Add($btnCopyQuery)

        $btnPasteQuery = [Terminal.Gui.Button]::new("Paste Query"); $btnPasteQuery.X=18; $btnPasteQuery.Y=$y
        $btnPasteQuery.add_Clicked({ Paste-LDAPQueryFromClipboard -LdapFilter $txtLdapFilter })
        $advView.Add($btnPasteQuery); $y+=2

        # Search button
        $btnSearchAdv = [Terminal.Gui.Button]::new("Execute LDAP Query"); $btnSearchAdv.X=2; $btnSearchAdv.Y=$y
        $advView.Add($btnSearchAdv); $y+=3

        # Results
        $lbl = [Terminal.Gui.Label]::new("Results:"); $lbl.X=2; $lbl.Y=$y; $advView.Add($lbl); $y+=1
        $txtResultsAdv = [Terminal.Gui.TextView]::new()
        $txtResultsAdv.X = 2
        $txtResultsAdv.Y = $y
        $txtResultsAdv.Width = [Terminal.Gui.Dim]::Fill(2)
        $txtResultsAdv.Height = [Terminal.Gui.Dim]::Fill(3)
        $txtResultsAdv.ReadOnly = $true
        $advView.Add($txtResultsAdv)

        # Copy Results button
        $btnCopyResultsAdv = [Terminal.Gui.Button]::new("Copy Results to Clipboard")
        $btnCopyResultsAdv.X = 2
        $btnCopyResultsAdv.Y = [Terminal.Gui.Pos]::Bottom($txtResultsAdv) + 1
        $btnCopyResultsAdv.add_Clicked({ Copy-SearchResultsToClipboard -TxtOutput $txtResultsAdv })
        $advView.Add($btnCopyResultsAdv)

        $advTab.View = $advView
        $searchTabView.AddTab($advTab, $false)

        # ==================== Wire up Search Buttons ====================

        # Basic Search
        $btnSearch.add_Clicked({
            Invoke-ADSearch -UserField $txtSearchName -DomainField $txtDomain -ObjType $cmbObjectType `
                           -TabView $searchTabView -TxtOutput $txtResults -ChkDisabledOnly $chkDisabledOnly `
                           -LdapFilter $txtLdapFilter -AdvTab $advTab
        })

        # Advanced Search
        $btnSearchAdv.add_Clicked({
            Invoke-ADSearch -UserField $txtSearchName -DomainField $txtDomainAdv -ObjType $cmbObjectType `
                           -TabView $searchTabView -TxtOutput $txtResultsAdv -ChkDisabledOnly $chkDisabledOnly `
                           -LdapFilter $txtLdapFilter -AdvTab $advTab
        })

        # Add TabView to dialog
        $dlg.Add($searchTabView)

        Debug-Log ": AD Search Dialog created, running" -Type "Success"

        # Run the dialog
        [Terminal.Gui.Application]::Run($dlg)

        Debug-Log ": AD Search Dialog closed" -Type "Info"

    } catch {
        Debug-Log ": Exception in Show-ADSearchDialog: $($_.Exception.Message)" -Type "Error"
        Show-Modal "Error" "Failed to open search dialog:`n$($_.Exception.Message)"
    }
}

# =====================================================
# Example Common LDAP Filters (for reference/docs)
# =====================================================

<#
Common LDAP Filters:

# All users
(objectClass=user)

# All enabled users
(&(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))

# All disabled users
(&(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))

# All groups
(objectClass=group)

# All computers
(objectClass=computer)

# Users with email
(&(objectClass=user)(mail=*))

# Users in specific OU (need DN)
(&(objectClass=user)(distinguishedName=*,OU=IT,DC=example,DC=com))

# Users created after date
(&(objectClass=user)(whenCreated>=20240101000000.0Z))

# Users whose password never expires
(&(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=65536))

# Users with admin in name
(&(objectClass=user)(name=*admin*))

# Security groups
(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=2147483648))

# Distribution groups
(&(objectClass=group)(!(groupType:1.2.840.113556.1.4.803:=2147483648)))
#>

## ------------------------- Refresh Tree Function ------------------------
function Refresh-TreeData {
    Debug-Log (": Refreshing tree data...") -Type "Info"

    # Show loading dialog
    $loadingDlg = Show-LoadingDialog -Message "Refreshing Active Directory data..."

    try {
        # Reload domain data
        Load-DomainData -domain $Script:CurrentDomain

        # Rebuild tree
        Build-Tree -domain $Script:CurrentDomain
        # Add after Build-Tree calls:
        Update-FilterStatusLabel -label $Script:FilterStatusLabel

        Debug-Log (": Tree refreshed successfully") -Type "Info"
    } finally {
        Close-LoadingDialog $loadingDlg
    }

    Show-Modal "Refreshed" "Active Directory data refreshed successfully"
}

function Show-DCPropertiesDialog {
    param($dc)

    # Accept either DC object or DC name
    if ($dc -is [string]) {
        $dcName = $dc
        Debug-Log ": Looking for DC: $dcName" -Type "Info"

        # Check if DCs array exists
        if (-not $Script:DCs) {
            Debug-Log ": Script:DCs is null or not initialized" -Type "Error"
            Show-Modal "Error" "Domain Controllers list is not loaded"
            return
        }

        Debug-Log ": Script:DCs has $($Script:DCs.Count) entries" -Type "Info"

        # Find the DC
        $dc = $Script:DCs | Where-Object { $_.Name -eq $dcName } | Select-Object -First 1

        if (-not $dc) {
            Debug-Log ": DC '$dcName' not found in Script:DCs" -Type "Error"
            Show-Modal "Not Found" "DC '$dcName' not found in the domain controllers list"
            return
        }

        Debug-Log ": Found DC: $($dc.Name)" -Type "Success"
    }

    if (-not $dc) {
        Debug-Log ": DC object is null after lookup" -Type "Error"
        Show-Modal "Error" "DC object is null"
        return
    }

    Debug-Log ": Showing DC properties for: $($dc.Name)" -Type "Info"

    try {
        # ---------------------- Buttons ----------------------
        $btnOK = [Terminal.Gui.Button]::new("OK")
        $btnOK.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })

        # ---------------------- Dialog ----------------------
        $dlg = [Terminal.Gui.Dialog]::new("Domain Controller Properties - $($dc.Name)", 110, 35, $btnOK)

        # ---------------------- TabView ----------------------
        $tabView = [Terminal.Gui.TabView]::new()
        $tabView.X = 0
        $tabView.Y = 0
        $tabView.Width = [Terminal.Gui.Dim]::Fill()
        $tabView.Height = [Terminal.Gui.Dim]::Fill(1)

        # ==================== General Tab ====================
        $generalTab = [Terminal.Gui.TabView+Tab]::new()
        $generalTab.Text = "General"
        $generalView = [Terminal.Gui.View]::new()
        $generalView.X = 0; $generalView.Y = 0
        $generalView.Width = [Terminal.Gui.Dim]::Fill()
        $generalView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1

        # Basic Information
        $lbl = [Terminal.Gui.Label]::new("Domain Controller Information"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Name:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($dc.Name ?? "Unknown"); $lbl.X=25; $lbl.Y=$y; $generalView.Add($lbl); $y+=1

        $hostname = if ($dc.HostName) { $dc.HostName } elseif ($dc.DNSHostName) { $dc.DNSHostName } else { $dc.Name }
        $lbl = [Terminal.Gui.Label]::new("Hostname:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($hostname); $lbl.X=25; $lbl.Y=$y; $generalView.Add($lbl); $y+=1

        if ($dc.Site) {
            $lbl = [Terminal.Gui.Label]::new("Site:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($dc.Site); $lbl.X=25; $lbl.Y=$y; $generalView.Add($lbl); $y+=1
        }

        if ($dc.Domain) {
            $lbl = [Terminal.Gui.Label]::new("Domain:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($dc.Domain); $lbl.X=25; $lbl.Y=$y; $generalView.Add($lbl); $y+=1
        }

        if ($dc.Forest) {
            $lbl = [Terminal.Gui.Label]::new("Forest:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($dc.Forest); $lbl.X=25; $lbl.Y=$y; $generalView.Add($lbl); $y+=1
        }

        if ($dc.Location) {
            $lbl = [Terminal.Gui.Label]::new("Location:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($dc.Location); $lbl.X=25; $lbl.Y=$y; $generalView.Add($lbl); $y+=1
        }

        $y+=1

        # Network Information
        $lbl = [Terminal.Gui.Label]::new("Network"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl); $y+=2

        if ($dc.IPv4Address -or $dc.IPAddress) {
            $ip = if ($dc.IPv4Address) { $dc.IPv4Address } else { $dc.IPAddress }
            $lbl = [Terminal.Gui.Label]::new("IPv4 Address:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($ip); $lbl.X=25; $lbl.Y=$y; $generalView.Add($lbl); $y+=1
        }

        if ($dc.IPv6Address) {
            $lbl = [Terminal.Gui.Label]::new("IPv6 Address:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($dc.IPv6Address); $lbl.X=25; $lbl.Y=$y; $generalView.Add($lbl); $y+=1
        }

        $y+=1

        # Operating System
        $lbl = [Terminal.Gui.Label]::new("Operating System"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl); $y+=2

        if ($dc.OperatingSystem -or $dc.OS) {
            $os = if ($dc.OperatingSystem) { $dc.OperatingSystem } else { $dc.OS }
            $lbl = [Terminal.Gui.Label]::new("OS:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($os); $lbl.X=25; $lbl.Y=$y; $generalView.Add($lbl); $y+=1
        }

        if ($dc.OperatingSystemVersion) {
            $lbl = [Terminal.Gui.Label]::new("OS Version:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($dc.OperatingSystemVersion); $lbl.X=25; $lbl.Y=$y; $generalView.Add($lbl); $y+=1
        }

        $y+=1

        # Capabilities
        $lbl = [Terminal.Gui.Label]::new("Capabilities"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl); $y+=2

        $isGC = if ($null -ne $dc.IsGlobalCatalog) { $dc.IsGlobalCatalog.ToString() } else { "Unknown" }
        $lbl = [Terminal.Gui.Label]::new("Global Catalog:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $lbl = [Terminal.Gui.Label]::new($isGC); $lbl.X=25; $lbl.Y=$y; $generalView.Add($lbl); $y+=1

        if ($null -ne $dc.IsReadOnly) {
            $lbl = [Terminal.Gui.Label]::new("Read-Only:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($dc.IsReadOnly.ToString()); $lbl.X=25; $lbl.Y=$y; $generalView.Add($lbl); $y+=1
        }

        if ($null -ne $dc.Enabled) {
            $lbl = [Terminal.Gui.Label]::new("Enabled:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($dc.Enabled.ToString()); $lbl.X=25; $lbl.Y=$y; $generalView.Add($lbl); $y+=1
        }

        $generalTab.View = $generalView
        $tabView.AddTab($generalTab, $false)

        # ==================== Roles Tab ====================
        $rolesTab = [Terminal.Gui.TabView+Tab]::new()
        $rolesTab.Text = "Roles"
        $rolesView = [Terminal.Gui.View]::new()
        $rolesView.X = 0; $rolesView.Y = 0
        $rolesView.Width = [Terminal.Gui.Dim]::Fill()
        $rolesView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        $lbl = [Terminal.Gui.Label]::new("FSMO Roles"); $lbl.X=2; $lbl.Y=$y; $rolesView.Add($lbl); $y+=2

        # Get FSMO roles from either FSMORoles or OperationMasterRoles
        $fsmoRoles = @()
        if ($dc.PSObject.Properties['FSMORoles'] -and $dc.FSMORoles -and $dc.FSMORoles.Count -gt 0) {
            $fsmoRoles = $dc.FSMORoles
        } elseif ($dc.PSObject.Properties['OperationMasterRoles'] -and $dc.OperationMasterRoles -and $dc.OperationMasterRoles.Count -gt 0) {
            $fsmoRoles = $dc.OperationMasterRoles
        }

        if ($fsmoRoles.Count -gt 0) {
            foreach ($role in $fsmoRoles) {
                $lbl = [Terminal.Gui.Label]::new("• $role"); $lbl.X=4; $lbl.Y=$y; $rolesView.Add($lbl); $y+=1
            }
        } else {
            $lbl = [Terminal.Gui.Label]::new("(None)"); $lbl.X=4; $lbl.Y=$y; $rolesView.Add($lbl); $y+=1
        }

        $rolesTab.View = $rolesView
        $tabView.AddTab($rolesTab, $false)

        # ==================== Replication Tab ====================
        $replTab = [Terminal.Gui.TabView+Tab]::new()
        $replTab.Text = "Replication"
        $replView = [Terminal.Gui.View]::new()
        $replView.X = 0; $replView.Y = 0
        $replView.Width = [Terminal.Gui.Dim]::Fill()
        $replView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        $lbl = [Terminal.Gui.Label]::new("Replication Status"); $lbl.X=2; $lbl.Y=$y; $replView.Add($lbl); $y+=2

        if ($dc.ReplicationHealth) {
            $lbl = [Terminal.Gui.Label]::new("Health:"); $lbl.X=4; $lbl.Y=$y; $replView.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($dc.ReplicationHealth); $lbl.X=25; $lbl.Y=$y; $replView.Add($lbl); $y+=1
        }

        if ($dc.LastReplication) {
            $lastRep = $dc.LastReplication.ToString('yyyy-MM-dd HH:mm:ss')
            $lbl = [Terminal.Gui.Label]::new("Last Replication:"); $lbl.X=4; $lbl.Y=$y; $replView.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($lastRep); $lbl.X=25; $lbl.Y=$y; $replView.Add($lbl); $y+=2
        }

        # Replication Partners
        if ($dc.ReplicationPartners -and $dc.ReplicationPartners.Count -gt 0) {
            $lbl = [Terminal.Gui.Label]::new("Replication Partners:"); $lbl.X=4; $lbl.Y=$y; $replView.Add($lbl); $y+=1

            foreach ($partner in $dc.ReplicationPartners) {
                $lbl = [Terminal.Gui.Label]::new("• $partner"); $lbl.X=6; $lbl.Y=$y; $replView.Add($lbl); $y+=1
            }
        }

        $replTab.View = $replView
        $tabView.AddTab($replTab, $false)

        # ==================== Services Tab ====================
        $servicesTab = [Terminal.Gui.TabView+Tab]::new()
        $servicesTab.Text = "Services"
        $servicesView = [Terminal.Gui.View]::new()
        $servicesView.X = 0; $servicesView.Y = 0
        $servicesView.Width = [Terminal.Gui.Dim]::Fill()
        $servicesView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1

        # Services Status
        if ($dc.Services) {
            $lbl = [Terminal.Gui.Label]::new("Service Status"); $lbl.X=2; $lbl.Y=$y; $servicesView.Add($lbl); $y+=2

            foreach ($service in $dc.Services.Keys | Sort-Object) {
                $status = $dc.Services[$service]
                $lbl = [Terminal.Gui.Label]::new("${service}:"); $lbl.X=4; $lbl.Y=$y; $servicesView.Add($lbl)
                $lbl = [Terminal.Gui.Label]::new($status); $lbl.X=20; $lbl.Y=$y; $servicesView.Add($lbl); $y+=1
            }
            $y+=1
        }

        # Boot Time
        if ($dc.LastBoot -or $dc.LastBootUpTime) {
            $lbl = [Terminal.Gui.Label]::new("System Information"); $lbl.X=2; $lbl.Y=$y; $servicesView.Add($lbl); $y+=2

            $lastBoot = if ($dc.LastBootUpTime) { $dc.LastBootUpTime } else { $dc.LastBoot }
            if ($lastBoot) {
                $bootTime = $lastBoot.ToString('yyyy-MM-dd HH:mm:ss')
                $uptime = (Get-Date) - $lastBoot
                $uptimeStr = "$($uptime.Days) days, $($uptime.Hours) hours"

                $lbl = [Terminal.Gui.Label]::new("Last Boot:"); $lbl.X=4; $lbl.Y=$y; $servicesView.Add($lbl)
                $lbl = [Terminal.Gui.Label]::new($bootTime); $lbl.X=20; $lbl.Y=$y; $servicesView.Add($lbl); $y+=1

                $lbl = [Terminal.Gui.Label]::new("Uptime:"); $lbl.X=4; $lbl.Y=$y; $servicesView.Add($lbl)
                $lbl = [Terminal.Gui.Label]::new($uptimeStr); $lbl.X=20; $lbl.Y=$y; $servicesView.Add($lbl); $y+=1
            }
        }

        $servicesTab.View = $servicesView
        $tabView.AddTab($servicesTab, $false)

        # ==================== Disk Space Tab ====================
        $diskTab = [Terminal.Gui.TabView+Tab]::new()
        $diskTab.Text = "Disk Space"
        $diskView = [Terminal.Gui.View]::new()
        $diskView.X = 0; $diskView.Y = 0
        $diskView.Width = [Terminal.Gui.Dim]::Fill()
        $diskView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1

        if ($dc.DiskSpace) {
            $lbl = [Terminal.Gui.Label]::new("Disk Usage"); $lbl.X=2; $lbl.Y=$y; $diskView.Add($lbl); $y+=2

            foreach ($drive in $dc.DiskSpace.Keys | Sort-Object) {
                $diskInfo = $dc.DiskSpace[$drive]

                $lbl = [Terminal.Gui.Label]::new("Drive ${drive}"); $lbl.X=4; $lbl.Y=$y; $diskView.Add($lbl); $y+=1

                if ($diskInfo.Total) {
                    $lbl = [Terminal.Gui.Label]::new("  Total:"); $lbl.X=6; $lbl.Y=$y; $diskView.Add($lbl)
                    $lbl = [Terminal.Gui.Label]::new($diskInfo.Total); $lbl.X=20; $lbl.Y=$y; $diskView.Add($lbl); $y+=1
                }

                if ($diskInfo.Used) {
                    $lbl = [Terminal.Gui.Label]::new("  Used:"); $lbl.X=6; $lbl.Y=$y; $diskView.Add($lbl)
                    $lbl = [Terminal.Gui.Label]::new($diskInfo.Used); $lbl.X=20; $lbl.Y=$y; $diskView.Add($lbl); $y+=1
                }

                if ($diskInfo.Free) {
                    $lbl = [Terminal.Gui.Label]::new("  Free:"); $lbl.X=6; $lbl.Y=$y; $diskView.Add($lbl)
                    $lbl = [Terminal.Gui.Label]::new($diskInfo.Free); $lbl.X=20; $lbl.Y=$y; $diskView.Add($lbl); $y+=1
                }

                if ($null -ne $diskInfo.PercentFree) {
                    $lbl = [Terminal.Gui.Label]::new("  % Free:"); $lbl.X=6; $lbl.Y=$y; $diskView.Add($lbl)
                    $lbl = [Terminal.Gui.Label]::new("$($diskInfo.PercentFree)%"); $lbl.X=20; $lbl.Y=$y; $diskView.Add($lbl); $y+=1
                }

                $y+=1
            }
        } else {
            $lbl = [Terminal.Gui.Label]::new("(No disk space information available)"); $lbl.X=4; $lbl.Y=$y; $diskView.Add($lbl); $y+=1
        }

        $diskTab.View = $diskView
        $tabView.AddTab($diskTab, $false)

        # Add TabView to dialog
        $dlg.Add($tabView)

        Debug-Log ": All DC tabs added, running dialog" -Type "Success"

        # Run the dialog
        [Terminal.Gui.Application]::Run($dlg)
        Debug-Log ": DC dialog closed normally" -Type "Info"

    } catch {
        Debug-Log ": Exception in Show-DCPropertiesDialog: $($_.Exception.Message)" -Type "Error"
        Show-Modal "Error" "Failed to display DC properties:`n$($_.Exception.Message)"
    }
}

function Show-GroupPropertiesDialog {
    param([string]$groupName)

    Debug-Log (": Showing group properties dialog for: $groupName") -Type "Info"

    if (-not $groupName) {
        Debug-Log (": Group name is null") -Type "Warn"
        return
    }

    # Find the group
    $group = $Script:Groups | Where-Object { $_.Name -eq $groupName } | Select-Object -First 1
    if (-not $group) {
        Show-Modal "Not Found" "Group '$groupName' not found"
        return
    }

    # ---------------- Buttons ----------------
    $btnOK     = [Terminal.Gui.Button]::new("OK")
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnApply  = [Terminal.Gui.Button]::new("Apply")

    # ---------------- Dialog ----------------
    $dlg = [Terminal.Gui.Dialog]::new(
        "Group Properties - $groupName",
        90, 35,
        $btnOK, $btnCancel, $btnApply
    )

    # ---------------- TabView ----------------
    $tabView = [Terminal.Gui.TabView]::new()
    $tabView.X = 0
    $tabView.Y = 0
    $tabView.Width  = [Terminal.Gui.Dim]::Fill()
    $tabView.Height = [Terminal.Gui.Dim]::Fill(3)

    # ==================== General Tab ====================
    $generalTab = [Terminal.Gui.TabView+Tab]::new()
    $generalTab.Text = "General"

    $generalView = [Terminal.Gui.View]::new()
    $y = 1

    $lbl = [Terminal.Gui.Label]::new("Display Name:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtName = [Terminal.Gui.TextField]::new($user.Name ?? ""); $txtName.X=20; $txtName.Y=$y; $txtName.Width=70
    $txtName.add_TextChanged({ $Script:groupChangesMade = $true })
    $generalView.Add($txtName)
    $y += 2

    $lbl = [Terminal.Gui.Label]::new("Description:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtDesc = [Terminal.Gui.TextField]::new($user.Description ?? ""); $txtDesc.X=20; $txtDesc.Y=$y; $txtDesc.Width=70
    $txtDesc.add_TextChanged({ $Script:groupChangesMade = $true })
    $generalView.Add($txtDesc)
    $y += 2

    $lbl = [Terminal.Gui.Label]::new("Office:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtOffice = [Terminal.Gui.TextField]::new($user.Office ?? ""); $txtOffice.X=20; $txtOffice.Y=$y; $txtOffice.Width=70
    $txtOffice.add_TextChanged({ $Script:groupChangesMade = $true })
    $generalView.Add($txtOffice)
    $y += 2

    Debug-Log ": User Phone value='$($user.OfficePhone)'" -Type "Debug"
    $lbl = [Terminal.Gui.Label]::new("Telephone:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtPhone = [Terminal.Gui.TextField]::new($user.OfficePhone ?? ""); $txtPhone.X=20; $txtPhone.Y=$y; $txtPhone.Width=70
    $txtPhone.add_TextChanged({ $Script:groupChangesMade = $true })
    $generalView.Add($txtPhone)
    $y += 2

    $lbl = [Terminal.Gui.Label]::new("Mobile Phone:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtMobile = [Terminal.Gui.TextField]::new($user.MobilePhone ?? ""); $txtMobile.X=20; $txtMobile.Y=$y; $txtMobile.Width=70
    $txtMobile.add_TextChanged({ $Script:groupChangesMade = $true })
    $generalView.Add($txtMobile)
    $y += 2

    Debug-Log ": User Email value='$($user.EmailAddress)'" -Type "Debug"
    $lbl = [Terminal.Gui.Label]::new("E-mail:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtEmail = [Terminal.Gui.TextField]::new($user.EmailAddress ?? ""); $txtEmail.X=20; $txtEmail.Y=$y; $txtEmail.Width=70
    $txtEmail.add_TextChanged({ $Script:groupChangesMade = $true })
    $generalView.Add($txtEmail)

    $generalTab.View = $generalView
    Register-Tab -Tab $generalTab

    # ==================== Members Tab ====================
    $membersTab = [Terminal.Gui.TabView+Tab]::new()
    $membersTab.Text = "Members"

    $membersView = [Terminal.Gui.View]::new()
    $members = $Script:Users | Where-Object { $_.Groups -contains $groupName } |
               ForEach-Object { $_.Name } | Sort-Object

    $lbl = [Terminal.Gui.Label]::new("Members:"); $lbl.X=2; $lbl.Y=1
    $membersView.Add($lbl)

    $lstMembers = [Terminal.Gui.ListView]::new()
    $lstMembers.SetSource($members)
    $lstMembers.X=2; $lstMembers.Y=3
    $lstMembers.Width=[Terminal.Gui.Dim]::Fill(2)
    $lstMembers.Height=[Terminal.Gui.Dim]::Fill(4)
    $membersView.Add($lstMembers)

    $membersTab.View = $membersView
    Register-Tab -Tab $membersTab

    # ---------------- Attach TabView ----------------
    $dlg.Add($tabView)

    # ---------------- Buttons ----------------
    $fields = @{
        txtDesc  = $txtDesc
        txtEmail = $txtEmail
    }

    $btnOK.add_Clicked({
        if (Apply-GroupChanges -group $group -fields $fields) {
            [Terminal.Gui.Application]::RequestStop()
        }
    }.GetNewClosure())

    $btnCancel.add_Clicked({
        if ($Script:groupChangesMade) {
            $result = [Terminal.Gui.MessageBox]::Query(
                60, 8,
                "Unsaved Changes",
                "You have unsaved changes. Discard them?",
                @("Yes","No")
            )
            if ($result -eq 0) {
                $Script:groupChangesMade = $false
                [Terminal.Gui.Application]::RequestStop()
            }
        } else {
            [Terminal.Gui.Application]::RequestStop()
        }
    }.GetNewClosure())

    $btnApply.add_Clicked({
        Apply-GroupChanges -group $group -fields $fields
    }.GetNewClosure())

    # ---------------- Run ----------------
    Debug-Log (": Show-GroupPropertiesDialog running") -Type "Info"
    Layout-Tabs
    [Terminal.Gui.Application]::Run($dlg)
    Debug-Log (": Show-GroupPropertiesDialog completed") -Type "Info"
}


# ----------------------- Show Computer Properties ----------------------
function Show-ComputerPropertiesDialog {
    param([string]$computerName)

    Debug-Log (": Showing computer properties for: $computerName") -Type "Info"

    $computer = $Script:Computers | Where-Object { $_.Name -eq $computerName } | Select-Object -First 1

    if (-not $computer) {
        Debug-Log (": Computer NOT found in Script:Computers for name: $computerName") -Type "Info"
        Show-Modal "Not Found" "Computer '$computerName' not found"
        return
    }

    Debug-Log (": Computer found: $($computer.Name)") -Type "Info"

    try {
        # ---------------------- Buttons ----------------------
        $btnOK = [Terminal.Gui.Button]::new("OK")
        $btnCancel = [Terminal.Gui.Button]::new("Cancel")

        $btnOK.add_Clicked({
            Debug-Log ": OK clicked" -Type "Info"
            [Terminal.Gui.Application]::RequestStop()
        })

        $btnCancel.add_Clicked({
            Debug-Log ": Cancel clicked" -Type "Info"
            [Terminal.Gui.Application]::RequestStop()
        })

        # ---------------------- Dialog ----------------------
        $dlg = [Terminal.Gui.Dialog]::new("Computer Properties - $($computer.Name)", 100, 40, $btnOK, $btnCancel)

        # ---------------------- TabView ----------------------
        $tabView = [Terminal.Gui.TabView]::new()
        $tabView.X = 0
        $tabView.Y = 0
        $tabView.Width = [Terminal.Gui.Dim]::Fill()
        $tabView.Height = [Terminal.Gui.Dim]::Fill(1)

        # ==================== General Tab ====================
        Debug-Log ": Creating General tab" -Type "Info"
        $generalTab = [Terminal.Gui.TabView+Tab]::new()
        $generalTab.Text = "General"
        $generalView = [Terminal.Gui.View]::new()
        $generalView.X = 0; $generalView.Y = 0
        $generalView.Width = [Terminal.Gui.Dim]::Fill()
        $generalView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        $lbl = [Terminal.Gui.Label]::new("Computer Information"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Name:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $txtName = [Terminal.Gui.Label]::new($computer.Name ?? ""); $txtName.X=25; $txtName.Y=$y
        $generalView.Add($txtName); $y+=1

        $lbl = [Terminal.Gui.Label]::new("DNS Host Name:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $txtDNS = [Terminal.Gui.Label]::new($computer.DNSHostName ?? ""); $txtDNS.X=25; $txtDNS.Y=$y
        $generalView.Add($txtDNS); $y+=1

        $lbl = [Terminal.Gui.Label]::new("Domain:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $txtDomain = [Terminal.Gui.Label]::new($computer.Domain ?? ""); $txtDomain.X=25; $txtDomain.Y=$y
        $generalView.Add($txtDomain); $y+=1

        $lbl = [Terminal.Gui.Label]::new("Description:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $txtDescription = [Terminal.Gui.TextField]::new($computer.Description ?? ""); $txtDescription.X=25; $txtDescription.Y=$y; $txtDescription.Width=65
        $generalView.Add($txtDescription); $y+=2

        # Operating System
        $lbl = [Terminal.Gui.Label]::new("Operating System"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl); $y+=2

        $lbl = [Terminal.Gui.Label]::new("OS:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $txtOS = [Terminal.Gui.Label]::new($computer.OperatingSystem ?? ""); $txtOS.X=25; $txtOS.Y=$y
        $generalView.Add($txtOS); $y+=1

        $lbl = [Terminal.Gui.Label]::new("Version:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $txtOSVer = [Terminal.Gui.Label]::new($computer.OperatingSystemVersion ?? ""); $txtOSVer.X=25; $txtOSVer.Y=$y
        $generalView.Add($txtOSVer); $y+=1

        if ($computer.OperatingSystemServicePack) {
            $lbl = [Terminal.Gui.Label]::new("Service Pack:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
            $txtSP = [Terminal.Gui.Label]::new($computer.OperatingSystemServicePack); $txtSP.X=25; $txtSP.Y=$y
            $generalView.Add($txtSP); $y+=1
        }

        # Network
        $y+=1
        $lbl = [Terminal.Gui.Label]::new("Network"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl); $y+=2

        $lbl = [Terminal.Gui.Label]::new("IPv4 Address:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $txtIPv4 = [Terminal.Gui.Label]::new($computer.IPv4Address ?? ""); $txtIPv4.X=25; $txtIPv4.Y=$y
        $generalView.Add($txtIPv4); $y+=1

        if ($computer.IPv6Address) {
            $lbl = [Terminal.Gui.Label]::new("IPv6 Address:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
            $txtIPv6 = [Terminal.Gui.Label]::new($computer.IPv6Address); $txtIPv6.X=25; $txtIPv6.Y=$y
            $generalView.Add($txtIPv6); $y+=1
        }

        $lbl = [Terminal.Gui.Label]::new("Location:"); $lbl.X=4; $lbl.Y=$y; $generalView.Add($lbl)
        $txtLocation = [Terminal.Gui.TextField]::new($computer.Location ?? ""); $txtLocation.X=25; $txtLocation.Y=$y; $txtLocation.Width=65
        $generalView.Add($txtLocation); $y+=1

        $generalTab.View = $generalView
        $tabView.AddTab($generalTab, $false)

        # ==================== Account Tab ====================
        Debug-Log ": Creating Account tab" -Type "Info"
        $accountTab = [Terminal.Gui.TabView+Tab]::new()
        $accountTab.Text = "Account"
        $accountView = [Terminal.Gui.View]::new()
        $accountView.X = 0; $accountView.Y = 0
        $accountView.Width = [Terminal.Gui.Dim]::Fill()
        $accountView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        $lbl = [Terminal.Gui.Label]::new("Account Status"); $lbl.X=2; $lbl.Y=$y; $accountView.Add($lbl); $y+=2

        $chkEnabled = [Terminal.Gui.CheckBox]::new("Computer Account Enabled"); $chkEnabled.X=4; $chkEnabled.Y=$y; $chkEnabled.Checked=($computer.Enabled??$true)
        $accountView.Add($chkEnabled); $y+=1

        $chkLocked = [Terminal.Gui.CheckBox]::new("Account Locked"); $chkLocked.X=4; $chkLocked.Y=$y; $chkLocked.Checked=($computer.LockedOut??$false); $chkLocked.Enabled=$false
        $accountView.Add($chkLocked); $y+=2

        # Password Settings
        $lbl = [Terminal.Gui.Label]::new("Password Settings"); $lbl.X=2; $lbl.Y=$y; $accountView.Add($lbl); $y+=2

        $chkPasswordExpired = [Terminal.Gui.CheckBox]::new("Password Expired"); $chkPasswordExpired.X=4; $chkPasswordExpired.Y=$y; $chkPasswordExpired.Checked=($computer.PasswordExpired??$false); $chkPasswordExpired.Enabled=$false
        $accountView.Add($chkPasswordExpired); $y+=1

        $chkPasswordNeverExpires = [Terminal.Gui.CheckBox]::new("Password never expires"); $chkPasswordNeverExpires.X=4; $chkPasswordNeverExpires.Y=$y; $chkPasswordNeverExpires.Checked=($computer.PasswordNeverExpires??$false); $chkPasswordNeverExpires.Enabled=$false
        $accountView.Add($chkPasswordNeverExpires); $y+=1

        $chkCannotChangePassword = [Terminal.Gui.CheckBox]::new("Cannot change password"); $chkCannotChangePassword.X=4; $chkCannotChangePassword.Y=$y; $chkCannotChangePassword.Checked=($computer.CannotChangePassword??$false); $chkCannotChangePassword.Enabled=$false
        $accountView.Add($chkCannotChangePassword); $y+=1

        $chkPasswordNotRequired = [Terminal.Gui.CheckBox]::new("Password not required"); $chkPasswordNotRequired.X=4; $chkPasswordNotRequired.Y=$y; $chkPasswordNotRequired.Checked=($computer.PasswordNotRequired??$false); $chkPasswordNotRequired.Enabled=$false
        $accountView.Add($chkPasswordNotRequired); $y+=2

        # Account Details
        $lbl = [Terminal.Gui.Label]::new("Account Details"); $lbl.X=2; $lbl.Y=$y; $accountView.Add($lbl); $y+=2

        $lbl = [Terminal.Gui.Label]::new("SAM Account:"); $lbl.X=4; $lbl.Y=$y; $accountView.Add($lbl)
        $txtSAM = [Terminal.Gui.Label]::new($computer.SamAccountName ?? ""); $txtSAM.X=25; $txtSAM.Y=$y
        $accountView.Add($txtSAM); $y+=1

        if ($computer.PasswordLastSet) {
            $lbl = [Terminal.Gui.Label]::new("Password last set: " + $computer.PasswordLastSet.ToString('yyyy-MM-dd HH:mm')); $lbl.X=4; $lbl.Y=$y
            $accountView.Add($lbl); $y+=1
        }

        if ($computer.LastLogonDate) {
            $lbl = [Terminal.Gui.Label]::new("Last logon: " + $computer.LastLogonDate.ToString('yyyy-MM-dd HH:mm')); $lbl.X=4; $lbl.Y=$y
            $accountView.Add($lbl); $y+=1
        }

        if ($computer.logonCount) {
            $lbl = [Terminal.Gui.Label]::new("Logon count: " + $computer.logonCount); $lbl.X=4; $lbl.Y=$y
            $accountView.Add($lbl); $y+=1
        }

        if ($computer.AccountExpirationDate) {
            $lbl = [Terminal.Gui.Label]::new("Account expires: " + $computer.AccountExpirationDate.ToString('yyyy-MM-dd')); $lbl.X=4; $lbl.Y=$y
            $accountView.Add($lbl); $y+=1
        }

        $accountTab.View = $accountView
        $tabView.AddTab($accountTab, $false)

        # ==================== Security Tab ====================
        Debug-Log ": Creating Security tab" -Type "Info"
        $securityTab = [Terminal.Gui.TabView+Tab]::new()
        $securityTab.Text = "Security"
        $securityView = [Terminal.Gui.View]::new()
        $securityView.X = 0; $securityView.Y = 0
        $securityView.Width = [Terminal.Gui.Dim]::Fill()
        $securityView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        $lbl = [Terminal.Gui.Label]::new("Delegation Settings"); $lbl.X=2; $lbl.Y=$y; $securityView.Add($lbl); $y+=2

        $chkTrustedForDelegation = [Terminal.Gui.CheckBox]::new("Trusted for delegation"); $chkTrustedForDelegation.X=4; $chkTrustedForDelegation.Y=$y; $chkTrustedForDelegation.Checked=($computer.TrustedForDelegation??$false); $chkTrustedForDelegation.Enabled=$false
        $securityView.Add($chkTrustedForDelegation); $y+=1

        $chkTrustedToAuth = [Terminal.Gui.CheckBox]::new("Trusted to authenticate for delegation"); $chkTrustedToAuth.X=4; $chkTrustedToAuth.Y=$y; $chkTrustedToAuth.Checked=($computer.TrustedToAuthForDelegation??$false); $chkTrustedToAuth.Enabled=$false
        $securityView.Add($chkTrustedToAuth); $y+=1

        $chkAccountNotDelegated = [Terminal.Gui.CheckBox]::new("Account not delegated"); $chkAccountNotDelegated.X=4; $chkAccountNotDelegated.Y=$y; $chkAccountNotDelegated.Checked=($computer.AccountNotDelegated??$false); $chkAccountNotDelegated.Enabled=$false
        $securityView.Add($chkAccountNotDelegated); $y+=1

        $chkNoPreAuth = [Terminal.Gui.CheckBox]::new("Does not require Kerberos preauthentication"); $chkNoPreAuth.X=4; $chkNoPreAuth.Y=$y; $chkNoPreAuth.Checked=($computer.DoesNotRequirePreAuth??$false); $chkNoPreAuth.Enabled=$false
        $securityView.Add($chkNoPreAuth); $y+=1

        $chkUseDES = [Terminal.Gui.CheckBox]::new("Use DES encryption types"); $chkUseDES.X=4; $chkUseDES.Y=$y; $chkUseDES.Checked=($computer.UseDESKeyOnly??$false); $chkUseDES.Enabled=$false
        $securityView.Add($chkUseDES); $y+=2

        # Kerberos
        $lbl = [Terminal.Gui.Label]::new("Kerberos Encryption"); $lbl.X=2; $lbl.Y=$y; $securityView.Add($lbl); $y+=2

        if ($computer.KerberosEncryptionType) {
            $kerbTypes = $computer.KerberosEncryptionType -join ', '
            $lbl = [Terminal.Gui.Label]::new("Encryption types: " + $kerbTypes); $lbl.X=4; $lbl.Y=$y
            $securityView.Add($lbl); $y+=1
        }

        if ($computer.'msDS-SupportedEncryptionTypes') {
            $lbl = [Terminal.Gui.Label]::new("Supported encryption: " + $computer.'msDS-SupportedEncryptionTypes'); $lbl.X=4; $lbl.Y=$y
            $securityView.Add($lbl); $y+=1
        }

        # SID
        $y+=1
        $lbl = [Terminal.Gui.Label]::new("Identifiers"); $lbl.X=2; $lbl.Y=$y; $securityView.Add($lbl); $y+=2

        $sid = $computer.SID ?? $computer.objectSid
        if ($sid) {
            $lbl = [Terminal.Gui.Label]::new("SID:"); $lbl.X=4; $lbl.Y=$y; $securityView.Add($lbl)
            $txtSID = [Terminal.Gui.Label]::new($sid.ToString()); $txtSID.X=25; $txtSID.Y=$y
            $securityView.Add($txtSID); $y+=1
        }

        if ($computer.ObjectGUID) {
            $lbl = [Terminal.Gui.Label]::new("Object GUID:"); $lbl.X=4; $lbl.Y=$y; $securityView.Add($lbl)
            $txtGUID = [Terminal.Gui.Label]::new($computer.ObjectGUID.ToString()); $txtGUID.X=25; $txtGUID.Y=$y
            $securityView.Add($txtGUID); $y+=1
        }

        $securityTab.View = $securityView
        $tabView.AddTab($securityTab, $false)

        # ==================== Member Of Tab ====================
        Debug-Log ": Creating Member Of tab" -Type "Info"
        $memberTab = [Terminal.Gui.TabView+Tab]::new()
        $memberTab.Text = "Member Of"
        $memberView = [Terminal.Gui.View]::new()
        $memberView.X = 0; $memberView.Y = 0
        $memberView.Width = [Terminal.Gui.Dim]::Fill()
        $memberView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        $lbl = [Terminal.Gui.Label]::new("Group Memberships:"); $lbl.X=2; $lbl.Y=$y; $memberView.Add($lbl); $y+=2

        $lstGroups = [Terminal.Gui.ListView]::new()
        $lstGroups.X = 2
        $lstGroups.Y = $y
        $lstGroups.Width = [Terminal.Gui.Dim]::Fill(2)
        $lstGroups.Height = 20  # Fixed height to leave room for buttons inside tab

        $groupList = @()
        if ($computer.MemberOf) {
            $groupList = $computer.MemberOf | ForEach-Object {
                if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
            }
        }

        if ($groupList.Count -gt 0) {
            $lstGroups.SetSource($groupList)
        } else {
            $lstGroups.SetSource(@("(No group memberships)"))
        }

        $memberView.Add($lstGroups)

        # =====================================================
# SNIPPET: Add/Remove Buttons for Member Of Tab
# =====================================================
# Add this code after you create the ListView ($lstGroups)
# and before you set the tab view ($memberTab.View = $memberView)

$memberView.Add($lstGroups)

# Add/Remove buttons - FIXED POSITIONS
$btnAdd = [Terminal.Gui.Button]::new("Add to Group...")
$btnAdd.X = 2
$btnAdd.Y = 23  # Hardcoded position (after ListView height of 20 + label + spacing)
$btnAdd.add_Clicked({
    Show-EditGroupMembershipDialog -User $user -OnUpdate {
        # Refresh code here
    }
})
$memberView.Add($btnAdd)

$btnRemove = [Terminal.Gui.Button]::new("Remove from Group")
$btnRemove.X = 22
$btnRemove.Y = 23  # Same Y position as Add button
$btnRemove.add_Clicked({
    # Remove code here
})
$memberView.Add($btnRemove)
$memberTab.View = $memberView

        # ==================== Service Principal Names Tab ====================
        Debug-Log ": Creating SPN tab" -Type "Info"
        $spnTab = [Terminal.Gui.TabView+Tab]::new()
        $spnTab.Text = "SPNs"
        $spnView = [Terminal.Gui.View]::new()
        $spnView.X = 0; $spnView.Y = 0
        $spnView.Width = [Terminal.Gui.Dim]::Fill()
        $spnView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        $lbl = [Terminal.Gui.Label]::new("Service Principal Names:"); $lbl.X=2; $lbl.Y=$y; $spnView.Add($lbl); $y+=2

        $lstSPNs = [Terminal.Gui.ListView]::new()
        $lstSPNs.X = 2
        $lstSPNs.Y = $y
        $lstSPNs.Width = [Terminal.Gui.Dim]::Fill(2)
        $lstSPNs.Height = [Terminal.Gui.Dim]::Fill(2)

        $spnList = @()
        if ($computer.ServicePrincipalNames) {
            $spnList = $computer.ServicePrincipalNames
        } elseif ($computer.servicePrincipalName) {
            $spnList = $computer.servicePrincipalName
        }

        if ($spnList.Count -gt 0) {
            $lstSPNs.SetSource($spnList)
        } else {
            $lstSPNs.SetSource(@("(No SPNs configured)"))
        }

        $spnView.Add($lstSPNs)

        $spnTab.View = $spnView
        $tabView.AddTab($spnTab, $false)

        # ==================== LAPS Tab ====================
        Debug-Log ": Creating LAPS tab" -Type "Info"
        $lapsTab = [Terminal.Gui.TabView+Tab]::new()
        $lapsTab.Text = "LAPS"
        $lapsView = [Terminal.Gui.View]::new()
        $lapsView.X = 0; $lapsView.Y = 0
        $lapsView.Width = [Terminal.Gui.Dim]::Fill()
        $lapsView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        $lbl = [Terminal.Gui.Label]::new("Local Administrator Password Solution"); $lbl.X=2; $lbl.Y=$y; $lapsView.Add($lbl); $y+=2

        # Check for LAPS in production mode
        $lapsEnabled = $false
        $lapsPassword = ""
        $lapsExpiry = ""

        if (-not $Script:DemoMode) {
            try {
                $lapsData = Get-ADComputer -Identity $computer.Name -Properties 'ms-Mcs-AdmPwd','ms-Mcs-AdmPwdExpirationTime' -ErrorAction SilentlyContinue
                if ($lapsData.'ms-Mcs-AdmPwd') {
                    $lapsEnabled = $true
                    $lapsPassword = $lapsData.'ms-Mcs-AdmPwd'
                    $lapsExpiry = if ($lapsData.'ms-Mcs-AdmPwdExpirationTime') {
                        [DateTime]::FromFileTime($lapsData.'ms-Mcs-AdmPwdExpirationTime').ToString('yyyy-MM-dd HH:mm:ss')
                    } else {
                        "Unknown"
                    }
                }
            } catch {
                Debug-Log (": Failed to retrieve LAPS info: $_") -Type "Warn"
            }
        } elseif ($computer.'ms-Mcs-AdmPwd') {
            # Demo mode with LAPS data
            $lapsEnabled = $true
            $lapsPassword = $computer.'ms-Mcs-AdmPwd'
            $lapsExpiry = if ($computer.'ms-Mcs-AdmPwdExpirationTime') {
                [DateTime]::FromFileTime($computer.'ms-Mcs-AdmPwdExpirationTime').ToString('yyyy-MM-dd HH:mm:ss')
            } else {
                "Unknown"
            }
        }

        if ($lapsEnabled) {
            $lbl = [Terminal.Gui.Label]::new("LAPS Status: ✓ Enabled"); $lbl.X=4; $lbl.Y=$y; $lapsView.Add($lbl); $y+=2

            $lbl = [Terminal.Gui.Label]::new("Administrator Password:"); $lbl.X=4; $lbl.Y=$y; $lapsView.Add($lbl); $y+=1
            $txtPassword = [Terminal.Gui.TextField]::new($lapsPassword); $txtPassword.X=4; $txtPassword.Y=$y; $txtPassword.Width=70; $txtPassword.ReadOnly=$true
            $lapsView.Add($txtPassword); $y+=2

            $lbl = [Terminal.Gui.Label]::new("Password Expires:"); $lbl.X=4; $lbl.Y=$y; $lapsView.Add($lbl)
            $txtExpiry = [Terminal.Gui.Label]::new($lapsExpiry); $txtExpiry.X=25; $txtExpiry.Y=$y
            $lapsView.Add($txtExpiry); $y+=2

            $lbl = [Terminal.Gui.Label]::new("⚠ This password provides full local administrator access"); $lbl.X=4; $lbl.Y=$y; $lapsView.Add($lbl); $y+=1
            $lbl = [Terminal.Gui.Label]::new("  Handle with appropriate security controls"); $lbl.X=4; $lbl.Y=$y; $lapsView.Add($lbl)
        } else {
            $lbl = [Terminal.Gui.Label]::new("LAPS Status: ✗ Not Enabled"); $lbl.X=4; $lbl.Y=$y; $lapsView.Add($lbl); $y+=2
            $lbl = [Terminal.Gui.Label]::new("This computer does not have LAPS configured or you"); $lbl.X=4; $lbl.Y=$y; $lapsView.Add($lbl); $y+=1
            $lbl = [Terminal.Gui.Label]::new("do not have permissions to view the password."); $lbl.X=4; $lbl.Y=$y; $lapsView.Add($lbl)
        }

        $lapsTab.View = $lapsView
        $tabView.AddTab($lapsTab, $false)

        # ==================== Advanced Tab ====================
        Debug-Log ": Creating Advanced tab" -Type "Info"
        $advancedTab = [Terminal.Gui.TabView+Tab]::new()
        $advancedTab.Text = "Advanced"
        $advancedView = [Terminal.Gui.View]::new()
        $advancedView.X = 0; $advancedView.Y = 0
        $advancedView.Width = [Terminal.Gui.Dim]::Fill()
        $advancedView.Height = [Terminal.Gui.Dim]::Fill()

        $y = 1
        $lbl = [Terminal.Gui.Label]::new("Distinguished Name:"); $lbl.X=2; $lbl.Y=$y; $advancedView.Add($lbl); $y+=1
        $txtDN = [Terminal.Gui.TextField]::new($computer.DistinguishedName ?? ""); $txtDN.X=2; $txtDN.Y=$y; $txtDN.Width=90; $txtDN.ReadOnly=$true
        $advancedView.Add($txtDN); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Canonical Name:"); $lbl.X=2; $lbl.Y=$y; $advancedView.Add($lbl); $y+=1
        $txtCN = [Terminal.Gui.TextField]::new($computer.CanonicalName ?? ""); $txtCN.X=2; $txtCN.Y=$y; $txtCN.Width=90; $txtCN.ReadOnly=$true
        $advancedView.Add($txtCN); $y+=2

        # Timestamps
        $lbl = [Terminal.Gui.Label]::new("Timestamps"); $lbl.X=2; $lbl.Y=$y; $advancedView.Add($lbl); $y+=2

        $created = if ($computer.Created -or $computer.whenCreated) {
            ($computer.Created ?? $computer.whenCreated).ToString('yyyy-MM-dd HH:mm:ss')
        } else { "Unknown" }
        $lbl = [Terminal.Gui.Label]::new("Created: " + $created); $lbl.X=4; $lbl.Y=$y; $advancedView.Add($lbl); $y+=1

        $modified = if ($computer.Modified -or $computer.whenChanged) {
            ($computer.Modified ?? $computer.whenChanged).ToString('yyyy-MM-dd HH:mm:ss')
        } else { "Unknown" }
        $lbl = [Terminal.Gui.Label]::new("Modified: " + $modified); $lbl.X=4; $lbl.Y=$y; $advancedView.Add($lbl); $y+=1

        if ($computer.LastBootUpTime) {
            $lbl = [Terminal.Gui.Label]::new("Last Boot: " + $computer.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss')); $lbl.X=4; $lbl.Y=$y
            $advancedView.Add($lbl); $y+=1
        }

        # USN
        $y+=1
        $lbl = [Terminal.Gui.Label]::new("Update Sequence Numbers"); $lbl.X=2; $lbl.Y=$y; $advancedView.Add($lbl); $y+=2

        if ($computer.uSNCreated) {
            $lbl = [Terminal.Gui.Label]::new("USN Created: " + $computer.uSNCreated); $lbl.X=4; $lbl.Y=$y
            $advancedView.Add($lbl); $y+=1
        }

        if ($computer.uSNChanged) {
            $lbl = [Terminal.Gui.Label]::new("USN Changed: " + $computer.uSNChanged); $lbl.X=4; $lbl.Y=$y
            $advancedView.Add($lbl); $y+=1
        }

        # Additional properties
        $y+=1
        $lbl = [Terminal.Gui.Label]::new("Additional Properties"); $lbl.X=2; $lbl.Y=$y; $advancedView.Add($lbl); $y+=2

        if ($computer.isCriticalSystemObject) {
            $lbl = [Terminal.Gui.Label]::new("✓ Critical System Object"); $lbl.X=4; $lbl.Y=$y
            $advancedView.Add($lbl); $y+=1
        }

        if ($computer.ProtectedFromAccidentalDeletion) {
            $lbl = [Terminal.Gui.Label]::new("✓ Protected from Accidental Deletion"); $lbl.X=4; $lbl.Y=$y
            $advancedView.Add($lbl); $y+=1
        }

        if ($computer.PrimaryGroup) {
            $lbl = [Terminal.Gui.Label]::new("Primary Group: " + $computer.PrimaryGroup); $lbl.X=4; $lbl.Y=$y
            $advancedView.Add($lbl); $y+=1
        }

        $advancedTab.View = $advancedView
        $tabView.AddTab($advancedTab, $false)

        # Add TabView to dialog
        $dlg.Add($tabView)

        Debug-Log ": All tabs added, running dialog" -Type "Success"

        # Run the dialog
        [Terminal.Gui.Application]::Run($dlg)
        Debug-Log ": Dialog closed normally" -Type "Info"

    } catch {
        Debug-Log ": Exception in Show-ComputerPropertiesDialog: $($_.Exception.Message)" -Type "Error"
        Debug-Log ": Stack trace: $($_.ScriptStackTrace)" -Type "Error"
        throw
    }
}

# ------------------------- Context Menu Handler ------------------------
function Show-ContextMenu {
    param(
        [string]$objectName,
        [string]$objectType
    )

$selName = $Script:tree.SelectedObject.Text
Debug-Log (": Selected object text: '$selName'") -Type "Info"

# Determine what type of object this is FIRST (before cleaning the name)
$selType = if ($selName -like "(U)*") {"user"}
           elseif ($selName -like "(DC)*") {"dc"}
           else {"group"}
Debug-Log (": Detected type: $selType") -Type "Info"

# NOW clean the name - remove prefixes like "(U) " or "(DC) " and status icons
$cleanName = $selName -replace '^\([^)]+\)\s*', '' -replace '^[○⊗🔒]\s*', ''
Debug-Log (": Cleaned name after removing prefix: '$cleanName'") -Type "Info"

# Extract just the name part if it has [SITE] suffix
if ($cleanName -match '^(.+?)\s+\[.+\]$') {
    $cleanName = $matches[1].Trim()
    Debug-Log (": Extracted name from [SITE] format: '$cleanName'") -Type "Info"
}

Debug-Log (": Final cleaned name: '$cleanName'") -Type "Info"

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

    Debug-Log (": Menu item selected: $selected") -Type "Info"

    [Terminal.Gui.Application]::RequestStop()

    if ($selected -ne "---") {
        Debug-Log (": Context menu selected: $selected") -Type "Info"

        switch ($selected) {
            "Properties" {
                Debug-Log (": Properties selected for type: isUser=$isUser, isGroup=$isGroup, isDC=$isDC") -Type "Info"

                if ($isUser) {
                    Debug-Log (": Looking for user: $cleanName") -Type "Info"
                    $user = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1

                    if ($user) {
                        Debug-Log (": Found user, calling Show-UserPropertiesDialog") -Type "Info"
                        Debug-Log (": User type: $($user.GetType().Name)") -Type "Info"
                        Debug-Log (": User Name: $($user.Name)") -Type "Info"

                        try {
                            Show-UserPropertiesDialog -user $user  # REMOVED -Global $Script
                            Debug-Log (": Show-UserPropertiesDialog returned") -Type "Info"
                        } catch {
                            Debug-Log (": Exception showing properties: $($_.Exception.Message)") -Type "Warn"
                            Debug-Log (": Stack trace: $($_.ScriptStackTrace)") -Type "Warn"
                            Show-Modal "Error" "Failed to show properties:`n$($_.Exception.Message)"
                        }
                    } else {
                        Debug-Log (": User '$cleanName' not found in Global:Users") -Type "Warn"
                        Show-Modal "Not Found" "User '$cleanName' not found"
                    }
                } elseif ($isGroup) {
                    Debug-Log (": Showing group properties") -Type "Info"
                    Show-GroupPropertiesDialog -groupName $cleanName
                } elseif ($isDC) {
                    Debug-Log (": Showing DC properties") -Type "Info"
                    Show-DCPropertiesDialog -dcName $cleanName
                } else {
                    Debug-Log (": Showing generic properties") -Type "Info"
                    $msg = "Object: $cleanName`nType: $objectType"
                    Show-Modal "Properties" $msg
                }
            }
            "Reset Password" { Show-ResetPasswordDialog -userName $cleanName }
            "Disable Account" { Toggle-UserAccount -userName $cleanName -disable $true }
            "Enable Account" { Toggle-UserAccount -userName $cleanName -disable $false }
            "Move to OU..." { Show-MoveObjectDialog -objectName $cleanName -objectType "User" }
            "Delete" { Show-DeleteObjectDialog -objectName $cleanName -objectType $objectType }
            "Add Member..." { Show-EditGroupMembershipDialog -groupName $cleanName }
            "Remove Member..." { Show-EditGroupMembershipDialog -groupName $cleanName }
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

function Show-Properties {
    Debug-Log (": Show-Properties called") -Type "Info"

    if (-not $Script:tree.SelectedObject) {
        Debug-Log (": No object selected") -Type "Info"
        Show-Modal "Debug" "No object selected in tree"
        return
    }

    $selName = $Script:tree.SelectedObject.Text
    Debug-Log (": Selected object text: '$selName'") -Type "Info"

    # Parse the tree text to get clean name and type
    $objInfo = Get-CleanObjectInfo -treeText $selName
    $cleanName = $objInfo.Name
    $selType = $objInfo.Type

    # Route to appropriate dialog
    if ($selType -eq "user") {
        $selUser = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1

        if ($selUser) {
            Debug-Log (": User found, calling Show-UserPropertiesDialog") -Type "Info"
            try {
                Show-UserPropertiesDialog -user $selUser
            } catch {
                Debug-Log (": Exception in Show-UserPropertiesDialog: $_") -Type "Info"
                Show-Modal "Error" "Failed to show properties`n$($_.Exception.Message)`n`nCheck console for details"
            }
        } else {
            Debug-Log (": User NOT found in Script:Users") -Type "Info"
            Show-Modal "Not Found" "User '$cleanName' not found"
        }
    } elseif ($selType -eq "group") {
        Debug-Log (": Group type selected: $cleanName") -Type "Info"
        Show-GroupPropertiesDialog -groupName $cleanName
    } elseif ($selType -eq "ou") {
        Debug-Log (": OU type selected: $cleanName") -Type "Info"
        Show-OUPropertiesDialog -ouName $cleanName
    } elseif ($selType -eq "dc") {
        Debug-Log (": DC type selected: $cleanName") -Type "Info"

        ## Find the DC object first
        $dcObj = $Script:DCs | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1

        if ($dcObj) {
            Debug-Log (": DC found: $($dcObj.Name)") -Type "Info"
            Show-DCPropertiesDialog -dc $dcObj  # Pass the object
        } else {
            Debug-Log (": DC NOT found in Script:DCs for name: $cleanName") -Type "Info"
            Show-Modal "Error" "DC '$cleanName' not found"
        }
    } elseif ($selType -eq "computer") {
        Debug-Log (": Computer type selected: $cleanName") -Type "Info"
        Show-ComputerPropertiesDialog -computerName $cleanName
    } else {
        Debug-Log (": Selected object type '$selType' not handled yet") -Type "Info"
    }
}

# Additional debug helper - call this to verify your demo data loaded correctly
function Test-DemoData {
    Debug-Log ("========== DEMO DATA CHECK ==========") -Type "Info"
    Debug-Log ("Global:Users count: $($Script:Users.Count)") -Type "Info"
    Debug-Log ("Global:DCs count: $($Script:DCs.Count)") -Type "Info"
    Debug-Log ("") -Type "Info"
    Debug-Log ("Users in memory:") -Type "Info"
    foreach ($u in $Script:Users) {
        $locked = if ($u.Locked) { "🔒" } else { "" }
        $disabled = if ($u.Disabled) { "⊗" } else { "○" }
        Debug-Log ("  $disabled$locked $($u.Name) - Groups: $($u.Groups -join ', ')") -Type "Info"
    }
    Debug-Log ("=====================================") -Type "Info"
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
            if ($Script:DemoMode) {
                Debug-Log (": Password reset for $userName (demo mode)") -Type "Info"
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
            if ($Script:DemoMode) {
                $user = $Script:Users | Where-Object { $_.Name -eq $userName } | Select-Object -First 1
                if ($user) {
                    $user.Disabled = $disable
                    Debug-Log (": Account $userName $action`d (demo mode)") -Type "Info"
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
$Script:SelectedObjects = @()
$Script:SelectionMode = $false

# ------------------------- Toggle Selection Mode ------------------------
function Toggle-SelectionMode {
    $Script:SelectionMode = -not $Script:SelectionMode

    if ($Script:SelectionMode) {
        Debug-Log (": Selection mode ENABLED") -Type "Info"
        [Terminal.Gui.MessageBox]::Query(60, 8, "Selection Mode",
            "Selection mode enabled!`n`nClick objects to select/deselect them.`nPress Ctrl+A to select all.`nPress Ctrl+D to deselect all.",
            "OK") | Out-Null
    } else {
        Debug-Log (": Selection mode DISABLED") -Type "Info"
        $Script:SelectedObjects = @()
        Build-Tree -domain $Script:CurrentDomain
        Update-FilterStatusLabel -label $Script:FilterStatusLabel
    }
}

# ------------------------- Update Selection Panel ------------------------
function Update-SelectionPanel {
    param($panel)

    if (-not $panel -or -not $panel.Tag) { return }

    $lblCount = $panel.Tag.CountLabel
    $lstSelected = $panel.Tag.ListView

    $count = $Script:SelectedObjects.Count
    $lblCount.Text = "$count object(s) selected"

    $displayNames = $Script:SelectedObjects | ForEach-Object {
        $name = $_ -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
        $name
    }

    $lstSelected.SetSource($displayNames)
    $panel.SetNeedsDisplay()
}

# ------------------------- Enhanced Tree with Selection Support ------------------------

function Handle-TreeClick {
    param($mouseArgs)

    if (-not $Script:tree.SelectedObject) { return }

    $selName = $Script:tree.SelectedObject.Text

    # Check if in selection mode
    if ($Script:SelectionMode) {
        # Toggle selection
        if ($Script:SelectedObjects -contains $selName) {
            # Deselect
            $Script:SelectedObjects = $Script:SelectedObjects | Where-Object { $_ -ne $selName }
            Debug-Log (": Deselected $selName") -Type "Info"
        } else {
            # Select
            $Script:SelectedObjects += $selName
            Debug-Log (": Selected $selName") -Type "Info"
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
    if (-not $Script:SelectionMode) {
      Show-Modal "Selection Mode" "Enable selection mode first (Ctrl+S)"
        return
    }

    $Script:SelectedObjects = @()

    # Get all users from tree
    foreach ($user in $Script:Users) {
        $statusIcon = if ($user.Disabled) { "⊗" } else { "○" }
        $displayName = "(U) $statusIcon $($user.Name)"
        $Script:SelectedObjects += $displayName
    }

    Debug-Log (": Selected all users ($($Script:SelectedObjects.Count))") -Type "Info"
    Update-SelectionPanel -panel $selectionPanel
    Show-Modal "Selected All" "Selected $($Script:SelectedObjects.Count) users"
}

function Deselect-AllObjects {
    $Script:SelectedObjects = @()
    Debug-Log (": Deselected all objects") -Type "Info"
    Update-SelectionPanel -panel $selectionPanel
}

# ------------------------- Bulk Disable/Enable ------------------------
function Invoke-BulkDisableEnable {
    param([bool]$disable)

    if ($Script:SelectedObjects.Count -eq 0) {
      Show-Modal "No Selection" "No objects selected. Select objects first."
        return
    }

    $action = if ($disable) { "disable" } else { "enable" }
    $result = [Terminal.Gui.MessageBox]::Query(60, 9, "Confirm Bulk Action",
        "Are you sure you want to $action $($Script:SelectedObjects.Count) user account(s)?",
        "Yes", "No")

    if ($result -eq 0) {
        $successCount = 0
        $failCount = 0
        $errors = @()

        foreach ($objName in $Script:SelectedObjects) {
            $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''

            try {
                if ($Script:DemoMode) {
                    $user = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                    if ($user) {
                        $user.Disabled = $disable
                        $successCount++
                        Debug-Log (": $action`d $cleanName (demo mode)") -Type "Info"
                    }
                } else {
                    if ($disable) {
                        Disable-ADAccount -Identity $cleanName -ErrorAction Stop
                    } else {
                        Enable-ADAccount -Identity $cleanName -ErrorAction Stop
                    }
                    $successCount++
                    Debug-Log (": $action`d $cleanName in AD") -Type "Info"
                }
            } catch {
                $failCount++
                $errors += "$cleanName`: $($_.Exception.Message)"
                Debug-Log (": Failed to $action $cleanName`: $_") -Type "Info"
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
        if (-not $Script:DemoMode) {
            Load-DomainData -domain $Script:CurrentDomain
        }
        Build-Tree -domain $Script:CurrentDomain
        Update-FilterStatusLabel -label $Script:FilterStatusLabel

        # Clear selection
        $Script:SelectedObjects = @()
        $Script:SelectionMode = $false
        Update-SelectionPanel -panel $selectionPanel
    }
}

# ------------------------- Bulk Move ------------------------
function Invoke-BulkMove {
    if ($Script:SelectedObjects.Count -eq 0) {
      Show-Modal "No Selection" "No objects selected. Select objects first."
        return
    }

    $dlg = [Terminal.Gui.Dialog]::new("Bulk Move - $($Script:SelectedObjects.Count) Objects", 70, 18)

    $lblInfo = [Terminal.Gui.Label]::new("Moving $($Script:SelectedObjects.Count) object(s) to:")
    $lblInfo.X=2; $lblInfo.Y=1; $dlg.Add($lblInfo)

    # Get list of OUs
    $ouList = if ($Script:DemoMode) {
        $Script:Users | Select-Object -ExpandProperty OU -Unique | Sort-Object
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
            "Move $($Script:SelectedObjects.Count) object(s) to:`n$targetOU?",
            "Yes", "No")

        if ($confirm -eq 0) {
            $successCount = 0
            $failCount = 0
            $errors = @()

            foreach ($objName in $Script:SelectedObjects) {
                $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''

                try {
                    if ($Script:DemoMode) {
                        $user = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                        if ($user) {
                            $user.OU = $targetOU
                            $successCount++
                            Debug-Log (": Moved $cleanName to $targetOU (demo mode)") -Type "Info"
                        }
                    } else {
                        $adObject = Get-ADObject -Filter "Name -eq '$cleanName'" -ErrorAction Stop
                        Move-ADObject -Identity $adObject.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
                        $successCount++
                        Debug-Log (": Moved $cleanName to $targetOU in AD") -Type "Info"
                    }
                } catch {
                    $failCount++
                    $errors += "$cleanName`: $($_.Exception.Message)"
                    Debug-Log (": Failed to move $cleanName`: $_") -Type "Info"
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
            if (-not $Script:DemoMode) {
                Load-DomainData -domain $Script:CurrentDomain
            }
            Build-Tree -domain $Script:CurrentDomain
            Update-FilterStatusLabel -label $Script:FilterStatusLabel

            # Clear selection
            $Script:SelectedObjects = @()
            $Script:SelectionMode = $false
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
    if ($Script:SelectedObjects.Count -eq 0) {
      Show-Modal "No Selection" "No objects selected. Select objects first."
        return
    }

    $dlg = [Terminal.Gui.Dialog]::new("Bulk Add to Group", 70, 18)

    $lblInfo = [Terminal.Gui.Label]::new("Add $($Script:SelectedObjects.Count) user(s) to group:")
    $lblInfo.X=2; $lblInfo.Y=1; $dlg.Add($lblInfo)

    # Get list of groups
    $groupList = if ($Script:DemoMode) {
        $allGroups = @()
        foreach ($u in $Script:Users) {
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
            "Add $($Script:SelectedObjects.Count) user(s) to group:`n$targetGroup?",
            "Yes", "No")

        if ($confirm -eq 0) {
            $successCount = 0
            $failCount = 0

            foreach ($objName in $Script:SelectedObjects) {
                $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''

                try {
                    if ($Script:DemoMode) {
                        $user = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                        if ($user -and $user.Groups -notcontains $targetGroup) {
                            $user.Groups += $targetGroup
                            $successCount++
                            Debug-Log (": Added $cleanName to $targetGroup (demo mode)") -Type "Info"
                        }
                    } else {
                        Add-ADGroupMember -Identity $targetGroup -Members $cleanName -ErrorAction Stop
                        $successCount++
                        Debug-Log (": Added $cleanName to $targetGroup in AD") -Type "Info"
                    }
                } catch {
                    $failCount++
                    Debug-Log (": Failed to add $cleanName`: $_") -Type "Info"
                }
            }

            [Terminal.Gui.MessageBox]::Query(60, 10, "Bulk Add Complete",
                "Successfully added $successCount user(s)`nFailed: $failCount",
                "OK") | Out-Null

            # Refresh tree
            Build-Tree -domain $Script:CurrentDomain
            Update-FilterStatusLabel -label $Script:FilterStatusLabel

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
    if (-not $Script:SelectionMode) {
        [Terminal.Gui.MessageBox]::Query(50, 7, "Selection Mode", "Enable selection mode first (Ctrl+S)", "OK") | Out-Null
        return
    }

    $Script:SelectedObjects = @()

    # Get all users from tree
    foreach ($user in $Script:Users) {
        $statusIcon = if ($user.Disabled) { "⊗" } else { "○" }
        $displayName = "(U) $statusIcon $($user.Name)"
        $Script:SelectedObjects += $displayName
    }

    Debug-Log (": Selected all users ($($Script:SelectedObjects.Count))") -Type "Info"
    Update-SelectionPanel -panel $selectionPanel
    Show-Modal "Selected All" "Selected $($Script:SelectedObjects.Count) users"
}

function Deselect-AllObjects {
    $Script:SelectedObjects = @()
    Debug-Log (": Deselected all objects") -Type "Info"
    Update-SelectionPanel -panel $selectionPanel
}

# ------------------------- Bulk Disable/Enable ------------------------
function Invoke-BulkDisableEnable {
    param([bool]$disable)

    if ($Script:SelectedObjects.Count -eq 0) {
      Show-Modal "No Selection" "No objects selected. Select objects first."
        return
    }

    $action = if ($disable) { "disable" } else { "enable" }
    $result = [Terminal.Gui.MessageBox]::Query(60, 9, "Confirm Bulk Action",
        "Are you sure you want to $action $($Script:SelectedObjects.Count) user account(s)?",
        "Yes", "No")

    if ($result -eq 0) {
        $successCount = 0
        $failCount = 0
        $errors = @()

        foreach ($objName in $Script:SelectedObjects) {
            $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''

            try {
                if ($Script:DemoMode) {
                    $user = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                    if ($user) {
                        $user.Disabled = $disable
                        $successCount++
                        Debug-Log (": $action`d $cleanName (demo mode)") -Type "Info"
                    }
                } else {
                    if ($disable) {
                        Disable-ADAccount -Identity $cleanName -ErrorAction Stop
                    } else {
                        Enable-ADAccount -Identity $cleanName -ErrorAction Stop
                    }
                    $successCount++
                    Debug-Log (": $action`d $cleanName in AD") -Type "Info"
                }
            } catch {
                $failCount++
                $errors += "$cleanName`: $($_.Exception.Message)"
                Debug-Log (": Failed to $action $cleanName`: $_") -Type "Info"
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
        if (-not $Script:DemoMode) {
            Load-DomainData -domain $Script:CurrentDomain
        }
        Build-Tree -domain $Script:CurrentDomain
        Update-FilterStatusLabel -label $Script:FilterStatusLabel

        # Clear selection
        $Script:SelectedObjects = @()
        $Script:SelectionMode = $false
        Update-SelectionPanel -panel $selectionPanel
    }
}

# ------------------------- Bulk Move ------------------------
function Invoke-BulkMove {
    if ($Script:SelectedObjects.Count -eq 0) {
      Show-Modal "No Selection" "No objects selected. Select objects first."
        return
    }

    $dlg = [Terminal.Gui.Dialog]::new("Bulk Move - $($Script:SelectedObjects.Count) Objects", 70, 18)

    $lblInfo = [Terminal.Gui.Label]::new("Moving $($Script:SelectedObjects.Count) object(s) to:")
    $lblInfo.X=2; $lblInfo.Y=1; $dlg.Add($lblInfo)

    # Get list of OUs
    $ouList = if ($Script:DemoMode) {
        $Script:Users | Select-Object -ExpandProperty OU -Unique | Sort-Object
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
            "Move $($Script:SelectedObjects.Count) object(s) to:`n$targetOU?",
            "Yes", "No")

        if ($confirm -eq 0) {
            $successCount = 0
            $failCount = 0
            $errors = @()

            foreach ($objName in $Script:SelectedObjects) {
                $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''

                try {
                    if ($Script:DemoMode) {
                        $user = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                        if ($user) {
                            $user.OU = $targetOU
                            $successCount++
                            Debug-Log (": Moved $cleanName to $targetOU (demo mode)") -Type "Info"
                        }
                    } else {
                        $adObject = Get-ADObject -Filter "Name -eq '$cleanName'" -ErrorAction Stop
                        Move-ADObject -Identity $adObject.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
                        $successCount++
                        Debug-Log (": Moved $cleanName to $targetOU in AD") -Type "Info"
                    }
                } catch {
                    $failCount++
                    $errors += "$cleanName`: $($_.Exception.Message)"
                    Debug-Log (": Failed to move $cleanName`: $_") -Type "Info"
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
            if (-not $Script:DemoMode) {
                Load-DomainData -domain $Script:CurrentDomain
            }
            Build-Tree -domain $Script:CurrentDomain
            Update-FilterStatusLabel -label $Script:FilterStatusLabel

            # Clear selection
            $Script:SelectedObjects = @()
            $Script:SelectionMode = $false
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
    if ($Script:SelectedObjects.Count -eq 0) {
      Show-Modal "No Selection" "No objects selected. Select objects first."
        return
    }

    $dlg = [Terminal.Gui.Dialog]::new("Bulk Add to Group", 70, 18)

    $lblInfo = [Terminal.Gui.Label]::new("Add $($Script:SelectedObjects.Count) user(s) to group:")
    $lblInfo.X=2; $lblInfo.Y=1; $dlg.Add($lblInfo)

    # Get list of groups
    $groupList = if ($Script:DemoMode) {
        $allGroups = @()
        foreach ($u in $Script:Users) {
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
            "Add $($Script:SelectedObjects.Count) user(s) to group:`n$targetGroup?",
            "Yes", "No")

        if ($confirm -eq 0) {
            $successCount = 0
            $failCount = 0

            foreach ($objName in $Script:SelectedObjects) {
                $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''

                try {
                    if ($Script:DemoMode) {
                        $user = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                        if ($user -and $user.Groups -notcontains $targetGroup) {
                            $user.Groups += $targetGroup
                            $successCount++
                            Debug-Log (": Added $cleanName to $targetGroup (demo mode)") -Type "Info"
                        }
                    } else {
                        Add-ADGroupMember -Identity $targetGroup -Members $cleanName -ErrorAction Stop
                        $successCount++
                        Debug-Log (": Added $cleanName to $targetGroup in AD") -Type "Info"
                    }
                } catch {
                    $failCount++
                    Debug-Log (": Failed to add $cleanName`: $_") -Type "Info"
                }
            }

            [Terminal.Gui.MessageBox]::Query(60, 10, "Bulk Add Complete",
                "Successfully added $successCount user(s)`nFailed: $failCount",
                "OK") | Out-Null

            # Refresh tree
            Build-Tree -domain $Script:CurrentDomain
            Update-FilterStatusLabel -label $Script:FilterStatusLabel

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
    param(
        [int]$panelWidth = 30,
        [int]$panelHeight = 15
    )

    # Create frame view
    $selPanel = [Terminal.Gui.FrameView]::new("Selected Objects")
    $selPanel.Width  = $panelWidth
    $selPanel.Height = $panelHeight
    $selPanel.X = [Terminal.Gui.Pos]::AnchorEnd($panelWidth)  # Right-align
    $selPanel.Y = 1  # Just below top margin

    # Count label
    $lblCount = [Terminal.Gui.Label]::new("0 objects selected")
    $lblCount.X = 1
    $lblCount.Y = 0
    $selPanel.Add($lblCount)

    # ListView
    $lstSelected = [Terminal.Gui.ListView]::new(@())
    $lstSelected.X = 1
    $lstSelected.Y = 1
    $lstSelected.Width  = [Terminal.Gui.Dim]::Fill(2)  # margin on both sides
    $lstSelected.Height = [Terminal.Gui.Dim]::Fill(6)  # leaves space for buttons
    $selPanel.Add($lstSelected)

    # Store references in Tag
    $selPanel | Add-Member -MemberType NoteProperty -Name Tag -Value @{
        CountLabel = $lblCount
        ListView   = $lstSelected
    } -Force

    # ----------------- Batch action buttons -----------------
    $yPos = [Terminal.Gui.Pos]::Bottom($lstSelected) + 1

    $btnBulkDisable = [Terminal.Gui.Button]::new("Disable All")
    $btnBulkDisable.X = 1
    $btnBulkDisable.Y = $yPos
    $btnBulkDisable.add_Clicked({ Invoke-BulkDisableEnable -disable $true })
    $selPanel.Add($btnBulkDisable)

    $yPos = [Terminal.Gui.Pos]::Bottom($btnBulkDisable) + 1
    $btnBulkEnable = [Terminal.Gui.Button]::new("Enable All")
    $btnBulkEnable.X = 1
    $btnBulkEnable.Y = $yPos
    $btnBulkEnable.add_Clicked({ Invoke-BulkDisableEnable -disable $false })
    $selPanel.Add($btnBulkEnable)

    $yPos = [Terminal.Gui.Pos]::Bottom($btnBulkEnable) + 1
    $btnBulkMove = [Terminal.Gui.Button]::new("Move All...")
    $btnBulkMove.X = 1
    $btnBulkMove.Y = $yPos
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

    $count = $Script:SelectedObjects.Count
    $lblCount.Text = "$count object(s) selected"

    $displayNames = $Script:SelectedObjects | ForEach-Object {
        $name = $_ -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
        $name
    }

    $lstSelected.SetSource($displayNames)
    $panel.SetNeedsDisplay()
}

# Add/Remove Group Member Implementation
# =====================================================
# ADD / REMOVE OUP MEMBERS AKA EDIT
# =====================================================
# Usage Examples
# =====================================================
function Show-EditGroupMembershipDialog {
    <#
    .SYNOPSIS
    Shows a dialog to edit group memberships with checkboxes

    .PARAMETER User
    User object to edit group memberships for

    .PARAMETER OnUpdate
    Optional scriptblock to call after successful changes
    #>
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$User,

        [scriptblock]$OnUpdate
    )

    Debug-Log ": Edit group membership for user $($User.Name)" -Type "Info"

    # Get current user groups
    $currentGroups = @()
    if ($User.Groups) {
        $currentGroups = $User.Groups
    } elseif ($User.MemberOf) {
        $currentGroups = $User.MemberOf | ForEach-Object {
            if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
        }
    }

    # Get all available groups
    if ($Script:DemoMode) {
        $allGroups = $Script:Groups | Sort-Object -Property Name
    } else {
        try {
            $allGroups = Get-ADGroup -Filter * -Properties Description -ErrorAction Stop | Sort-Object -Property Name
        } catch {
            Show-Modal "Error" "Failed to retrieve groups:`n$($_.Exception.Message)"
            return
        }
    }

    if ($allGroups.Count -eq 0) {
        Show-Modal "No Groups" "No groups available in the directory"
        return
    }

    # Create dialog
    $dlg = [Terminal.Gui.Dialog]::new("Edit Group Membership - $($User.Name)", 90, 35)

    $lblInfo = [Terminal.Gui.Label]::new("Check/uncheck groups and click Apply to save changes:")
    $lblInfo.X = 2
    $lblInfo.Y = 1
    $dlg.Add($lblInfo)

    # Create scrollable view for checkboxes
    $scrollView = [Terminal.Gui.ScrollView]::new()
    $scrollView.X = 2
    $scrollView.Y = 3
    $scrollView.Width = [Terminal.Gui.Dim]::Fill(2)
    $scrollView.Height = [Terminal.Gui.Dim]::Fill(5)
    $scrollView.ShowVerticalScrollIndicator = $true

    # Track checkboxes and their corresponding groups
    $checkboxes = @{}
    $y = 0

    foreach ($group in $allGroups) {
        $groupName = $group.Name
        $isChecked = $currentGroups -contains $groupName

        $description = if ($group.Description) { " - $($group.Description)" } else { "" }
        $chk = [Terminal.Gui.CheckBox]::new("$groupName$description")
        $chk.X = 0
        $chk.Y = $y
        $chk.Checked = $isChecked

        # Store reference
        $checkboxes[$groupName] = $chk

        $scrollView.Add($chk)
        $y++
    }

    # Set content size for scrolling
    $scrollView.ContentSize = [Terminal.Gui.Size]::new(80, $y)
    $dlg.Add($scrollView)

    # Buttons
    $btnApply = [Terminal.Gui.Button]::new("Apply Changes")
    $btnApply.X = 2
    $btnApply.Y = [Terminal.Gui.Pos]::Bottom($scrollView) + 1
    $btnApply.add_Clicked({
        # Determine what changed
        $toAdd = @()
        $toRemove = @()

        foreach ($groupName in $checkboxes.Keys) {
            $isCurrentlyMember = $currentGroups -contains $groupName
            $isChecked = $checkboxes[$groupName].Checked

            if ($isChecked -and -not $isCurrentlyMember) {
                # Need to add
                $toAdd += $groupName
            } elseif (-not $isChecked -and $isCurrentlyMember) {
                # Need to remove
                $toRemove += $groupName
            }
        }

        if ($toAdd.Count -eq 0 -and $toRemove.Count -eq 0) {
            Show-Modal "No Changes" "No group membership changes were made"
            return
        }

        # Show confirmation
        $changeMsg = ""
        if ($toAdd.Count -gt 0) {
            $changeMsg += "Add to $($toAdd.Count) group(s):`n  " + ($toAdd -join "`n  ") + "`n`n"
        }
        if ($toRemove.Count -gt 0) {
            $changeMsg += "Remove from $($toRemove.Count) group(s):`n  " + ($toRemove -join "`n  ")
        }

        $confirm = [Terminal.Gui.MessageBox]::ErrorQuery(
            "Confirm Changes",
            "Apply these changes for $($User.Name)?`n`n$changeMsg",
            "Yes", "No"
        )

        if ($confirm -ne 0) {  # Not "Yes"
            return
        }

        # Apply changes
        $addedCount = 0
        $removedCount = 0
        $errors = @()

        # Add to groups
        foreach ($groupName in $toAdd) {
            try {
                if ($Script:DemoMode) {
                    if (-not $User.Groups) { $User.Groups = @() }
                    $User.Groups += $groupName
                    Debug-Log ": Added $($User.Name) to group $groupName (demo)" -Type "Success"
                    $addedCount++
                } else {
                    Add-ADGroupMember -Identity $groupName -Members $User.SamAccountName -ErrorAction Stop
                    Debug-Log ": Added $($User.SamAccountName) to group $groupName" -Type "Success"
                    $addedCount++
                }
            } catch {
                $errors += "Failed to add to ${$groupName}: $($_.Exception.Message)"
                Debug-Log ": Failed to add: $($_.Exception.Message)" -Type "Error"
            }
        }

        # Remove from groups
        foreach ($groupName in $toRemove) {
            try {
                if ($Script:DemoMode) {
                    $User.Groups = $User.Groups | Where-Object { $_ -ne $groupName }
                    Debug-Log ": Removed $($User.Name) from group $groupName (demo)" -Type "Success"
                    $removedCount++
                } else {
                    Remove-ADGroupMember -Identity $groupName -Members $User.SamAccountName -Confirm:$false -ErrorAction Stop
                    Debug-Log ": Removed $($User.SamAccountName) from group $groupName" -Type "Success"
                    $removedCount++
                }
            } catch {
                $errors += "Failed to remove from ${$groupName}: $($_.Exception.Message)"
                Debug-Log ": Failed to remove: $($_.Exception.Message)" -Type "Error"
            }
        }

        # Show results
        $resultMsg = "Changes applied:"
        if ($addedCount -gt 0) { $resultMsg += "`nAdded to $addedCount group(s)" }
        if ($removedCount -gt 0) { $resultMsg += "`nRemoved from $removedCount group(s)" }

        if ($errors.Count -gt 0) {
            $resultMsg += "`n`nErrors:`n" + ($errors -join "`n")
            Show-Modal "Partial Success" $resultMsg
        } else {
            Show-Modal "Success" $resultMsg
        }

        # Call OnUpdate callback if provided
        if ($OnUpdate) {
            & $OnUpdate
        }

        # Close the dialog
        [Terminal.Gui.Application]::RequestStop()
    })
    $dlg.Add($btnApply)

    # Cancel button
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.X = [Terminal.Gui.Pos]::Right($btnApply) + 2
    $btnCancel.Y = [Terminal.Gui.Pos]::Bottom($scrollView) + 1
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
    $dlg.Add($btnCancel)

    # Run the dialog
    [Terminal.Gui.Application]::Run($dlg)
}

# =====================================================
# Usage Example
# =====================================================

<#
# Call from User Properties Member Of tab - ONE button instead of two!
$btnEdit = [Terminal.Gui.Button]::new("Edit Groups...")
$btnEdit.add_Clicked({
    Show-EditGroupMembershipDialog -User $user -OnUpdate {
        # Refresh group list
        $refreshedGroups = @()
        if ($user.Groups) {
            $refreshedGroups = $user.Groups
        } elseif ($user.MemberOf) {
            $refreshedGroups = $user.MemberOf | ForEach-Object {
                if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
            }
        }

        if ($refreshedGroups.Count -gt 0) {
            $lstGroups.SetSource($refreshedGroups)
        } else {
            $lstGroups.SetSource(@("(No group memberships)"))
        }
    }
})
#>

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

    Debug-Log (": Showing OU properties dialog for: $ouName") -Type "Info"

    if (-not $ouName) {
        Debug-Log (": OU name is null") -Type "Warn"
        return
    }

    # Find the OU in raw data
    $ou = $Script:rawOUs | Where-Object { $_.Name -eq $ouName } | Select-Object -First 1

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
    $txtName.add_TextChanged({ $Script:ouChangesMade = $true })
    $view.Add($txtName)
    $y+=2

    # Description
    $lbl = [Terminal.Gui.Label]::new("Description:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
    $txtDesc = [Terminal.Gui.TextField]::new($ou.Description ?? ""); $txtDesc.X=20; $txtDesc.Y=$y; $txtDesc.Width=50
    $txtDesc.add_TextChanged({ $Script:ouChangesMade = $true })
    $view.Add($txtDesc)
    $y+=2

    # Path (read-only)
    $lbl = [Terminal.Gui.Label]::new("Path:"); $lbl.X=2; $lbl.Y=$y; $view.Add($lbl)
    $txtPath = [Terminal.Gui.TextField]::new($ou.Path ?? ""); $txtPath.X=20; $txtPath.Y=$y; $txtPath.Width=50; $txtPath.ReadOnly=$true
    $view.Add($txtPath)
    $y+=2

    # Show object count in this OU
    $objectCount = ($Script:Users | Where-Object { $_.OU -contains $ouName }).Count
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
        if ($Script:ouChangesMade) {
            $result = [Terminal.Gui.MessageBox]::Query(60, 8, "Unsaved Changes",
                "You have unsaved changes. Discard them?",
                @("Yes", "No"))
            if ($result -eq 0) {
                $Script:ouChangesMade = $false
                [Terminal.Gui.Application]::RequestStop()
            }
        } else {
            [Terminal.Gui.Application]::RequestStop()
        }
    })

    $btnApply.add_Clicked({
        Apply-OUChanges -ou $ou -fields $fields
    })

    Debug-Log (": Show-OUPropertiesDialog running") -Type "Info"
    [Terminal.Gui.Application]::Run($dlg)
    Debug-Log (": Show-OUPropertiesDialog completed") -Type "Info"
}

function Apply-OUChanges {
    param($ou, $fields)

    $originalName = $fields.originalName
    $newName = $fields.txtName.Text.ToString()
    $newDesc = $fields.txtDesc.Text.ToString()

    Debug-Log (": Applying changes for OU: $originalName -> $newName") -Type "Info"

    try {
        if ($Script:DemoMode) {
            # Check if renaming
            $isRename = $originalName -ne $newName

            if ($isRename) {
                Debug-Log (": Renaming OU from '$originalName' to '$newName'") -Type "Info"

                # Update the OU itself
                $ou.Name = $newName
                $ou.Description = $newDesc

                # Update all users that reference this OU
                foreach ($user in $Script:Users) {
                    if ($user.OU -contains $originalName) {
                        Debug-Log (": Updating user $($user.Name) OU reference") -Type "Info"
                        $user.OU = $user.OU | ForEach-Object { if ($_ -eq $originalName) { $newName } else { $_ } }
                    }
                }

                # Update raw users too
                foreach ($rawUser in $Script:rawUsers) {
                    if ($rawUser.OU -contains $originalName) {
                        $rawUser.OU = $rawUser.OU | ForEach-Object { if ($_ -eq $originalName) { $newName } else { $_ } }
                    }
                }

                # Update in rawOUs
                $rawOU = $Script:rawOUs | Where-Object { $_.Name -eq $originalName } | Select-Object -First 1
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

                $rawOU = $Script:rawOUs | Where-Object { $_.Name -eq $originalName } | Select-Object -First 1
                if ($rawOU) {
                    $rawOU.Description = $newDesc
                }

                Show-Modal "Success" "OU changes applied (demo mode)"
            }

            Debug-Log ("SUCCESS: OU changes applied (demo mode)") -Type "Info"

            # Refresh the tree to show changes
            Refresh-Data -domain $Script:CurrentDomain

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

                    Debug-Log ("SUCCESS: OU renamed in AD from $originalName to $newName") -Type "Info"
                    Show-Modal "Success" "OU renamed successfully"
                } else {
                    throw "OU not found in AD"
                }
            } else {
                # Just update description
                $adOU = Get-ADOrganizationalUnit -Filter "Name -eq '$originalName'" -ErrorAction Stop | Select-Object -First 1
                if ($adOU) {
                    Set-ADOrganizationalUnit -Identity $adOU.DistinguishedName -Description $newDesc -ErrorAction Stop
                    Debug-Log ("SUCCESS: OU description updated in AD") -Type "Info"
                    Show-Modal "Success" "OU changes applied"
                }
            }

            # Refresh from AD
            Refresh-Data -domain $Script:CurrentDomain
        }

        $Script:ouChangesMade = $false
        return $true

    } catch {
        Debug-Log (": Failed to apply OU changes: $($_.Exception.Message)") -Type "Warn"
        Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)"
        return $false
    }
}

## ===== STEP 1: Environment & Logging =====

## Don't let users do stupid stuff, either unintentionally or willingly
if ($DemoMode -and $PSBoundParameters.ContainsKey('Domain')) {
    Debug-Log "Invalid startup: -Domain cannot be used with -DemoMode" -Type "Error"
    return
}

## Echo basic info for debugging
Debug-Log "DemoMode: $DemoMode" "info"
Debug-Log "Logging: $Logging" -Type "Info"
Debug-Log "LogFile: $LogFile" -Type "Info"

## Initialize logging if requested
if ($Logging -or $LogFile) {
    Debug-Log "Logging condition TRUE" "Debug"

    if (-not $LogFile) {
        $LogFile = Join-Path $PSScriptRoot "dsa_tui_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        Debug-Log "Auto-generated LogFile: $LogFile" -Type "Debug"
    }

    # Convert to absolute path if needed
    if (-not [System.IO.Path]::IsPathRooted($LogFile)) {
        $LogFile = Join-Path (Get-Location).Path $LogFile
    }

    $Script:Logging = $true
    $Script:LogFile = $LogFile

    # Spinner setup for status updates
    $Script:statusSpinner = @('|', '/', '-', '\')
    $Script:statusSpinnerIndex = 0

    # Dummy status object for UI
    $Script:StatusItem = [PSCustomObject]@{ Title = "Initializing..." }

    Debug-Log "Attempting to create log at: $LogFile" -Type "Debug"
    try {
        $Script:LogStream = [System.IO.StreamWriter]::new($LogFile, $false)
        $Script:LogStream.AutoFlush = $true
        Debug-Log "SUCCESS: Log file created at $LogFile" -Type "Debug"

        $Script:LogStream.WriteLine("=== DSA-TUI Log Started $(Get-Date) ===")
        $Script:LogStream.WriteLine("DemoMode: $DemoMode")
        $Script:LogStream.WriteLine("Theme: $Theme")
        $Script:LogStream.Flush()
    } catch {
        Debug-Log "FAILED to create log: $_" -Type "Error"
        $Script:Logging = $false
    }
} else {
    Debug-Log "Logging condition FALSE - no logging enabled" -Type "Warn"
}

# Set globals for demo mode and theme
$Script:DemoMode = $DemoMode
$Script:ThemeMode = $Theme

## ===== STEP 2: Module Checks & Terminal.Gui =====

Debug-Log "Performing pre-flight module checks..." -Type "Info"

# Required module: Terminal.Gui via ConsoleGuiTools
$requiredOK = Test-RequiredModule -Name "Microsoft.PowerShell.ConsoleGuiTools"
if (-not $requiredOK) {
    Debug-Log "Missing required module Microsoft.PowerShell.ConsoleGuiTools. Exiting." -Type "Error"
    exit
}

# Optional module: ActiveDirectory
$adAvailable = Test-RequiredModule -Name "ActiveDirectory" -Optional
if (-not $adAvailable) {
    Debug-Log "ActiveDirectory module missing. Falling back to DEMO mode..." -Type "Warn"
    $Script:DemoMode = $true
}

# Optional modules
$Script:HasPSWriteColor = Test-RequiredModule -Name "PSWriteColor" -Optional
$null = Test-RequiredModule -Name "Terminal-Icons" -Optional

# Ensure Terminal.Gui.dll is loaded
Debug-Log "Checking Terminal.Gui assembly..." -Type "Debug"
if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Terminal.Gui' })) {
    $mod = Get-Module -ListAvailable Microsoft.PowerShell.ConsoleGuiTools | Select-Object -First 1
    if ($mod) {
        $dll = Join-Path $mod.ModuleBase 'Terminal.Gui.dll'
        if (Test-Path $dll) {
            Add-Type -Path $dll -ErrorAction Stop
            Debug-Log "Loaded Terminal.Gui from $dll" -Type "Debug"
        } else {
            Write-Error "Terminal.Gui.dll not found in $($mod.ModuleBase). Install Microsoft.PowerShell.ConsoleGuiTools."
            return
        }
    } else {
        Write-Error "Microsoft.PowerShell.ConsoleGuiTools module not found."
        return
    }
} else {
    Debug-Log "Terminal.Gui assembly already loaded." -Type "Info"
}

## ===== STEP 3: Forest/Domain Initialization & Globals =====

Debug-Log "Initializing forest/domain globals..." -Type "Info"

# Colours and demo mode
$Script:DemoMode = $DemoMode
$Script:ThemeMode = $Theme

# Spinner setup
$Script:SpinnerTimer  = $null
$Script:SpinnerActive = $false

# Tab & layout placeholders
$Script:LayoutInProgress = $false
$Script:TabRows          = @()
$Script:AllTabs          = @()
$Script:ActiveTab        = $null
$Script:TabRowHeight     = 1

## ===== STEP 4: Load Domain/Demo Data & Build Tree =====

########################### HERE IS WHERE IT DOESNT SHOW A UI FOR AGES ##############################################

## ==================== STEP 5: Initialize Terminal.Gui UI ====================

[Terminal.Gui.Application]::Init()
$top = [Terminal.Gui.Application]::Top

## Check if date is special and use emoji accordingly
Initialize-DirectoryEmoji

## ------------------------- Main Window (ONLY ONCE) -------------------------
$win = [Terminal.Gui.Window]::new("$($Script:ProjectName) $($Script:DirectoryEmoji) Active Directory $BuildVersion $($Global:FruitName)")
$win.X = 0
$win.Y = 0
$win.Width  = [Terminal.Gui.Dim]::Fill()
$win.Height = [Terminal.Gui.Dim]::Fill(1)

## Apply theme to window early (safe)
## Select theme before proceeding. Save the mode string
$Script:ThemeMode = $Theme

## Get the selected colour scheme
$themeData = Get-Theme -mode $Theme  # Store in $themeData, not $cs
Debug-Log "Applying theme: $Theme" -Type "Info"
$themeData = Get-Theme -mode $Theme
#Apply-Theme -ThemeData $themeData -TopLevel [Terminal.Gui.Application]::Top -MainWindow $win -Menu $menu -Status $Script:StatusBar
Apply-Theme -ThemeData $themeData -TopLevel [Terminal.Gui.Application]::Top -MainWindow $win -Menu $menu -Status $Script:StatusBar
Debug-Log "Theme applied successfully" -Type "Info"

$top.Add($win)

# ------------------------- Filter Panel (right, top) -------------------------
$filterPanel = Create-FilterPanel
if (-not ($filterPanel -is [Terminal.Gui.View])) {
    $filterPanel = [Terminal.Gui.FrameView]::new("Filters")
}

$filterPanel.Width  = 40 ## <-- if you need a wider panel, change this
$filterPanel.Height = 17 ## <- chnage this if you make add new filters on the panel
$filterPanel.X = [Terminal.Gui.Pos]::AnchorEnd(40)
$filterPanel.Y = 0
$win.Add($filterPanel)

# -------------------------{} Selected Objects Panel (right, below filters) }-------------------------
$selectedObjectsPanel = Create-SelectionPanel
if (-not ($selectedObjectsPanel -is [Terminal.Gui.View])) {
    $selectedObjectsPanel = [Terminal.Gui.FrameView]::new("Selected Objects")
}

$selectedObjectsPanel.Width  = 40
$selectedObjectsPanel.Height = 10
$selectedObjectsPanel.X = [Terminal.Gui.Pos]::AnchorEnd(40)
$selectedObjectsPanel.Y = 17
$win.Add($selectedObjectsPanel)

############################# This needs to move to after the TUI loads ######################################

# Forest/domain structure
if ($Script:DemoMode) {
    Debug-Log "DemoMode enabled: creating demo forest structure..." -Type "Info"
    $Script:ForestName = "jukebox.example"
    $Script:RootDomain = "example.com"
    $Script:Domains    = @('example.com', 'example.net', 'example.org')
    $Script:Sites      = @('GLA', 'EDI', 'LND', 'CPH', 'KGE', 'ODE', 'BON', 'BRL', 'MUC', 'NEW', 'LIV')
    $Script:CurrentDomain = $Script:RootDomain
} else {
    Debug-Log "Production mode: querying AD forest..." -Type "Info"
    try {
        if ($Domain) {
            Debug-Log "Querying specified domain: $Domain" -Type "Info"
            $targetDomain = Get-ADDomain -Server $Domain -ErrorAction Stop
            $forest = Get-ADForest -Server $targetDomain.Forest -ErrorAction Stop
        } else {
            $targetDomain = Get-ADDomain -ErrorAction Stop
            $forest = Get-ADForest -ErrorAction Stop
        }

        $Script:ForestName = $forest.Name.Split('.')[0].ToUpper()
        $Script:RootDomain  = $forest.RootDomain
        $Script:Domains     = $forest.Domains
        $Script:Sites       = $forest.Sites | ForEach-Object { $_.Name }
        $Script:CurrentDomain = if ($Domain) { $Domain } else { $Script:RootDomain }
    } catch {
        Debug-Log ("Failed to query AD domain/forest: $_") -Type "Error"
        Debug-Log ("Falling back to minimal domain info.") -Type "Warn"
        $Script:ForestName = "DOMAIN"
        $Script:RootDomain  = if ($Domain) { $Domain } else { $env:USERDNSDOMAIN }
        $Script:Domains     = @($Script:RootDomain)
        $Script:Sites       = @()
        $Script:CurrentDomain = $Script:RootDomain
    }
}

$Script:Domain = $Script:CurrentDomain  ## Compatibility

# Initialize object arrays
$Script:CurrentDC      = $null
$Script:Users          = @()
$Script:Groups         = @()
$Script:DCs            = @()
$Script:ADObjects      = @()
$Script:SelectedObjects = @()
$Script:SelectionMode   = $false

# Global Search filters
$Script:FilterOptions = @{
  ShowDisabledUsers = $true
  ShowEnabledUsers  = $true
  ShowLockedUsers   = $true
  ShowGroups        = $true
  ShowDCs           = $true
  ShowComputers     = $true
  ShowOUs           = $true
  NameFilter        = ""
  SortBy            = "Name"
  SortDescending    = $false
}

# -------------------------{ Try to load data and refresh status bar after TUI }-------------------------
## Do this AFTER blanking out the arrays just above
Debug-Log "Loading domain data for $($Script:CurrentDomain)..." -Type "Info"
##Update-Status "Loading domain data for $($Script:CurrentDomain)..." -spinner
Load-DomainData -domain $Script:CurrentDomain
##Update-Status "Refresh complete" -final
Debug-Log ("POST-LOAD: Users=$($Script:Users.Count), Objects=$($Script:ADObjects.Count), DCs=$($Script:DCs.Count)") -Type "Info"

Debug-Log "Forest/Domain initialization complete: CurrentDomain=$($Script:CurrentDomain)" -Type "Info"

#####################################################################################################################

# ------------------------- TreeView -------------------------
Debug-Log ": Initializing TreeView..." -Type "Info"

# Create a FrameView to hold the tree
$treeFrame = [Terminal.Gui.FrameView]::new("Active Directory Objects")
$treeFrame.X = 0
$treeFrame.Y = 0  # Y=1 to leave room for menu
$treeFrame.Width = [Terminal.Gui.Dim]::Fill(42)  # Leave room for right panels
$treeFrame.Height = [Terminal.Gui.Dim]::Fill()

# Create and configure the tree
$Script:tree = [Terminal.Gui.TreeView]::new()
$Script:tree.X = 0
$Script:tree.Y = 0
$Script:tree.Width = [Terminal.Gui.Dim]::Fill()
$Script:tree.Height = [Terminal.Gui.Dim]::Fill()

# Build and populate tree
$rootNode = Build-Tree -domain $Script:CurrentDomain
if ($null -ne $rootNode) {
    $Script:tree.ClearObjects()
    $Script:tree.AddObject($rootNode)
    Debug-Log ": Root node added to TreeView" -Type "Success"
} else {
    Debug-Log ": FATAL - Build-Tree returned null root node" -Type "Error"
}

# Add tree to frame, then frame to window
$treeFrame.Add($Script:tree)
$win.Add($treeFrame)

Debug-Log ": TreeView created and added to window successfully" -Type "Success"
Update-Status "Initializing TreeView..." -spinner

# ------------------------- Status Bar -------------------------
$Script:StatusItem = [Terminal.Gui.StatusItem]::new(0, "Ready", $null)

$Script:StatusBar = [Terminal.Gui.StatusBar]::new(@(
  [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F1, "~F1~ Help", { Show-Modal "Shortcuts" "F1 - Help`nF2 - Password Generator`nF3 - New`nF5 - Refresh`nF6 - Themes`nF7 - Search`nF9 Menus`nF10 - Quit`nF11 Full Screen" }),
  [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F2, "~F2~ Password Generator", { Generate-RandomPassword }),
  [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F3, "~F3~ New", { Show-NewObjectWizard }),
  [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F5, "~F5~ Refresh", { Refresh-Data -domain $Script:CurrentDomain -RebuildTree }),
  [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F6, "~F6~ Themes", { Show-ThemeSelector }),
  [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F7, "~F7~ Search", { Show-ADSearchDialog }),
  [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F9, "~F9~ Menus", { }),
  [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F10, "~F10~ Quit", { [Terminal.Gui.Application]::RequestStop() }),
  [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F11, "~F11~ Full Screen", { }),
  $Script:StatusItem
))

$top.Add($Script:StatusBar)

## Final State of status bar after startup
Update-Status "Refresh complete" -final

## ------------------------- Menu -------------------------
$menu = Build-MainMenu
$top.Add($menu)

## ------------------------- Global Key Handlers -------------------------
$top.add_KeyPress({
  param($e)
  $handled = $false

  switch ($e.KeyEvent.Key) {
    ([Terminal.Gui.Key]::F1)  { Show-Modal "Shortcuts" "F1 - Help`nF2 - Password Generator`nF3 - New`nF5 - Refresh`nF6 - Themes`nF7 - Search`nF9 Menus`nF10 - Quit`nF11 Full Screen" ; $handled = $true }
    ([Terminal.Gui.Key]::F2)  { Generate-RandomPassword ; $handled = $true }
    ([Terminal.Gui.Key]::F3)  { Show-NewObjectWizard ; $handled = $true }
    ([Terminal.Gui.Key]::F5)  { Refresh-Data -domain $Script:CurrentDomain -RebuildTree ; $handled = $true }
    ([Terminal.Gui.Key]::F6)  { Show-ThemeSelector ; $handled = $true }
    ([Terminal.Gui.Key]::F7)  { Show-ADSearchDialog ; $handled = $true }
    ([Terminal.Gui.Key]::F10) { [Terminal.Gui.Application]::RequestStop() ; $handled = $true }
  }

  $e.Handled = $handled
})

## ------------------------- Debug View Tree -------------------------
Debug-Log "=== FULL VIEW TREE DUMP ===" -Type "Info"
Debug-DumpViewTree -View $top
Debug-Log "=== END VIEW TREE DUMP ===" -Type "Info"

## ------------------------- Run -------------------------
[Terminal.Gui.Application]::Run($top)

## ------------------------- Cleanup -------------------------
Stop-Spinner
[Terminal.Gui.Application]::Shutdown()
Debug-Log "Application shut down cleanly" -Type "Info"

## He fights for the users...
Debug-Log "End of line..." -Type "Info"
