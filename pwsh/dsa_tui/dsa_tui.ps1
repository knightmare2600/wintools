<#

DSA-TUI Text Mode version of dsa.msc for powershell
Locked-in baseline: dynamic resize, menu, demo data mirrors prod format, Change Domain fixed, fixed DC selection, full production AD object detection, properties modal, AD search popup

If you need to bulk add another AD property to demo data:

$Script:rawUsers | ForEach-Object {
  $sam = if ($_.SamAccountName) { $_.SamAccountName }
         elseif ($_.Email) { ($_.Email -split '@')[0] }
         else { '' }

  $upn = if ($_.UserPrincipalName) { $_.UserPrincipalName }
         elseif ($_.Email) { $_.Email }
         else { '' }

  "@{ Name = '$($_.Name)'; SamAccountName = '$sam'; UserPrincipalName = '$upn' },"
}

Then replace }, in the output with , and use column mode to your advantage

----------{ How this code works }----------
ORDER OF OPERATIONS: From AD Data to Display

1. DATA LOADING (happens once at startup or refresh)
 Load-ADData or Load-DefaultDemoData
   ↓
 Populates: $Script:Users, $Script:Groups, $Script:Computers, $Script:DCs, $Script:rawOUs

2. TREE BUILDING (happens when you switch domains or refresh)
 Build-Tree
   ↓
 Build-DomainContent (for each domain)
   ↓
   Get-OrCreateChildNode (creates each OU/container node)
   ↓
   Looks up OU data from $Script:rawOUs Sets node.Tag = @{ Type='ou'; Object=$ouData }
   ↓
   Creates user/group/computer nodes under their parent OUs. Sets node.Tag = @{ Type='user'; Object=$userObject }

3. USER CLICKS ON A NODE
 TreeView selection changes
   ↓
  Show-Properties gets called
   ↓
  Reads $node.Tag.Type to determine what was clicked
   ↓
  Calls the appropriate dialog:
    - Show-UserPropertiesDialog (if Type='user')
    - Show-GroupPropertiesDialog (if Type='group')
    - Show-OUPropertiesDialog (if Type='ou')
    - Show-ComputerPropertiesDialog (if Type='computer')
    - Show-DCPropertiesDialog (if Type='dc')
   ↓
  Dialog reads $node.Tag.Object to get the actual data and displays it in the UI

KEY CONCEPT:
  Tag acts as the bridge - it stores both the Type and the actual Object data, so when you click something,
  Show-Properties knows exactly what it is and where to find its data.

## grep and awk type stuff for pwsh
Select-String .\users.ps1 -Pattern "Company\s*=\s*'([^']+)'" | ForEach-Object { $_.Matches[0].Groups[1].Value } |  Sort-Object -Unique

Recent changelog

3.2.4.03 (It was DNS, it's always DNS...)
  - Add DNS Lookup modal function which works like whttp://hatsmydns.net
  - Fix scoping issue with FSMO function in AD health dialog
  - Fix spacing issue with environment info panel
  - Rework menus to move AD specific sub functions such as IPSec or Printers to a dedicated menu
  - Create a Utilities menu to hold password generator, DNS lookup, etc

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~{ TODO / COME BACK TO }~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

REMAINING FEATURES TO IMPLEMENT:
  - AD Health tabs like Group Policy and domain controllers the tab pane could be smaller with a search box in them to help out
  - https://learn.microsoft.com/en-us/previous-versions/windows/desktop/dacx/how-to-set-up-a-resource-property

  BUGS:

#>

param(
  [switch]$DemoMode,
  [switch]$Logging,
  [string]$LogFile,
  [string]$Domain,              ## User can specify domain
  [switch]$ImportDemoData,      ## Trigger demo data import
  [string]$DemoDataFile,        ## Path to user's chosen CSV demo data file
  [ValidateSet("british", "class91", "dark", "db-1980s", "dsb", "gemstones", "intercity-swallow", "irn-bru", "light", "matrix", "ns", "network-southeast", "panam", "procomm", "scotrail", "twa", "viarail", "viarail-soft")]
  [string]$Theme,
  [Bool]$LaunchReady            ## Will check prerequisites only then exit
)

## $PSBoundParameters contains only the parameters that were actually provided. But $args contains ALL command-line
## arguments, including invalid ones. Extract parameter names from command line (anything starting with -)
$providedParams = $args | Where-Object { $_ -match '^-' } | ForEach-Object {  $_ -replace '^-+', '' }

## Define valid parameter names (including aliases)
$validParams = @(
  'DemoMode', 'Logging', 'LogFile', 'Domain', 'ImportDemoData', 'DemoDataFile', 'Theme', 'LaunchReady',
  ## Common parameters built in to pwsh - DO NOT USE:
  'Verbose', 'Debug', 'ErrorAction', 'WarningAction', 'WhatIf', 'Confirm'
)

## Find invalid parameters
$invalidParams = $providedParams | Where-Object {
  $param = $_
  $isValid = $validParams | Where-Object { $_ -ieq $param }
  -not $isValid  # Return true if NOT valid
}

## If invalid parameters found, show error and exit. Use Write-Host as Debug-Log is not available yet
if ($invalidParams) {
  Write-Host "`n❌ ERROR: Invalid parameter(s) detected:" ($invalidParams -join ', ') -ForegroundColor Red
  Write-Host "`nValid parameters: -DemoMode, -Logging, -LogFile, -Domain -ImportDemoData, -DemoDataFile -Theme, -LaunchReady`n" -ForegroundColor Cyan
  Write-Host "Example: " -NoNewline -ForegroundColor Green
  Write-Host ".\dsa_tui.ps1 -DemoMode -Theme panam -ImportDemoData -DemoDataFile .\data.csv`n" -ForegroundColor White
  exit 1
}

## Debug when errors scroll past too fast, and Debug-Log can't capture them
#$ErrorActionPreference = 'Stop'

## Define the build version, project and code names once only - up here to ease patching. The rest in main
## execution loop, where they belong
$Script:ProjectName  = "DSA-TUI pwsh dsa.msc TUI"
$Script:FruitName    = "Blåbær"
$Script:BuildVersion = "3.2.3.10"

## Global emojis
$Script:Icons = @{
  ## Core AD object types
  User      = "👤"
  Group     = "👥"
  Computer  = "💻"
  OU        = "📁"
  DC        = "🖥"
  ## Status indicators
  Locked    = "🔒"
  Disabled  = "⊗"
  Enabled   = "●"
}

class OUNode {
  [string]$Name
  [string]$Path
  [string]$Description
  [System.Collections.Generic.List[object]]$Children
  [object]$Tag

  OUNode([string]$name) {
    $this.Name = $name
    $this.Path = ""
    $this.Description = ""
    $this.Children = [System.Collections.Generic.List[object]]::new()
  }

  ## Optional: Add a constructor that takes all properties
  OUNode([string]$name, [string]$path, [string]$description) {
    $this.Name = $name
    $this.Path = $path
    $this.Description = $description
    $this.Children = [System.Collections.Generic.List[object]]::new()
  }

  [string] ToString() { return $this.Name }
}

## Initialise FilterOptions at script startup
$Script:FilterOptions = @{
  ShowDisabledUsers       = $true
  ShowEnabledUsers        = $true
  ShowPasswordExpiring72h = $true
  ShowPasswordExpired     = $true
  ShowLockedUsers         = $true
  ShowGroups              = $true
  ShowDCs                 = $true
  ShowComputers           = $true
  ShowOUs                 = $true
  ShowUsersNoGroups       = $true
  ShowDevicesNoLAPS       = $true
  ShowDevicesNoBitlocker  = $true
  NameFilter              = ""
  NameOperator            = "Contains"
  QuickFilter             = "All"
  SortBy                  = "Name"
  SortDescending          = $false
}

## Initialise CurrentDCName at script startup
$Script:CurrentDCName   = "(None)" ## No DC set initially
$Script:ImportedRawData = $null    ## Stores data from imported files
$Script:ImportSource    = $null    ## Track where data came from ('File', 'AD', 'Demo')

## Diagnostic - Check what type it is
Write-Host "FilterOptions type: $($Script:FilterOptions.GetType().Name)" -Type "Insight"
Write-Host "FilterOptions is hashtable: $($Script:FilterOptions -is [hashtable])" -Type "Insight"

##--------------------------------------------------------------------------------------------------------------##
## Any functions added in here, make sure to keep chronology when calling them from inside other functions...   ##
##--------------------------------------------------------------------------------------------------------------##

## ----------{ Test For Required Modules }----------
function Test-Requirement {
  param(
    [Parameter(Mandatory)]
    [ValidateSet("Module","WindowsFeature","WindowsCapability","ChocoApp")]
    [string]$Type,
    [Parameter(Mandatory)]
    [string]$Name,
    [string]$MinimumVersion = $null,
    [string]$InstallMsg,
    [switch]$Optional
  )

  ## Check PowerShell version
  if ($PSVersionTable.PSVersion.Major -lt 7) {
    Debug-Log "This script requires PowerShell 7.x or higher. You are running $($PSVersionTable.PSVersion)." -Type "Problem"
    return $false
  }

  ## Detect platform
  $Platform  = $PSVersionTable.PSEdition
  $OS        = $PSVersionTable.Platform

  ## Check elevated/admin on Windows
  $IsAdmin = $false
  if ($IsWindows) { $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator") }

  ## Windows-specific helpers
  $IsWindowsServer = $false
  $HasServerManagerModule = $false
  if ($IsWindows) {
    $osInfo                 = Get-CimInstance Win32_OperatingSystem
    $IsWindowsServer        = $osInfo.ProductType -in 2,3  # 2=Domain Controller, 3=Server
    $HasServerManagerModule = Get-Module -ListAvailable -Name ServerManager -ErrorAction SilentlyContinue
  }

  ## Main switch by Type
  switch ($Type) {

    "Module" {
      $module = Get-Module -ListAvailable -Name $Name -ErrorAction SilentlyContinue
      if ($module) {
        if ($MinimumVersion -and ($module.Version -lt [version]$MinimumVersion)) {
          Debug-Log "Module '$Name' found but version is too old. Need $MinimumVersion or later." -Type "Warning"
          if ($InstallMsg) { Debug-Log "Please run: $InstallMsg" -Type "Insight" }
          return $false
        }
        Debug-Log "Module '$Name' is installed. Importing..." -Type "Success"
        try {
          Import-Module $Name -ErrorAction Stop
          ## Special handling for ConsoleGuiTools
          if ($Name -eq 'Microsoft.PowerShell.ConsoleGuiTools') {
            if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Terminal.Gui' })) {
              $dll = Join-Path $module.ModuleBase 'Terminal.Gui.dll'
              if (Test-Path $dll) {
                Add-Type -Path $dll -ErrorAction Stop
                Debug-Log "Loaded Terminal.Gui assembly from $dll" -Type "Tracing"
              }
            }
          }
        } catch {
          Debug-Log "Failed to import module '$Name': $_" -Type "Warning"
          if ($InstallMsg) { Debug-Log "Please run: $InstallMsg" -Type "Insight" }
          return $false
        }
        return $true
      } else {
        if ($Optional) {
          Debug-Log "Optional module '$Name' is NOT installed." -Type "Warning"
        } else {
          Debug-Log "Module '$Name' is NOT installed." -Type "Problem"
        }
        if ($InstallMsg) { Debug-Log "Please run: $InstallMsg" -Type "Insight" }
        return $false
      }
    }
    "WindowsFeature" {
      if (-not $IsWindows) {
        Debug-Log "Cannot check Windows Feature '$Name' — not running on Windows." -Type "Warning"
        return $false
      }
      if (-not $IsWindowsServer -or -not $HasServerManagerModule) {
        Debug-Log "Windows Feature '$Name' cannot be checked on this OS (not Server / ServerManager module missing)." -Type "Warning"
        if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Insight" }
        return $false
      }
      if (-not $IsAdmin) {
        Debug-Log "Cannot check/install Windows Feature '$Name' — requires admin privileges." -Type "Warning"
        if ($InstallMsg) { Debug-Log "Suggested action (run elevated): $InstallMsg" -Type "Insight" }
        return $false
      }
      try {
        Import-Module ServerManager -ErrorAction Stop
        $feature = Get-WindowsFeature $Name -ErrorAction SilentlyContinue
      } catch {
        Debug-Log "Failed to query Windows Feature '$Name': $_" -Type "Warning"
        if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Insight" }
        return $false
      }
      if ($feature) {
        if ($feature.Installed) {
          Debug-Log "Windows Feature '$Name' is already installed." -Type "Success"
          return $true
        } else {
          Debug-Log "Windows Feature '$Name' is NOT installed." -Type "Problem"
          if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Insight" }
          return $false
        }
      } else {
        Debug-Log "Windows Feature '$Name' not found on this system." -Type "Warning"
        if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Insight" }
        return $false
      }
    }
    "WindowsCapability" {
      if (-not $IsWindows) {
        Debug-Log "Cannot check Windows Capability '$Name' — not running on Windows." -Type "Warning"
       return $false
      }
      if (-not $IsAdmin) {
        Debug-Log "Cannot query/install Windows Capability '$Name' — requires admin privileges." -Type "Warning"
        if ($InstallMsg) { Debug-Log "Suggested action (run elevated): $InstallMsg" -Type "Insight" }
        return $false
      }
      try {
        $cap = Get-WindowsCapability -Name $Name -Online -ErrorAction SilentlyContinue
      } catch {
        Debug-Log "Failed to query Windows Capability '$Name': $_" -Type "Warning"
        if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Insight" }
        return $false
      }
      if ($cap) {
        if ($cap.State -eq 'Installed') {
          Debug-Log "Windows Capability '$Name' is already installed." -Type "Success"
          return $true
        } else {
          Debug-Log "Windows Capability '$Name' is NOT installed." -Type "Problem"
          if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Insight" }
          return $false
        }
      } else {
        Debug-Log "Windows Capability '$Name' not found on this system." -Type "Warning"
        if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Insight" }
        return $false
      }
    }
    "ChocoApp" {
      if (-not $IsWindows) {
        Debug-Log "Cannot check ChocoApp '$Name' — Chocolatey only works on Windows." -Type "Warning"
        return $false
      }
      $installed = choco list --local-only | Where-Object { $_ -match "^$Name" }
      if ($installed) {
        Debug-Log "Choco App '$Name' is installed." -Type "Success"
        return $true
      } else {
        Debug-Log "Choco App '$Name' is NOT installed." -Type "Problem"
        if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Insight" }
        return $false
      }
    }
  }
}

function Debug-DumpViewTree {
  param(
    [Terminal.Gui.View]$View,
    [int]$Indent = 0
  )

  if (-not $View) { return }
  ## Uncomment if you'd like tree view style padding
  #  $pad = (' ' * ($Indent * 2))
  $type = $View.GetType().Name
  $x = if ($View.X) { $View.X.ToString() } else { "null" }
  $y = if ($View.Y) { $View.Y.ToString() } else { "null" }
  $w = if ($View.Width) { $View.Width.ToString() } else { "null" }
  $h = if ($View.Height) { $View.Height.ToString() } else { "null" }
  $vis = $View.Visible
  $scheme = if ($View.ColorScheme) { "HasTheme" } else { "DefaultTheme" }
  Debug-Log ("$pad$type | X=$x Y=$y W=$w H=$h Visible=$vis Theme=$scheme") -Type "Tracing"
  if ($View.Subviews -and $View.Subviews.Count -gt 0) { foreach ($child in $View.Subviews) { Debug-DumpViewTree -View $child -Indent ($Indent + 1) } }
}

## ----------{ Debug Logging }----------
function Debug-Log {
  param(
    [string]$Message,
    [ValidateSet('Insight','Warning','Problem','Success','Tracing')]
    [string]$Type = 'Insight'
  )

  $timestamp = Get-Date -Format 'HH:mm:ss'
  ## Emoji + colours for each type
  switch ($Type) {
    'Insight' { $emoji = 'ℹ'  ; $color = 'DarkYellow' }    ## Insight (FYI type Info)
    'Warning' { $emoji = '⚠' ; $color = 'Yellow'     }    ## Warning
    'Problem' { $emoji = '✖' ; $color = 'Red'        }    ## Error
    'Success' { $emoji = '✔' ; $color = 'Green'      }    ## Success
    'Tracing' { $emoji = '⚙' ; $color = 'Cyan'       }    ## Tracing (Debug)
  }

  $line = "[$timestamp] $emoji  $Type : $Message"

  ## Show in console when Logging switch is enabled
  if ($Script:Logging) {
    ## Use PSWriteColor if available, otherwise fall back to Write-Host
    if ($Script:HasPSWriteColor) {
      try     { Write-Color -Encoding UTF8 -Text $line -Color $color
      } catch { Write-Host $line -ForegroundColor $color  ## Fallback
      }
    } else    { Write-Host $line -ForegroundColor $color }
  }

  ## Write to log file if enabled
  if ($Script:Logging -and $Script:LogStream) {
    try {
      $Script:LogStream.WriteLine($line)
    } catch { }
  }
}

function Build-MainMenu {
  [CmdletBinding()]
  param()
  Debug-Log "Building main menu..." -Type "Insight"
  ## ----------{ Menu Items }----------
  $mFile         = [Terminal.Gui.MenuItem]::new("_Exit","Exit application (F10)",[Action]{ [Terminal.Gui.Application]::RequestStop() })
  $mNew          = [Terminal.Gui.MenuItem]::new("N_ew Object","Create a new object (F3)",[Action]{ Show-NewObjectWizard })
  $mProps        = [Terminal.Gui.MenuItem]::new("_Properties","Edit selected properties",[Action]{ Show-Properties })
  $mDemoExport   = [Terminal.Gui.MenuItem]::new("_Export Data", "Export demo data", [Action]{ Show-ExportDataDialog })
  $mDemoImport   = [Terminal.Gui.MenuItem]::new("_Import Data", "Import demo data", [Action]{
    $Script:selectedFile = $null
    Show-FileBrowserDialog -Mode 'Open'
    if ($Script:selectedFile -and (Test-Path -LiteralPath $Script:selectedFile)) {
      $result = Import-DataFile -FilePath $Script:selectedFile
      if ($result) {
        Show-InfoPanel -UpdateOnly
        Refresh-Data -Domain $Script:CurrentDomain -RebuildTree
      }
    }
  })
  $mUndo              = [Terminal.Gui.MenuItem]::new("_Undo","Undo last action",[Action]{ Debug-Log (" Undo placeholder") -Type "Insight" })
  $mChangeDomain      = [Terminal.Gui.MenuItem]::new("Change _Domain","Change Domain",[Action]{ Show-ChangeDomainDialog })
  $mChangeDC          = [Terminal.Gui.MenuItem]::new("Select D_C","Change Domain Controller",[Action]{ Show-ChangeDCDialog })
  $mSearchAD          = [Terminal.Gui.MenuItem]::new("_Search AD","Search AD (F7)",[Action]{ $func = ${function:Show-ADSearchDialog} ;  & $func})
  $mRefresh           = [Terminal.Gui.MenuItem]::new("_Refresh","Refresh AD data (F5)",[Action]{
    Debug-Log (" Refresh menu clicked - scheduling refresh...") -Type "Insight"
    [Terminal.Gui.Application]::MainLoop.AddTimeout([TimeSpan]::FromMilliseconds(100), {
    Debug-Log (" Timeout callback - starting refresh...") -Type "Insight"
    try {
      Set-StatusBar "Refreshing..." -Icon 'Working'
      $result = Refresh-Data -domain $Script:CurrentDomain -RebuildTree
      if ($result) { Set-StatusBar "Refresh complete" -Icon 'Success' } else { Set-StatusBar "Refresh failed" -Icon 'Error' }
        Debug-Log (" Refresh completed with result: $result") -Type "Insight"
      } catch {
        Debug-Log (" Refresh crashed: $($_.Exception.Message)") -Type "Insight"
        Set-StatusBar "Refresh error" -Icon 'Error'
      }
      return $false
    })
    Debug-Log (" Refresh scheduled") -Type "Insight"
  })

  $mQuickFilter       = [Terminal.Gui.MenuItem]::new("_Quick Filter","Apply quick filters",[Action]{Show-QuickFilterDialog})
  $mSelectionMode     = [Terminal.Gui.MenuItem]::new("_Selection Mode (Ctrl+S)","Toggle selection mode",[Action]{Toggle-SelectionMode})
  $mSelectAll         = [Terminal.Gui.MenuItem]::new("Select _All (Ctrl+A)","Select all objects",[Action]{Manage-Selection -Action 'SelectAll'})
  $mDeselectAll       = [Terminal.Gui.MenuItem]::new("_Unselect All (Ctrl+U)","Deselect all objects",[Action]{Manage-Selection -Action 'DeselectAll'})
  $mBulkDisable       = [Terminal.Gui.MenuItem]::new("_Disable Selected", "Disable selected accounts", [Action]{$objects = Get-SelectedObjectsAsObjects ; Manage-AccountStatus -Objects $objects -Action 'Disable'})
  $mBulkEnable        = [Terminal.Gui.MenuItem]::new("_Enable Selected", "Enable selected accounts", [Action]{$objects = Get-SelectedObjectsAsObjects ; Manage-AccountStatus -Objects $objects -Action 'Enable'})
  $mBulkAddGroup      = [Terminal.Gui.MenuItem]::new("Add to _Group...","Add selected users to group",[Action]{Invoke-BulkAddToGroup})
  $mBulkEdit          = [Terminal.Gui.MenuItem]::new("_Edit Attribute (Bulk)", "Change attribute of selected objects", [Action]{$objects = Get-SelectedObjectsAsObjects ; Set-BulkAttribute -Objects $objects -ShowDialog})
  $mStaleAll          = [Terminal.Gui.MenuItem]::new("Show _Stale Accounts", "Find all inactive accounts", [Action]{Show-ADSearchDialog})
  $mPasswordGenerator = [Terminal.Gui.MenuItem]::new("Pass_word Generator","Password Generator (F2)",[Action]{Generate-RandomPassword})
  $mLAPSPasswords     = [Terminal.Gui.MenuItem]::new("_LAPS Passwords","Lookup LAPS Creds (Fx)",[Action]{Show-LAPSSearchModal})
  $mADHealth          = [Terminal.Gui.MenuItem]::new("AD H_ealth Status", "AD Health/Replication Status", [Action]{Show-ADHealthDialog })
  $mShortcuts         = [Terminal.Gui.MenuItem]::new("_Shortcuts","Keyboard shortcuts (F1)",[Action]{Show-Modal "Shortcuts" "F1  - Help              `nF2  - Password Generator`nF3  - New AD Object     `nF5  - Refresh AD Data   `nF6  - Theme Selection   `nF7  - Search AD         `nF8  - Focus Tree/Pane   `nF9  - Show Menubar      `nF10 - Quit DSA-TUI      `nF11 - Enter Full Screen `nF12 - Tree Context Menu `n" })
  $mAboutDSATUI       = [Terminal.Gui.MenuItem]::new("Abou_t","About Project",[Action]{Show-Modal "About" "$($Script:ProjectName)`n`nCodename: $($Script:FruitName)`nv$($Script:BuildVersion) STABLE`nGPL-3 Copyleft`nBy Knightmare2600 (https://github.com/knightmare2600)" })
  $mWhyBlaabaer       = [Terminal.Gui.MenuItem]::new("Why _Blåbær?","Why the $($Script:FruitName) codename?",[Action]{Show-BlaabaerInfo })
  $mTheme             = [Terminal.Gui.MenuItem]::new("_Theme","Change theme (F6)",[Action]{Show-ThemeSelector })
  $menuItemGPOs       = [Terminal.Gui.MenuItem]::new("Group Policy Ob_jects", "Show Group Policies", {Show-GPOListDialog -Domain $Script:CurrentDomain})
  $mDNSLookup         = [Terminal.Gui.MenuItem]::new("DNS _Lookup","DNS query tool",[Action]{Show-DNSLookupDialog})
  $menuItemTrusts     = [Terminal.Gui.MenuItem]::new("Trust _Relationships", "Show Trust Relationships", {Show-TrustsDialog -Domain $Script:CurrentDomain})
  $menuItemIPSec      = [Terminal.Gui.MenuItem]::new("IP_Sec Policies", "Show IPSec Policies", {Show-IPSecPoliciesDialog -Domain $Script:CurrentDomain})
  $menuItemIPSecHelp  = [Terminal.Gui.MenuItem]::new("_IPSec Help", "IPSec Policies Help", {Show-IPSecHelpDialog})
  $menuItemPrinters   = [Terminal.Gui.MenuItem]::new("Print _Queues", "Show Printer Queues", {Show-PrintQueuesDialog -Domain $Script:CurrentDomain})
  $menuItemDNSDialog  = [Terminal.Gui.MenuItem]::new("_DNS Manager", "AD DNS Manager", { Show-DNSDialog })
  ## Submenu: Actions > AD Properties
  $mCopyTemplate      = [Terminal.Gui.MenuItem]::new("Copy as _Template", "Clone from selected object", [Action]{
    ## Get the currently selected node from tree
    if (-not $Script:tree -or -not $Script:tree.SelectedObject) { Show-Modal "No Selection" "Please select a user or group to use as a template" ; return }
    $selectedNode = $Script:tree.SelectedObject
    $sourceObject = $selectedNode.Tag
    if (-not $sourceObject) { Show-Modal "Invalid Selection" "Cannot copy this object type" ; return }
    ## Validate it's a user or group
    $isUser  = $sourceObject.PSObject.Properties.Match('SamAccountName') -and -not $sourceObject.PSObject.Properties.Match('ComputerType')
    $isGroup = $sourceObject.PSObject.Properties.Match('Members')
    if (-not $isUser -and -not $isGroup) { Show-Modal "Unsupported Type" "Can only copy Users or Groups as templates" ; return }
    Copy-ADObject -SourceObject $sourceObject -ShowDialog
  })

  ## ----------{ Menu Bar }----------
  ## Be mindful of nested menus here! Define sub-menus FIRST
  $mExtraFunctions = [Terminal.Gui.MenuBarItem]::new("_Utilities",@($mPasswordGenerator, $mDNSLookup, $menuItemPrinters, $menuItemIPSec))
  $mADFunctions    = [Terminal.Gui.MenuBarItem]::new("_AD Properties",@($mLAPSPasswords, $mCopyTemplate, $menuItemGPOs, $menuItemTrusts, $menuItemDNSDialog, $mADHealth, $mStaleAll))
  ## THEN create the menu bar using them
  $menu = [Terminal.Gui.MenuBar]::new(@(
    [Terminal.Gui.MenuBarItem]::new("_File", @($mRefresh, $mDemoExport, $mDemoImport, $mTheme, $mFile)),
    [Terminal.Gui.MenuBarItem]::new("_Action", @($mADFunctions, $mNew, $mProps, $mQuickFilter, $mUndo, $mChangeDomain, $mChangeDC, $mSearchAD, $mExtraFunctions)),
    [Terminal.Gui.MenuBarItem]::new("_Selection", @($mSelectionMode, $mSelectAll, $mDeselectAll, $mBulkEdit, $mBulkAddGroup, $mBulkEnable, $mBulkDisable)),
    [Terminal.Gui.MenuBarItem]::new("_Help", @($mShortcuts, $menuItemIPSecHelp, $mAboutDSATUI, $mWhyBlaabaer))
  ))
  Debug-Log "Main menu created successfully" -Type "Insight"
  return $menu
}
function script:Set-ObjectCheckboxes {
  <#
  .SYNOPSIS
  Creates and updates object-type-specific checkboxes in one unified function

  .PARAMETER Mode
  'Create' = Add checkboxes to view
  'Update' = Update checkbox states from data
  #>
  param(
    $View,
    $State,
    $Data,
    [string]$ObjectType,
    [ValidateSet('Create', 'Update')]
    [string]$Mode = 'Create'
  )

  switch ($ObjectType) {
    'User' {
      if ($Mode -eq 'Create') {
        $State.chkSearchLocked   = [Terminal.Gui.CheckBox]::new("Account Locked")
        $State.chkSearchLocked.X = 2
        $State.chkSearchLocked.Y = [Terminal.Gui.Pos]::Bottom($State.txtSearchOutput) + 1
        $State.chkSearchLocked.CanFocus = $true
        $View.Add($State.chkSearchLocked)
        $State.chkSearchDisabled   = [Terminal.Gui.CheckBox]::new("Account Disabled")
        $State.chkSearchDisabled.X = 2
        $State.chkSearchDisabled.Y = [Terminal.Gui.Pos]::Bottom($State.chkSearchLocked) + 1
        $State.chkSearchDisabled.CanFocus = $true
        $View.Add($State.chkSearchDisabled)
      } else {
        if ($State.chkSearchLocked) { $State.chkSearchLocked.Checked = [bool]($Data.Locked ?? $Data.LockedOut) }
        if ($State.chkSearchDisabled) { $State.chkSearchDisabled.Checked = [bool]($Data.Disabled ?? -not $Data.Enabled) }
      }
    }
    'Group' {
      if ($Mode -eq 'Create') {
        $State.chkSecurity = [Terminal.Gui.CheckBox]::new("Security Group")
        $State.chkSecurity.X = 2
        $State.chkSecurity.Y = [Terminal.Gui.Pos]::Bottom($State.txtSearchOutput) + 1
        $State.chkSecurity.CanFocus = $true
        $View.Add($State.chkSecurity)
        $State.chkDistribution = [Terminal.Gui.CheckBox]::new("Distribution Group")
        $State.chkDistribution.X = 22
        $State.chkDistribution.Y = [Terminal.Gui.Pos]::Bottom($State.txtSearchOutput) + 1
        $State.chkDistribution.CanFocus = $true
        $View.Add($State.chkDistribution)
        $State.chkGlobal = [Terminal.Gui.CheckBox]::new("Global")
        $State.chkGlobal.X = 48
        $State.chkGlobal.Y = [Terminal.Gui.Pos]::Bottom($State.txtSearchOutput) + 1
        $State.chkGlobal.CanFocus = $true
        $View.Add($State.chkGlobal)
        $State.chkDomainLocal = [Terminal.Gui.CheckBox]::new("Domain Local")
        $State.chkDomainLocal.X = 60
        $State.chkDomainLocal.Y = [Terminal.Gui.Pos]::Bottom($State.txtSearchOutput) + 1
        $State.chkDomainLocal.CanFocus = $true
        $View.Add($State.chkDomainLocal)
        $State.chkUniversal = [Terminal.Gui.CheckBox]::new("Universal")
        $State.chkUniversal.X = 78
        $State.chkUniversal.Y = [Terminal.Gui.Pos]::Bottom($State.txtSearchOutput) + 1
        $State.chkUniversal.CanFocus = $true
        $View.Add($State.chkUniversal)
      } else {
        if ($State.chkSecurity) {
          $groupCategory = $Data.GroupCategory ?? $Data.Type
          $State.chkSecurity.Checked = ($groupCategory -eq 'Security')
          $State.chkDistribution.Checked = ($groupCategory -eq 'Distribution')
        }
        if ($State.chkGlobal) {
          $groupScope = $Data.GroupScope ?? $Data.Scope ?? 'Global'
          $State.chkGlobal.Checked = ($groupScope -eq 'Global')
          $State.chkDomainLocal.Checked = ($groupScope -eq 'DomainLocal')
          $State.chkUniversal.Checked = ($groupScope -eq 'Universal')
        }
      }
    }
    'Computer' {
      if ($Mode -eq 'Create') {
        $State.chkSearchEnabled = [Terminal.Gui.CheckBox]::new("Enabled")
        $State.chkSearchEnabled.X = 2
        $State.chkSearchEnabled.Y = [Terminal.Gui.Pos]::Bottom($State.txtSearchOutput) + 1
        $State.chkSearchEnabled.CanFocus = $true
        $View.Add($State.chkSearchEnabled)
        $State.chkSearchDisabled = [Terminal.Gui.CheckBox]::new("Disabled")
        $State.chkSearchDisabled.X = 2
        $State.chkSearchDisabled.Y = [Terminal.Gui.Pos]::Bottom($State.chkSearchEnabled) + 1
        $State.chkSearchDisabled.CanFocus = $true
        $View.Add($State.chkSearchDisabled)
      } else {
        if ($State.chkSearchEnabled) {
          $State.chkSearchEnabled.Checked = [bool]$Data.Enabled
          $State.chkSearchDisabled.Checked = -not [bool]$Data.Enabled
        }
      }
    }
    'DomainController' {
      if ($Mode -eq 'Create') {
        $State.chkSearchGC = [Terminal.Gui.CheckBox]::new("Global Catalog")
        $State.chkSearchGC.X = 2
        $State.chkSearchGC.Y = [Terminal.Gui.Pos]::Bottom($State.txtSearchOutput) + 1
        $State.chkSearchGC.CanFocus = $true
        $View.Add($State.chkSearchGC)
        $State.chkSearchRODC = [Terminal.Gui.CheckBox]::new("Read-Only DC")
        $State.chkSearchRODC.X = 2
        $State.chkSearchRODC.Y = [Terminal.Gui.Pos]::Bottom($State.chkSearchGC) + 1
        $State.chkSearchRODC.CanFocus = $true
        $View.Add($State.chkSearchRODC)
      } else {
        if ($State.chkSearchGC) { $State.chkSearchGC.Checked = [bool]$Data.IsGlobalCatalog }
        if ($State.chkSearchRODC) { $State.chkSearchRODC.Checked = [bool]$Data.IsReadOnly }
      }
    }
  }
}

function Show-Properties {
  Debug-Log "Show-Properties called" -Type "Insight"
  $node = $Script:tree.SelectedObject
  if (-not $node) {
    Debug-Log "No object selected" -Type "Insight"
    Show-Modal "Debug" "No object selected in tree"
    return
  }
  Debug-Log "Selected node text: '$($node.Text)'" -Type "Insight"
  $tag = $node.Tag
  Debug-Log "Tag is null: $($null -eq $tag)" -Type "Insight"
  if ($tag) {
    Debug-Log "Tag.Type: '$($tag.Type)'" -Type "Insight"
    Debug-Log "Tag.Object is null: $($null -eq $tag.Object)" -Type "Insight"
    if ($tag.Object) { Debug-Log "Tag.Object type: $($tag.Object.GetType().Name)" -Type "Insight" }
  }

  ## Get the actual AD object from the Tag
  $obj = $tag.Object

  ##  Handle containers without objects
  if (-not $obj) {
    if ($tag.Type -eq 'container') {
      Debug-Log "Container selected (no properties to show)" -Type "Insight"
      Show-Modal "Container" "This is a container node.`n`nSelect an individual object to view its properties."
      return
    }
    Debug-Log "No object attached to this node" -Type "Warning"
    return
  }

  ##  Use the Type property first (more reliable)
  switch ($tag.Type) {
    'user' {
      Debug-Log "USER object selected: $($obj.Name)" -Type "Insight"
      Show-UserPropertiesDialog -user $obj
      return
    }
    'group' {
      Debug-Log "GROUP object selected: $($obj.Name)" -Type "Insight"
      Show-GroupPropertiesDialog -group $obj
      return
    }
    'dc' {
      Debug-Log "DC object selected: $($obj.Name)" -Type "Insight"
      Show-DCPropertiesDialog -dc $obj
      return
    }
    'dc-container' {
      Debug-Log "DC container selected (no properties to show)" -Type "Insight"
      Show-Modal "Domain Controllers" "This is a container for Domain Controllers in this domain`n`nSelect an individual DC to view its properties."
      return
    }
    'computer' {
      Debug-Log "COMPUTER object selected: $($obj.Name)" -Type "Insight"
      Show-ComputerPropertiesDialog -computerName $obj.Name
      return
    }
    'ou' {
      Debug-Log "Showing OU properties for $($obj.Name)" -Type "Insight"
      Show-OUPropertiesDialog -ou $obj
      return
    }
    'container' {
      Debug-Log "Container selected (no properties to show)" -Type "Insight"
      Show-Modal "Container" "This is a container node.`n`nSelect an individual object to view its properties."
      return
    }
  }

  ##  Fallback: Try to detect type from properties (for backward compatibility)
  if ($obj.PSObject.Properties.Match('SamAccountName').Count -gt 0) {
    Debug-Log "Detected USER object (fallback): $($obj.Name)" -Type "Insight"
    Show-UserPropertiesDialog -user $obj
    return
  }
  if ($obj.PSObject.Properties.Match('GroupScope').Count -gt 0 -or
    $obj.PSObject.Properties.Match('Members').Count -gt 0) {
    Debug-Log "Detected GROUP object (fallback): $($obj.Name)" -Type "Insight"
    Show-GroupPropertiesDialog -group $obj
    return
  }
  if ($obj.PSObject.Properties.Match('Site').Count -gt 0) {
    Debug-Log "Detected DC object (fallback): $($obj.Name)" -Type "Insight"
    Show-DCPropertiesDialog -dc $obj
    return
  }
  if ($obj.PSObject.Properties.Match('OperatingSystem').Count -gt 0) {
    Debug-Log "Detected COMPUTER object (fallback): $($obj.Name)" -Type "Insight"
    Show-ComputerPropertiesDialog -computer $obj
    return
  }
  Debug-Log "Selected object type unknown, cannot show properties" -Type "Warning"
}

function Show-UserPropertiesDialog {
  param($user)

  ## Capture functions for this builder's closure
  $showAuditLogFunc = ${function:Show-AuditLogDialog}
  if (-not $user) {
    Debug-Log "User object is null" -Type "Warning"
    return
  }
  Debug-Log "Show-UserPropertiesDialog starting for: $($user.Name)" -Type "Insight"

  ## ----------{ General Tab }---------
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $user, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "User Information"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Name:" -FieldName 'txtName' -State $state -Value $user.Name -IsTextField $false
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Display Name:" -FieldName 'txtDisplayName' -State $state -Value ($user.DisplayName ?? "")
      $emailAddr = if ($user.EmailAddress) { $user.EmailAddress } elseif ($user.mail) { $user.mail } else { "" }
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Email:" -FieldName 'txtEmail' -State $state -Value $emailAddr
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Description:" -FieldName 'txtDescription' -State $state -Value ($user.Description ?? "")
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Contact Information"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Office:" -FieldName 'txtOfficePhone' -State $state -Value ($user.OfficePhone ?? "") -Width 30
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Mobile:" -FieldName 'txtMobilePhone' -State $state -Value ($user.MobilePhone ?? "") -Width 30
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Account Expiration"
      $hasExpiration  = $false
      $expirationDate = $null

      try {
        if ($user.PSObject.Properties.Match('AccountExpirationDate') -and $user.AccountExpirationDate) {
          $hasExpiration = $true
          $expirationDate = $user.AccountExpirationDate
        }
      } catch {}

      $state.chkNeverExpires = [Terminal.Gui.CheckBox]::new("Never Expires")
      $state.chkNeverExpires.X = 2; $state.chkNeverExpires.Y = $y
      $state.chkNeverExpires.Checked = -not $hasExpiration
      $view.Add($state.chkNeverExpires)
      $y += 1

      $lbl = [Terminal.Gui.Label]::new("Expires On:")
      $lbl.X = 2; $lbl.Y = $y
      $view.Add($lbl)

      $expirationDateStr = if ($expirationDate) { $expirationDate.ToString("yyyy-MM-dd") } else { "" }
      $state.txtExpirationDate = [Terminal.Gui.TextField]::new($expirationDateStr)
      $state.txtExpirationDate.X = 20; $state.txtExpirationDate.Y = $y; $state.txtExpirationDate.Width = 20
      $state.txtExpirationDate.ReadOnly = $state.chkNeverExpires.Checked
      $view.Add($state.txtExpirationDate)

      $lblFmt = [Terminal.Gui.Label]::new("(yyyy-MM-dd)")
      $lblFmt.X = 43; $lblFmt.Y = $y
      $view.Add($lblFmt)
      $y += 1

      $lbl = [Terminal.Gui.Label]::new("Days Until Expiry:")
      $lbl.X = 2; $lbl.Y = $y
      $view.Add($lbl)

      $state.lblDaysUntilExpiry = [Terminal.Gui.Label]::new("N/A")
      $state.lblDaysUntilExpiry.X = 22; $state.lblDaysUntilExpiry.Y = $y
      $view.Add($state.lblDaysUntilExpiry)
      $y += 1

      ## Validation message label (hidden by default)
      $state.lblValidationError = [Terminal.Gui.Label]::new("")
      $state.lblValidationError.X = 20
      $state.lblValidationError.Y = $y
      $state.lblValidationError.ColorScheme = [Terminal.Gui.Colors]::Error
      $view.Add($state.lblValidationError)

      ## Event handlers
      $null = $state.chkNeverExpires.add_Toggled({
      $state.txtExpirationDate.ReadOnly = $state.chkNeverExpires.Checked
      if ($state.chkNeverExpires.Checked) {
        $state.txtExpirationDate.Text   = [NStack.ustring]::Make("")
        $state.lblDaysUntilExpiry.Text  = [NStack.ustring]::Make("N/A")
        $state.lblValidationError.Text  = [NStack.ustring]::Make("")
      }
    }.GetNewClosure())

  $null = $state.txtExpirationDate.add_TextChanged({
  $text = $state.txtExpirationDate.Text.ToString()
  ## Clear validation error
  $state.lblValidationError.Text = [NStack.ustring]::Make("")

  if ([string]::IsNullOrWhiteSpace($text)) {
    $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("N/A")
    return
  }

  ## Validate YYYY-MM-DD format first
  if ($text -notmatch '^\d{4}-\d{2}-\d{2}$') {
    $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("Invalid format")
    $state.lblValidationError.Text = [NStack.ustring]::Make("Use YYYY-MM-DD format")
    return
  }

  ## Extract components and validate ranges
  $parts = $text -split '-'
  $year  = [int]$parts[0]
  $month = [int]$parts[1]
  $day   = [int]$parts[2]

  ## Validate month (1-12)
  if ($month -lt 1 -or $month -gt 12) {
    $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("Invalid month")
    $state.lblValidationError.Text = [NStack.ustring]::Make("Month must be 01-12")
    return
  }

  ## Validate day (1-31, but also check for month)
  if ($day -lt 1 -or $day -gt 31) {
    $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("Invalid day")
    $state.lblValidationError.Text = [NStack.ustring]::Make("Day must be 01-31")
    return
  }

    ## Try to parse as actual date (this catches Feb 30, etc.)
    try {
      $dt = [DateTime]::new($year, $month, $day)
      ## Calculate days until expiry
      $days = ($dt - (Get-Date)).Days
      if ($days -lt 0) {
        $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("EXPIRED ($([Math]::Abs($days)) days ago)")
        $state.lblValidationError.Text = [NStack.ustring]::Make("⚠ Date is in the past")
      } elseif ($days -eq 0) {
        $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("TODAY")
      } else {
        $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("$days days")
      }
    } catch {
      $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("Invalid date")
      $state.lblValidationError.Text = [NStack.ustring]::Make("Not a valid calendar date")
    }
  }.GetNewClosure())

  $null = $state.txtExpirationDate.add_TextChanged({
    $text = $state.txtExpirationDate.Text.ToString()
    if ([string]::IsNullOrWhiteSpace($text)) {
      $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("N/A")
      return
    }
    try {
      $dt = [DateTime]::Parse($text)
      $days = ($dt - (Get-Date)).Days
      if ($days -lt 0)       { $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("EXPIRED ($([Math]::Abs($days)) days ago)")
      } elseif ($days -eq 0) { $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("TODAY")
      } else                 { $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("$days days")          }
    } catch {
      $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("Invalid date")
    }
  }.GetNewClosure())
  }
}

## ----------{ Account Tab }---------
$accountTab = @{
  Name = "Account"
  Builder = {
    param($view, $user, $state)
    $y = 1

    Add-SectionHeader -View $view -Y ([ref]$y) -Text "Logon Information"
    Add-LabelAndField -View $view -Y ([ref]$y) -Label "User logon name (UPN):" -FieldName 'txtUserPrincipalName' -State $state -Value ($user.UserPrincipalName ?? "") -FieldX 35 -Width 50
    Add-LabelAndField -View $view -Y ([ref]$y) -Label "User logon name (Pre Win 2000):" -FieldName 'txtSamAccountName' -State $state -Value ($user.SamAccountName ?? "") -FieldX 35 -Width 50
    Add-SectionHeader -View $view -Y ([ref]$y) -Text "Account Status"

    $isEnabled = if ($user.PSObject.Properties['Enabled']) { $user.Enabled } else { -not $user.Disabled }
    $state.chkEnabled = [Terminal.Gui.CheckBox]::new("Account Enabled")
    $state.chkEnabled.X=4; $state.chkEnabled.Y=$y; $state.chkEnabled.Checked=$isEnabled
    $view.Add($state.chkEnabled); $y+=1

    $isLocked = if ($user.PSObject.Properties['LockedOut']) { $user.LockedOut } else { $user.Locked ?? $false }
    $state.chkLocked = [Terminal.Gui.CheckBox]::new("Account Locked")
    $state.chkLocked.X=4; $state.chkLocked.Y=$y; $state.chkLocked.Checked=$isLocked; $state.chkLocked.Enabled=$false
    $view.Add($state.chkLocked); $y+=2

    Add-SectionHeader -View $view -Y ([ref]$y) -Text "Password Settings" -SpaceBefore 0

    $state.chkPasswordExpired = [Terminal.Gui.CheckBox]::new("Password Expired")
    $state.chkPasswordExpired.X=4; $state.chkPasswordExpired.Y=$y; $state.chkPasswordExpired.Checked=($user.PasswordExpired??$false); $state.chkPasswordExpired.Enabled=$false
    $view.Add($state.chkPasswordExpired); $y+=1

    $state.chkMustChangePassword = [Terminal.Gui.CheckBox]::new("User must change password at next logon")
    $state.chkMustChangePassword.X=4; $state.chkMustChangePassword.Y=$y
    $state.chkMustChangePassword.Checked = if ($user.PasswordNeverExpires){$false}else{ if ($user.PSObject.Properties['pwdLastSet']){ $user.pwdLastSet -eq 0 } else { $false } }
    $view.Add($state.chkMustChangePassword); $y+=1

    $state.chkCannotChangePassword = [Terminal.Gui.CheckBox]::new("User cannot change password")
    $state.chkCannotChangePassword.X=4; $state.chkCannotChangePassword.Y=$y; $state.chkCannotChangePassword.Checked=($user.CannotChangePassword??$false)
    $view.Add($state.chkCannotChangePassword); $y+=1

    $state.chkPasswordNeverExpires = [Terminal.Gui.CheckBox]::new("Password never expires")
    $state.chkPasswordNeverExpires.X=4; $state.chkPasswordNeverExpires.Y=$y; $state.chkPasswordNeverExpires.Checked=($user.PasswordNeverExpires??$false)
    $view.Add($state.chkPasswordNeverExpires); $y+=2
    Add-SectionHeader -View $view -Y ([ref]$y) -Text "Logon History" -SpaceBefore 0

    $lbl = [Terminal.Gui.Label]::new("Last logon: "+($user.LastLogonDate?.ToString('yyyy-MM-dd HH:mm') ?? 'Never'))
    $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1

    if ($user.PSObject.Properties['PasswordLastSet'] -and $user.PasswordLastSet) {
      $lbl = [Terminal.Gui.Label]::new("Password last set: "+$user.PasswordLastSet.ToString('yyyy-MM-dd HH:mm'))
      $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
    }

    if ($user.PSObject.Properties['LogonCount'] -or $user.PSObject.Properties['logonCount']) {
      $logonCount = if ($user.LogonCount) { $user.LogonCount } else { $user.logonCount }
      $lbl = [Terminal.Gui.Label]::new("Logon count: $logonCount")
      $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
    }

    ## Audit Log Button - ALWAYS show, not conditional
    Add-SectionHeader -View $view -Y ([ref]$y) -Text "Audit & Security" -SpaceBefore 1
    $btnAuditLog   = [Terminal.Gui.Button]::new("View Audit Log...")
    $btnAuditLog.X = 4
    $btnAuditLog.Y = $y
    $btnAuditLog.add_Clicked({ Show-AuditLogDialog -Object $user -ObjectType 'User' })  ## NO .GetNewClosure() !
    $view.Add($btnAuditLog)
    $y += 1
    }
  }

  ## ----------{ Address Tab }---------
  $addressTab = @{
    Name = "Address"
    Builder = {
      param($view, $user, $state)
      $y = 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Street:" -FieldName 'txtStreet' -State $state -Value ($user.StreetAddress ?? "") -Width 70
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "City:" -FieldName 'txtCity' -State $state -Value ($user.City ?? "") -Width 70
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Postal Code:" -FieldName 'txtPostal' -State $state -Value ($user.PostalCode ?? "") -Width 20
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Country:" -FieldName 'txtCountry' -State $state -Value ($user.Country ?? "") -Width 70
    }
  }

  ## ----------{ Profile Tab }---------
  $profileTab = @{
    Name = "Profile"
    Builder = {
      param($view, $user, $state)
      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "User Profile"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Profile path:" -FieldName 'txtProfilePath' -State $state -Value ($user.ProfilePath ?? "") -Width 70
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Logon script:" -FieldName 'txtLogonScript' -State $state -Value ($user.ScriptPath ?? "") -Width 70
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Home Folder"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Home directory:" -FieldName 'txtHomeDirectory' -State $state -Value ($user.HomeDirectory ?? "") -Width 70
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Home drive:" -FieldName 'txtHomeDrive' -State $state -Value ($user.HomeDrive ?? "") -Width 5
    }
  }

  ## ----------{ Organization Tab }---------
  $organizationTab = @{
    Name = "Organization"
    Builder = {
      param($view, $user, $state)
      $y = 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Title:" -FieldName 'txtTitle' -State $state -Value ($user.Title ?? "") -Width 70
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Department:" -FieldName 'txtDept' -State $state -Value ($user.Department ?? "") -Width 70
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Company:" -FieldName 'txtCompany' -State $state -Value ($user.Company ?? "") -Width 70
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Manager:" -FieldName 'txtManager' -State $state -Value ($user.Manager ?? "") -Width 70
    }
  }

  ## ----------{ Member Of Tab }---------
  $memberOfTab = @{
    Name = "Member Of"
    Builder = {
      param($view, $user, $state)
      $y = 1
      $lbl = [Terminal.Gui.Label]::new("Group Memberships:")
      $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $state.lstGroups = [Terminal.Gui.ListView]::new()
      $state.lstGroups.X = 2; $state.lstGroups.Y = $y
      $state.lstGroups.Width = [Terminal.Gui.Dim]::Fill(2)
      $state.lstGroups.Height = 25
      ## Force array: Use @() to ensure it's always an array
      $state.groupList = @()
      if ($user.Groups) { $state.groupList = @($user.Groups)
      } elseif ($user.MemberOf) { $state.groupList = @($user.MemberOf | ForEach-Object { if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ } })
      }
      if ($state.groupList.Count -gt 0) { $state.lstGroups.SetSource($state.groupList)
      } else { $state.lstGroups.SetSource(@("(No group memberships)"))
      }
      $view.Add($state.lstGroups)

      $btnAdd = [Terminal.Gui.Button]::new("Add to Group...")
      $btnAdd.X = 2; $btnAdd.Y = 28
      $btnAdd.add_Clicked({
        Show-EditGroupMembershipDialog -User $user -OnUpdate {
          $refreshedGroups = @()
          if ($user.Groups) { $refreshedGroups = @($user.Groups)
          } elseif ($user.MemberOf) { $refreshedGroups = @($user.MemberOf | ForEach-Object { if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ } })
          }
          if ($refreshedGroups.Count -gt 0) { $state.lstGroups.SetSource($refreshedGroups)
          } else { $state.lstGroups.SetSource(@("(No group memberships)"))
          }
          $state.groupList = $refreshedGroups
        }
      }.GetNewClosure())
      $view.Add($btnAdd)

      $btnRemove   = [Terminal.Gui.Button]::new("Remove from Group")
      $btnRemove.X = 22; $btnRemove.Y = 28
      $btnRemove.add_Clicked({
        $selectedIndex = $state.lstGroups.SelectedItem
        ## Force array
        $currentGroups = @()
        if ($user.Groups) { $currentGroups = @($user.Groups)
        } elseif ($user.MemberOf) { $currentGroups = @($user.MemberOf | ForEach-Object {
            if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
          })
        }
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $currentGroups.Count) {
          $selectedGroup = $currentGroups[$selectedIndex]
          if ($selectedGroup -eq "(No group memberships)") {
            Show-Modal "Info" "No group selected"
            return
          }
          $confirmDlg = Show-Modal "Confirm Removal" "Remove $($user.Name) from group '$selectedGroup'?" -YesNo
          if ($confirmDlg -eq 0) {
            try {
              if ($Script:DemoMode) { $user.Groups = @($user.Groups | Where-Object { $_ -ne $selectedGroup })
              } else { Remove-ADGroupMember -Identity $selectedGroup -Members $user.SamAccountName -Confirm:$false
              }
              ## Force array
              $updatedGroups = @()
              if ($user.Groups) { $updatedGroups = @($user.Groups)
              } elseif ($user.MemberOf) { $updatedGroups = @($user.MemberOf | Where-Object { $_ -ne $selectedGroup } | ForEach-Object {
                  if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
                })
              }
              if ($updatedGroups.Count -gt 0) { $state.lstGroups.SetSource($updatedGroups)
              } else { $state.lstGroups.SetSource(@("(No group memberships)"))
              }
              $state.groupList = $updatedGroups
              Show-Modal "Success" "Successfully removed $($user.Name) from group '$selectedGroup'"
            } catch {
              Show-Modal "Error" "Failed to remove from group:`n$($_.Exception.Message)"
            }
          }
        } else {
          Show-Modal "Info" "Please select a group to remove"
        }
      }.GetNewClosure())
      $view.Add($btnRemove)
    }
  }

  ## ----------{ Apply Logic }---------
  $applyLogic = {
    param($user, $state)
    try {
      $changesMade = $false
      if ($state.txtSamAccountName) {
        $newSamAccountName = $state.txtSamAccountName.Text.ToString().Trim()
        if ($newSamAccountName -ne $user.SamAccountName -and -not [string]::IsNullOrWhiteSpace($newSamAccountName)) {
          if ($Script:DemoMode) { $user.SamAccountName = $newSamAccountName
          } else {
            Set-UnifiedObject -ObjectType User -Object $user -Properties @{SamAccountName = $newSamAccountName}
            $user.SamAccountName = $newSamAccountName
          }
          $changesMade = $true
        }
      }
      if ($state.txtUserPrincipalName) {
        $newUPN = $state.txtUserPrincipalName.Text.ToString().Trim()
        if ($newUPN -ne $user.UserPrincipalName -and -not [string]::IsNullOrWhiteSpace($newUPN)) {
          if ($Script:DemoMode) { $user.UserPrincipalName = $newUPN
          } else {
            Set-UnifiedObject -ObjectType User -Object $user -Properties @{ UserPrincipalName = $newUPN }
            $user.UserPrincipalName = $newUPN
          }
          $changesMade = $true
        }
      }
      if ($state.txtDisplayName) {
        $newDisplayName = $state.txtDisplayName.Text.ToString().Trim()
        if ($newDisplayName -ne $user.DisplayName -and -not [string]::IsNullOrWhiteSpace($newDisplayName)) {
          if ($Script:DemoMode) { $user.DisplayName = $newDisplayName
          } else {
            Set-UnifiedObject -ObjectType User -Object $user -Properties @{ DisplayName = $newDisplayName }
            $user.DisplayName = $newDisplayName
          }
          $changesMade = $true
        }
      }
      if ($state.txtEmail) {
        $newEmail = $state.txtEmail.Text.ToString().Trim()
        $currentEmail = if ($user.EmailAddress) { $user.EmailAddress } elseif ($user.mail) { $user.mail } else { "" }
        if ($newEmail -ne $currentEmail) {
          if ($Script:DemoMode) {
            $user.EmailAddress = $newEmail
            $user.mail = $newEmail
          } else {
            Set-UnifiedObject -ObjectType User -Object $user -Properties @{ EmailAddress = $newEmail }
            $user.EmailAddress = $newEmail
          }
          $changesMade = $true
        }
      }
      ## Handle Account Expiration
      if ($state.chkNeverExpires) {
        if ($state.chkNeverExpires.Checked) {
          ## User wants account to never expire
          if ($user.PSObject.Properties.Match('AccountExpirationDate') -and $user.AccountExpirationDate) {
            if ($Script:DemoMode) {
              $user.AccountExpirationDate = $null
            } else {
              Set-UnifiedObject -ObjectType User -Object $user -Properties @{ AccountExpirationDate = $null }
              $user.AccountExpirationDate = $null
            }
            $changesMade = $true
          }
        } else {
          ## User set an expiration date
          $expiryText = $state.txtExpirationDate.Text.ToString().Trim()
          if (-not [string]::IsNullOrWhiteSpace($expiryText)) {
          ## Validate format
          if ($expiryText -notmatch '^\d{4}-\d{2}-\d{2}$') {
            Show-Modal "Error" "Invalid date format. Use YYYY-MM-DD"
            return
          }
      ## Try to parse the date
      try {
        $parts = $expiryText -split '-'
        $year  = [int]$parts[0]
        $month = [int]$parts[1]
        $day   = [int]$parts[2]

        ## Validate ranges
        if ($month -lt 1 -or $month -gt 12) {
          Show-Modal "Error" "Invalid month. Must be 01-12, not $month"
          return
        }
        if ($day -lt 1 -or $day -gt 31) {
          Show-Modal "Error" "Invalid day. Must be 01-31, not $day"
          return
        }
        ## Create the date (this will fail for invalid dates like Feb 30)
        $newExpiryDate = [DateTime]::new($year, $month, $day)
        ## Check if date is different from current
        $currentExpiry = if ($user.PSObject.Properties.Match('AccountExpirationDate')) { $user.AccountExpirationDate } else { $null }
        if ($null -eq $currentExpiry -or $newExpiryDate -ne $currentExpiry) {
          if ($Script:DemoMode) {
            $user.AccountExpirationDate = $newExpiryDate
          } else {
            Set-UnifiedObject -ObjectType User -Object $user -Properties @{ AccountExpirationDate = $newExpiryDate }
            $user.AccountExpirationDate = $newExpiryDate
          }
          $changesMade = $true
          }
          } catch [System.ArgumentOutOfRangeException] {
            Show-Modal "Error" "Invalid date: '$expiryText' is not a valid calendar date (e.g., Feb 30 doesn't exist)"
            return
          } catch {
            Show-Modal "Error" "Failed to parse date: $($_.Exception.Message)"
            return
            }
          }
        }
      }
      if ($changesMade) { Show-Modal "Success" "Changes applied successfully"
      } else { Show-Modal "Info" "No changes to apply"
      }
    } catch { Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)" }
  }
  ## ----------{ Create Dialog }---------
  ## Note: Search tab is auto-added by New-PropertiesDialog
  $tabs = @($generalTab, $accountTab, $addressTab, $profileTab, $organizationTab, $memberOfTab)
  New-PropertiesDialog -Title "User Properties - $($user.Name)" -Width 100 -Height 40 -Tabs $tabs -Data $user -OnApply $applyLogic -IncludeSearchTab $true -SearchTabConfig @{ObjectType='User'}
}

##  ----------{ Helper functions for property dialogs }---------
function Add-LabelAndField {
  param(
    [Parameter(Mandatory)]$View,
    [Parameter(Mandatory)][ref]$Y,
    [Parameter(Mandatory)][string]$Label,
    [Parameter(Mandatory)][string]$FieldName,
    [Parameter(Mandatory)]$State,
    [string]$Value = "",
    [int]$LabelX = 2,
    [int]$FieldX = 20,
    [int]$Width = 60,
    [bool]$ReadOnly = $false,
    [bool]$IsTextField = $true
  )

  $lbl   = [Terminal.Gui.Label]::new($Label)
  $lbl.X = $LabelX
  $lbl.Y = $Y.Value
  $View.Add($lbl)

  if ($IsTextField) {
    $State.$FieldName          = [Terminal.Gui.TextField]::new($Value)
    $State.$FieldName.X        = $FieldX
    $State.$FieldName.Y        = $Y.Value
    $State.$FieldName.Width    = $Width
    $State.$FieldName.ReadOnly = $ReadOnly
  } else {
    $State.$FieldName   = [Terminal.Gui.Label]::new($Value)
    $State.$FieldName.X = $FieldX
    $State.$FieldName.Y = $Y.Value
  }
  $View.Add($State.$FieldName)
  $Y.Value += 1
}

function Add-SectionHeader {
  param(
    [Parameter(Mandatory)]$View,
    [Parameter(Mandatory)][ref]$Y,
    [Parameter(Mandatory)][string]$Text,
    [int]$X = 2,
    [int]$SpaceBefore = 1,
    [int]$SpaceAfter = 2
  )
  $Y.Value += $SpaceBefore
  $lbl      = [Terminal.Gui.Label]::new($Text)
  $lbl.X    = $X
  $lbl.Y    = $Y.Value
  $View.Add($lbl)
  $Y.Value += $SpaceAfter
}

## ----------{ Pretty Themes Selection }----------
function Show-ThemeSelector {
  ## Theme list
  $themes = @("british", "class91", "dark", "db-1980s", "dsb", "gemstones", "intercity-swallow", "irn-bru", "light", "matrix", "ns", "network-southeast", "panam", "procomm", "scotrail", "twa", "viarail", "viarail-soft")

  ## Split into two columns
  $half = [math]::Ceiling($themes.Count / 2)
  $leftThemes  = $themes[0..($half-1)]
  $rightThemes = $themes[$half..($themes.Count-1)]

  ## Determine current theme (case-insensitive)
  $currentTheme = $Script:ThemeMode
  Debug-Log (" Global ThemeMode = ${Global:ThemeMode}") -Type "Tracing"
  Debug-Log (" Current theme for selection = ${currentTheme}") -Type "Tracing"

  $currentIndex = -1
  for ($i = 0; $i -lt $themes.Count; $i++) {
    if ($themes[$i].ToLower() -eq $currentTheme.ToLower()) {
      $currentIndex = $i
      break
    }
  }

  Debug-Log (" Index of current theme in $themes array = ${currentIndex}") -Type "Tracing"
  ## Calculate which column gets the selection
  $leftSelected  = if ($currentIndex -ge 0 -and $currentIndex -lt $leftThemes.Count) { $currentIndex } else { -1 }
  $rightSelected = if ($currentIndex -ge $leftThemes.Count) { $currentIndex - $leftThemes.Count } else { -1 }
  Debug-Log (" LeftSelected = ${leftSelected}, RightSelected = ${rightSelected}") -Type "Tracing"

  ## Create dialog
  $dlg = [Terminal.Gui.Dialog]::new("Select Theme", 60, 16)
  $lbl = [Terminal.Gui.Label]::new("Choose a colour theme:")
  $lbl.X = 2; $lbl.Y = 1
  $dlg.Add($lbl)

  ## Left column RadioGroup
  $rdoLeft = [Terminal.Gui.RadioGroup]::new($leftThemes)
  $rdoLeft.X = 2
  $rdoLeft.Y = 3
  $rdoLeft.SelectedItem = $leftSelected
  $dlg.Add($rdoLeft)

  ## Right column RadioGroup
  $rdoRight = [Terminal.Gui.RadioGroup]::new($rightThemes)
  $rdoRight.X = 32
  $rdoRight.Y = 3
  $rdoRight.SelectedItem = $rightSelected
  $dlg.Add($rdoRight)
  Debug-Log (" rdoLeft.SelectedItem = ${rdoLeft.SelectedItem}, rdoRight.SelectedItem = ${rdoRight.SelectedItem}") -Type "Tracing"

  ## Sync columns so only one can be selected
  $rdoLeft.add_SelectedItemChanged({ if ($rdoLeft.SelectedItem -ge 0) { $rdoRight.SelectedItem = -1 } })
  $rdoRight.add_SelectedItemChanged({ if ($rdoRight.SelectedItem -ge 0) { $rdoLeft.SelectedItem = -1  }})

  ## Apply Button
  $btnApply = [Terminal.Gui.Button]::new("Apply")
  $btnApply.add_Clicked({
    $sel = if ($rdoLeft.SelectedItem -ge 0) { $leftThemes[$rdoLeft.SelectedItem]
    } elseif ($rdoRight.SelectedItem -ge 0) { $rightThemes[$rdoRight.SelectedItem]
    } else { "dark"
    }

    Debug-Log "Theme selected on Apply = ${sel}" -Type "Tracing"
    Debug-Log "Switching to theme: ${sel}" -Type "Insight"
    $Script:ThemeMode = $sel
    $newTheme = Get-Theme -mode $sel  # Get NEW theme
    ## Apply to ALL components including tree frame
    Apply-Theme -ThemeData $newTheme -TopLevel [Terminal.Gui.Application]::Top -MainWindow $win -Menu $menu -Status $Script:StatusBar
    ## Apply to tree frame specifically
    if ($treeFrame -and $newTheme.MainWindow) { $treeFrame.ColorScheme = $newTheme.MainWindow }
    ## Apply to filter panel
    if ($filterPanel -and $newTheme.MainWindow) { $filterPanel.ColorScheme = $newTheme.MainWindow }
    ## Apply to selected objects panel
    if ($selectedObjectsPanel -and $newTheme.MainWindow) { $selectedObjectsPanel.ColorScheme = $newTheme.MainWindow }

    ## Force redraw everything
    [Terminal.Gui.Application]::Top.SetNeedsDisplay()
    [Terminal.Gui.Application]::Refresh()
    Show-Modal "Theme Changed" "Theme changed to: ${sel}"
    [Terminal.Gui.Application]::RequestStop()
  })
  $dlg.AddButton($btnApply)

  ## Cancel Button
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $dlg.AddButton($btnCancel)

  ## Run the dialog
  [Terminal.Gui.Application]::Run($dlg)
}

## Allow exporting of data in various formats
function Show-ExportDataDialog {
  <#
  .SYNOPSIS
  Export demo data in various formats
  #>

  Debug-Log "Opening Export Data dialog" -Type "Insight"

  ## Capture functions
  $debugLogFunc = ${function:Debug-Log}
  $showModalFunc = ${function:Show-Modal}

  ## Generate default filename from forest/domain name
  $defaultFilename = if ($Script:ForestName) {
    "$($Script:ForestName)_ad_data"
  } elseif ($Script:CurrentDomain) {
    "$($Script:CurrentDomain)_ad_data"
  } else {
    "ad_export"
  }

  ## Create dialog
  $dlg = [Terminal.Gui.Dialog]::new("Export Data", 70, 28)
  $y = 1

  ## Header
  $lblHeader = [Terminal.Gui.Label]::new("Select export format and object types:")
  $lblHeader.X = 2
  $lblHeader.Y = $y
  $dlg.Add($lblHeader)
  $y += 2

  ## Format selection
  $lblFormat = [Terminal.Gui.Label]::new("Export Format:")
  $lblFormat.X = 2
  $lblFormat.Y = $y
  $dlg.Add($lblFormat)
  $y += 1

  $rdoFormat = [Terminal.Gui.RadioGroup]::new()
  $rdoFormat.X = 4
  $rdoFormat.Y = $y
  $rdoFormat.RadioLabels = [NStack.ustring[]]@(
    "TDF (PowerShell Data) - For re-import",
    "CSVDE (AD Standard) - csvde.exe compatible",
    "Simple CSV - Basic spreadsheet format"
  )
  $rdoFormat.SelectedItem = 0
  $dlg.Add($rdoFormat)
  $y += 4

  ## Object type selection
  $lblObjects = [Terminal.Gui.Label]::new("Include Object Types:")
  $lblObjects.X = 2
  $lblObjects.Y = $y
  $dlg.Add($lblObjects)
  $y += 1

  $chkUsers = [Terminal.Gui.CheckBox]::new("Users ($($Script:Users.Count))")
  $chkUsers.X = 4
  $chkUsers.Y = $y
  $chkUsers.Checked = $true
  $dlg.Add($chkUsers)
  $y += 1

  $chkGroups = [Terminal.Gui.CheckBox]::new("Groups ($($Script:Groups.Count))")
  $chkGroups.X = 4
  $chkGroups.Y = $y
  $chkGroups.Checked = $true
  $dlg.Add($chkGroups)
  $y += 1

  $chkComputers = [Terminal.Gui.CheckBox]::new("Computers ($($Script:Computers.Count))")
  $chkComputers.X = 4
  $chkComputers.Y = $y
  $chkComputers.Checked = $true
  $dlg.Add($chkComputers)
  $y += 1

  $chkDCs = [Terminal.Gui.CheckBox]::new("Domain Controllers ($($Script:DCs.Count))")
  $chkDCs.X = 4
  $chkDCs.Y = $y
  $chkDCs.Checked = $true
  $dlg.Add($chkDCs)
  $y += 2

  ## Filename input
  $lblFilename = [Terminal.Gui.Label]::new("Filename:")
  $lblFilename.X = 2
  $lblFilename.Y = $y
  $dlg.Add($lblFilename)

  $txtFilename = [Terminal.Gui.TextField]::new($defaultFilename)
  $txtFilename.X = 13
  $txtFilename.Y = $y
  $txtFilename.Width = 40
  $dlg.Add($txtFilename)

  $lblExtension = [Terminal.Gui.Label]::new(".csv")
  $lblExtension.X = 54
  $lblExtension.Y = $y
  $dlg.Add($lblExtension)
  $y += 2

  ## Update extension when format changes
  $rdoFormat.add_SelectedItemChanged({
    $ext = switch ($rdoFormat.SelectedItem) {
      0 { ".tdf" }
      1 { ".csv" }
      2 { ".csv" }
      default { ".csv" }
    }
    $lblExtension.Text = $ext
  }.GetNewClosure())

  ## Status label
  $lblStatus = [Terminal.Gui.Label]::new("")
  $lblStatus.X = 2
  $lblStatus.Y = $y
  $lblStatus.Width = [Terminal.Gui.Dim]::Fill(2)
  $dlg.Add($lblStatus)

  ## CAPTURE variables BEFORE the closure
  $capturedUsers = $Script:Users
  $capturedGroups = $Script:Groups
  $capturedComputers = $Script:Computers
  $capturedDCs = $Script:DCs

  ## Save button
  $btnSave = [Terminal.Gui.Button]::new("Save")
  $btnSave.X = 2
  $btnSave.Y = [Terminal.Gui.Pos]::AnchorEnd(2)
  $btnSave.add_Clicked({
    ## Validate filename
    $filename = $txtFilename.Text.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($filename)) {
      & $showModalFunc -title "Invalid Filename" -msg "Please enter a filename."
      return
    }

    ## Collect objects from CAPTURED variables
    $objectsToExport = @()
    if ($chkUsers.Checked) { $objectsToExport += $capturedUsers }
    if ($chkGroups.Checked) { $objectsToExport += $capturedGroups }
    if ($chkComputers.Checked) { $objectsToExport += $capturedComputers }
    if ($chkDCs.Checked) { $objectsToExport += $capturedDCs }

    if ($objectsToExport.Count -eq 0) {
      & $showModalFunc -title "No Selection" -msg "Please select at least one object type to export."
      return
    }

    ## Determine format and extension
    $format = switch ($rdoFormat.SelectedItem) {
      0 { "TDF" }
      1 { "CSVDE" }
      2 { "CSV" }
      default { "CSV" }
    }

    $extension = switch ($format) {
      "TDF" { ".tdf" }
      "CSVDE" { ".csv" }
      "CSV" { ".csv" }
    }

    ## Build full path
    $fullPath = Join-Path (Get-Location) "$filename$extension"

    ## Check if file exists
    if (Test-Path $fullPath) {
      $overwrite = & $showModalFunc -title "Overwrite?" -msg "File already exists:`n`n$fullPath`n`nOverwrite?" -YesNo
      if ($overwrite -ne 0) { return }
    }

    try {
      $lblStatus.Text = "Exporting $($objectsToExport.Count) objects..."
      [Terminal.Gui.Application]::Refresh()

      switch ($format) {
        "TDF" {
          ## Export as PowerShell data file (clixml)
          $objectsToExport | Export-Clixml -Path $fullPath -Depth 10 -Force
          & $debugLogFunc "Exported TDF format" -Type "Success"
        }
        "CSVDE" {
          & $debugLogFunc "Exporting CSVDE format" -Type "Insight"
          ## CSVDE requires DN as first column and all AD attributes
          $csvdeObjects = foreach ($obj in $objectsToExport) {
            $record = [ordered]@{
              ## Required Core Properties (DN must be first)
              DN                 = $obj.distinguishedName
              objectClass        = $obj.objectClass
              distinguishedName  = $obj.distinguishedName
              instanceType       = if ($obj.instanceType) { $obj.instanceType } else { 4 }
              name               = $obj.Name
              cn                 = if ($obj.cn) { $obj.cn } else { $obj.Name }
              ## User Identity
              sAMAccountName     = $obj.sAMAccountName
              userPrincipalName  = $obj.userPrincipalName
              displayName        = $obj.displayName
              description        = $obj.description
              ## Person Details
              givenName          = $obj.givenName
              sn                 = $obj.sn
              mail               = if ($obj.mail) { $obj.mail } else { "" }
              title              = $obj.title
              department         = if ($obj.Department) { $obj.Department } else { "" }
              company            = if ($obj.Company) { $obj.Company } else { "" }
              ## Timestamps (AD format: yyyyMMddHHmmss.0Z)
              whenCreated        = if ($obj.whenCreated -and $obj.whenCreated -is [DateTime]) {
                $obj.whenCreated.ToString('yyyyMMddHHmmss.0Z')
              } else { "" }
              whenChanged        = if ($obj.whenChanged -and $obj.whenChanged -is [DateTime]) {
                $obj.whenChanged.ToString('yyyyMMddHHmmss.0Z')
              } else { "" }
              ## Security & GUIDs
              objectGUID         = if ($obj.objectGUID) { $obj.objectGUID } else { "" }
              objectSid          = if ($obj.objectSid) { $obj.objectSid } else { "" }
              userAccountControl = if ($obj.userAccountControl) { $obj.userAccountControl } else { "" }
              ## Group & Access Control
              primaryGroupID     = if ($obj.primaryGroupID) {
                $obj.primaryGroupID
              } elseif ($obj.objectClass -eq 'user') {
                "513"
              } else { "" }
              adminCount         = if ($obj.adminCount) { $obj.adminCount } else { "" }
              ## Password & Logon Timestamps (FileTime format)
              pwdLastSet         = if ($obj.pwdLastSet) { $obj.pwdLastSet } else { "" }
              lastLogon          = if ($obj.lastLogon) { $obj.lastLogon } else { "" }
              lastLogonTimestamp = if ($obj.lastLogonTimestamp) { $obj.lastLogonTimestamp } else { "" }
              logonCount         = if ($obj.logonCount) { $obj.logonCount } else { "0" }
              ## Password Security
              badPwdCount        = if ($obj.badPwdCount) { $obj.badPwdCount } else { "0" }
              badPasswordTime    = if ($obj.badPasswordTime) { $obj.badPasswordTime } else { "0" }
              ## Account Expiration & Lockout
              accountExpires     = if ($obj.accountExpires) { $obj.accountExpires } else { "0" }
              lockoutTime        = if ($obj.lockoutTime) { $obj.lockoutTime } else { "0" }
              ## Group Membership (semicolon-separated DNs)
              memberOf           = if ($obj.memberOf) {
                if ($obj.memberOf -is [array]) {
                  $obj.memberOf -join ';'
                } else {
                  $obj.memberOf
                }
              } else { "" }
              member             = if ($obj.member) {
                if ($obj.member -is [array]) {
                  $obj.member -join ';'
                } else {
                  $obj.member
                }
              } else { "" }
              groupType             = if ($obj.objectClass -eq 'group') {
                if ($obj.groupType) { $obj.groupType } else { "-2147483646" }
              } else { "" }
              ## Computer Properties
              operatingSystem        = if ($obj.operatingSystem) { $obj.operatingSystem } else { "" }
              operatingSystemVersion = if ($obj.operatingSystemVersion) {
                $obj.operatingSystemVersion
              } else { "" }
              dNSHostName          = if ($obj.dNSHostName) { $obj.dNSHostName } else { "" }
              servicePrincipalName = if ($obj.servicePrincipalName) {
                if ($obj.servicePrincipalName -is [array]) {
                  $obj.servicePrincipalName -join ';'
                } else {
                  $obj.servicePrincipalName
                }
              } else { "" }
              ## Object Category
              objectCategory     = if ($obj.objectCategory) {
                $obj.objectCategory
              } else {
                switch ($obj.objectClass) {
                  'user'     { "CN=Person,CN=Schema,CN=Configuration,DC=example,DC=com" }
                  'computer' { "CN=Computer,CN=Schema,CN=Configuration,DC=example,DC=com" }
                  'group'    { "CN=Group,CN=Schema,CN=Configuration,DC=example,DC=com" }
                  default    { "" }
                }
              }
            }
            [PSCustomObject]$record
          }
          $csvdeObjects | Export-Csv -Path $fullPath -NoTypeInformation -Force -Encoding UTF8
          & $debugLogFunc "Exported CSVDE format with $($csvdeObjects.Count) objects" -Type "Success"
        }
        "CSV" {
          & $debugLogFunc "Exporting Simple CSV format" -Type "Insight"
          ## Simple flattened export
          $simpleObjects = foreach ($obj in $objectsToExport) {
            [PSCustomObject]@{
              Name            = $obj.Name
              Type            = $obj.objectClass
              SamAccountName  = $obj.sAMAccountName
              Email           = $obj.mail
              DisplayName     = $obj.displayName
              Title           = $obj.title
              Department      = $obj.Department
              Enabled         = if ($obj.Enabled -ne $null)   { $obj.Enabled } else { $true }
              Disabled        = if ($obj.Disabled -ne $null)  { $obj.Disabled } else { $false }
              LockedOut       = if ($obj.LockedOut -ne $null) { $obj.LockedOut } else { $false }
              Domain          = $obj.Domain
              OU              = if ($obj.OU) { $obj.OU -join ' > ' } else { "" }
              Groups          = if ($obj.Groups) { $obj.Groups -join '; ' } else { "" }
              Description     = $obj.description
              LastLogon       = $obj.LastLogonDate
              PasswordLastSet = $obj.PasswordLastSet
            }
          }
          $simpleObjects | Export-Csv -Path $fullPath -NoTypeInformation -Force -Encoding UTF8
          & $debugLogFunc "Exported Simple CSV with $($simpleObjects.Count) objects" -Type "Success"
        }
      }
      & $debugLogFunc "Exported $($objectsToExport.Count) objects to ${format}: $fullPath" -Type "Success"
      & $showModalFunc -title "Export Complete" -msg "Successfully exported $($objectsToExport.Count) objects to:`n`n$fullPath"
      [Terminal.Gui.Application]::RequestStop()
    } catch {
      & $debugLogFunc "Failed to export: $($_.Exception.Message)" -Type "Problem"
      & $showModalFunc -title "Export Failed" -msg "Could not export data:`n`n$($_.Exception.Message)"
      $lblStatus.Text = "Export failed: $($_.Exception.Message)"
    }
  }.GetNewClosure())
  $dlg.AddButton($btnSave)

  ## Cancel button
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({
    [Terminal.Gui.Application]::RequestStop()
  }.GetNewClosure())
  $dlg.AddButton($btnCancel)
  [Terminal.Gui.Application]::Run($dlg)
}

## ----------{ Copy User/group (Template-based) }---------
function Copy-ADObject {
  <#
  .SYNOPSIS
  Clone a user or group using an existing object as a template

  .PARAMETER SourceObject
  The user or group to use as a template

  .PARAMETER NewName
  Name for the new object

  .PARAMETER CopyMemberships
  If specified, copies group memberships (users) or members (groups)

  .PARAMETER ShowDialog
  If specified, shows interactive dialog for configuration

  .EXAMPLE
  Copy-ADObject -SourceObject $templateUser -NewName "John Smith" -CopyMemberships

  .EXAMPLE
  Copy-ADObject -SourceObject $user -ShowDialog
  #>

  param(
    [Parameter(Mandatory=$true)]
    [object]$SourceObject,
    [string]$NewName,
    [switch]$CopyMemberships,
    [switch]$ShowDialog
  )

  ## Detect object type
  $isUser  = $SourceObject.PSObject.Properties.Match('SamAccountName') -and -not $SourceObject.PSObject.Properties.Match('ComputerType')
  $isGroup = $SourceObject.PSObject.Properties.Match('Members')

  if (-not $isUser -and -not $isGroup) {
    Show-Modal "Unsupported Type" "Can only copy Users or Groups"
    return
  }

  $objectType = if ($isUser) { 'User' } else { 'Group' }
  Debug-Log "Copy $objectType initiated - Template: $($SourceObject.Name)" -Type "Insight"

  ## ----------{ Interactive dialog }---------
  if ($ShowDialog) {
    $dlg = [Terminal.Gui.Dialog]::new("Copy $objectType - $($SourceObject.Name)", 70, 22)

    ## Info section
    $lblInfo = [Terminal.Gui.Label]::new("Creating new $objectType based on:")
    $lblInfo.X = 2
    $lblInfo.Y = 1
    $dlg.Add($lblInfo)

    $lblTemplate = [Terminal.Gui.Label]::new("Template: $($SourceObject.Name)")
    $lblTemplate.X = 2
    $lblTemplate.Y = 2
    $dlg.Add($lblTemplate)

    ## New name
    $lblName = [Terminal.Gui.Label]::new("New ${objectType} Name:")
    $lblName.X = 2
    $lblName.Y = 4
    $dlg.Add($lblName)
    $txtName = [Terminal.Gui.TextField]::new("")
    $txtName.X = 2
    $txtName.Y = 5
    $txtName.Width = 66

    $dlg.Add($txtName)
    if ($isUser) {
      ## SamAccountName
      $lblSam = [Terminal.Gui.Label]::new("Login Name (SamAccountName):")
      $lblSam.X = 2
      $lblSam.Y = 7
      $dlg.Add($lblSam)
      $txtSam = [Terminal.Gui.TextField]::new("")
      $txtSam.X = 2
      $txtSam.Y = 8
      $txtSam.Width = 66
      $dlg.Add($txtSam)

      ## Auto-generate SamAccountName from display name
      $txtName.add_TextChanged({
        $name = $txtName.Text.ToString()
        if ($name -match '^(\w+)\s+(\w+)') {
          $first = $Matches[1].ToLower()
          $last  = $Matches[2].ToLower()
          $txtSam.Text = [NStack.ustring]::Make("${first}.${last}")
        }
      }.GetNewClosure())

      ## Email
      $lblEmail = [Terminal.Gui.Label]::new("Email Address:")
      $lblEmail.X = 2
      $lblEmail.Y = 10
      $dlg.Add($lblEmail)

      $txtEmail = [Terminal.Gui.TextField]::new("")
      $txtEmail.X = 2
      $txtEmail.Y = 11
      $txtEmail.Width = 66
      $dlg.Add($txtEmail)
      $yNext = 13
      } else {
        $yNext = 7
      }

      ## Copy memberships checkbox
      $chkMemberships = [Terminal.Gui.CheckBox]::new((if ($isUser) { "Copy group memberships" } else { "Copy group members" }))
      $chkMemberships.X = 2
      $chkMemberships.Y = $yNext
      $chkMemberships.Checked = $true
      $dlg.Add($chkMemberships)

      ## Fields to copy section
      $lblFields = [Terminal.Gui.Label]::new("Fields to copy:")
      $lblFields.X = 2
      $lblFields.Y = $yNext + 2
      $dlg.Add($lblFields)

      $lblFieldsList = [Terminal.Gui.Label]::new("")
      $lblFieldsList.X = 2
      $lblFieldsList.Y = $yNext + 3
      $lblFieldsList.Width = 66
      $dlg.Add($lblFieldsList)

      if ($isUser) { $fieldsList = "Department, Title, Company, Manager, OU, Phone, Address"
      } else { $fieldsList = "Description, ManagedBy, OU" }
      $lblFieldsList.Text = [NStack.ustring]::Make($fieldsList)

      ## Create button
      $btnCreate = [Terminal.Gui.Button]::new("Create")
      $btnCreate.add_Clicked({
        $name = $txtName.Text.ToString()

        if ([string]::IsNullOrWhiteSpace($name)) {
          Show-Modal "Missing Name" "Please enter a name for the new $objectType"
         return
        }

        if ($isUser) {
          $sam = $txtSam.Text.ToString()
          $email = $txtEmail.Text.ToString()

          if ([string]::IsNullOrWhiteSpace($sam)) {
            Show-Modal "Missing Login" "Please enter a login name (SamAccountName)"
            return
          }
          [Terminal.Gui.Application]::RequestStop()
          ## Create with additional params
          $params = @{
            NewName         = $name
            SamAccountName  = $sam
            EmailAddress    = $email
            CopyMemberships = $chkMemberships.Checked
          }
          Copy-ADObject -SourceObject $SourceObject @params
          } else {
            [Terminal.Gui.Application]::RequestStop()
            Copy-ADObject -SourceObject $SourceObject -NewName $name -CopyMemberships:$chkMemberships.Checked
          }
      }.GetNewClosure())
      $dlg.AddButton($btnCreate)

      ## Cancel button
      $btnCancel = [Terminal.Gui.Button]::new("Cancel")
      $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
      $dlg.AddButton($btnCancel)
      [Terminal.Gui.Application]::Run($dlg)
      return
  }

  ## ----------{ Direct Creation }---------
  if (-not $NewName) {
    Show-Modal "Missing Name" "NewName parameter required when not using -ShowDialog"
    return
  }

  ## Additional parameters for users (from dialog or direct call)
  $samAccountName = $PSBoundParameters['SamAccountName']
  $emailAddress   = $PSBoundParameters['EmailAddress']

  ## Create new object based on template
  try {
    if ($Script:DemoMode) {
      ## ----------{ Demo Mode - Copy User }---------
      if ($isUser) {
        ## Generate SamAccountName if not provided
        if (-not $samAccountName) {
          if ($NewName -match '^(\w+)\s+(\w+)') { $samAccountName = "$($Matches[1]).$($Matches[2])".ToLower()
          } else { $samAccountName = $NewName.ToLower().Replace(' ', '.') }
        }

        ## Generate email if not provided
        if (-not $emailAddress -and $SourceObject.EmailAddress) {
          $domain = $SourceObject.EmailAddress -replace '^[^@]+@', ''
          $emailAddress = "${samAccountName}@${domain}"
        }

        ## Create new user object
        $newUser = [PSCustomObject]@{
          Name              = $NewName
          SamAccountName    = $samAccountName
          DisplayName       = $NewName
          EmailAddress      = $emailAddress
          Enabled           = $true
          Disabled          = $false
          LockedOut         = $false
          OU                = $SourceObject.OU
          Groups            = if ($CopyMemberships) { $SourceObject.Groups } else { @() }
          MemberOf          = if ($CopyMemberships) { $SourceObject.MemberOf } else { @() }
          Title             = $SourceObject.Title
          Department        = $SourceObject.Department
          Company           = $SourceObject.Company
          Manager           = $SourceObject.Manager
          OfficePhone       = $SourceObject.OfficePhone
          MobilePhone       = $SourceObject.MobilePhone
          StreetAddress     = $SourceObject.StreetAddress
          City              = $SourceObject.City
          PostalCode        = $SourceObject.PostalCode
          Country           = $SourceObject.Country
          UserPrincipalName = "${samAccountName}@$($Script:CurrentDomain)"
          DistinguishedName = "CN=$NewName,$($SourceObject.DistinguishedName -replace '^CN=[^,]+,', '')"
          Domain            = $Script:CurrentDomain
        }

      $Script:Users += $newUser
      $Script:rawUsers += $newUser

      Debug-Log "Created user '$NewName' (demo mode) - Template: $($SourceObject.Name)" -Type "Success"
      if ($CopyMemberships) { Debug-Log "Copied $($newUser.Groups.Count) group memberships" -Type "Insight" }
      Show-Modal "User Created" "Successfully created user '$NewName'`n`nLogin: $samAccountName`nEmail: $emailAddress$(if ($CopyMemberships) { "`n`nCopied $($newUser.Groups.Count) group memberships" } else { '' })"

      ## ----------{ Demo mode - copy group }---------
      } else {
        ## Create new group object
        $newGroup = [PSCustomObject]@{
          Name              = $NewName
          Description       = $SourceObject.Description
          Email             = $SourceObject.Email
          ManagedBy         = $SourceObject.ManagedBy
          Members           = if ($CopyMemberships) { $SourceObject.Members } else { @() }
          MemberOf          = $SourceObject.MemberOf
          DistinguishedName = "CN=$NewName,$($SourceObject.DistinguishedName -replace '^CN=[^,]+,', '')"
          Domain            = $Script:CurrentDomain
        }
        $Script:Groups += $newGroup
        $Script:rawDemoGroups += $newGroup
        Debug-Log "Created group '$NewName' (demo mode) - Template: $($SourceObject.Name)" -Type "Success"
        if ($CopyMemberships) { Debug-Log "Copied $($newGroup.Members.Count) members" -Type "Insight" }
        Show-Modal "Group Created" "Successfully created group '$NewName'$(if ($CopyMemberships) { "`n`nCopied $($newGroup.Members.Count) members" } else { '' })"
      }
    } else {
      ## ----------{ Production Mode }---------
      if ($isUser) {
        ## Generate SamAccountName if not provided
        if (-not $samAccountName) {
          if ($NewName -match '^(\w+)\s+(\w+)') { $samAccountName = "$($Matches[1]).$($Matches[2])".ToLower()
          } else { $samAccountName = $NewName.ToLower().Replace(' ', '.') }
        }
        ## Create user params
        $params = @{
          Name           = $NewName
          SamAccountName = $samAccountName
          DisplayName    = $NewName
          Path           = $SourceObject.DistinguishedName -replace '^CN=[^,]+,', ''
          Enabled        = $true
          ErrorAction    = 'Stop'
        }

        ## Copy standard fields
        if ($SourceObject.Title)         { $params['Title']         = $SourceObject.Title }
        if ($SourceObject.Department)    { $params['Department']    = $SourceObject.Department }
        if ($SourceObject.Company)       { $params['Company']       = $SourceObject.Company }
        if ($SourceObject.Manager)       { $params['Manager']       = $SourceObject.Manager }
        if ($SourceObject.OfficePhone)   { $params['OfficePhone']   = $SourceObject.OfficePhone }
        if ($SourceObject.StreetAddress) { $params['StreetAddress'] = $SourceObject.StreetAddress }
        if ($SourceObject.City)          { $params['City']          = $SourceObject.City }
        if ($SourceObject.PostalCode)    { $params['PostalCode']    = $SourceObject.PostalCode }
        if ($SourceObject.Country)       { $params['Country']       = $SourceObject.Country }
        if ($emailAddress)               { $params['EmailAddress']  = $emailAddress }

        ## Create user
        New-ADUser @params

        ## Copy group memberships
        if ($CopyMemberships) {
          $sourceGroups = Get-ADPrincipalGroupMembership -Identity $SourceObject.SamAccountName | Where-Object { $_.Name -ne 'Domain Users' }
          foreach ($group in $sourceGroups) { Add-ADGroupMember -Identity $group.SamAccountName -Members $samAccountName -ErrorAction SilentlyContinue }
          Debug-Log "Copied $($sourceGroups.Count) group memberships" -Type "Insight"
        }
        Debug-Log "Created user '$NewName' in AD - Template: $($SourceObject.Name)" -Type "Success"
        Show-Modal "User Created" "Successfully created user '$NewName' in AD$(if ($CopyMemberships) { "`n`nCopied $($sourceGroups.Count) group memberships" } else { '' })"
      } else {
        ## Create group
        $params = @{
          Name        = $NewName
          Path        = $SourceObject.DistinguishedName -replace '^CN=[^,]+,', ''
          GroupScope  = 'Global'
          ErrorAction = 'Stop'
        }
        if ($SourceObject.Description) { $params['Description'] = $SourceObject.Description }
        if ($SourceObject.ManagedBy) { $params['ManagedBy'] = $SourceObject.ManagedBy }
        New-ADGroup @params
        ## Copy members
        if ($CopyMemberships) {
          $sourceMembers = Get-ADGroupMember -Identity $SourceObject.Name
          foreach ($member in $sourceMembers) { Add-ADGroupMember -Identity $NewName -Members $member.SamAccountName -ErrorAction SilentlyContinue }
          Debug-Log "Copied $($sourceMembers.Count) members" -Type "Insight"
        }
        Debug-Log "Created group '$NewName' in AD - Template: $($SourceObject.Name)" -Type "Success"
        Show-Modal "Group Created" "Successfully created group '$NewName' in AD$(if ($CopyMemberships) { "`n`nCopied $($sourceMembers.Count) members" } else { '' })"
      }
    }
    ## Refresh UI
    Refresh-Data -domain $Script:CurrentDomain
    Build-Tree -domain $Script:CurrentDomain
    if ($Script:FilterStatusLabel) { Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel }
  } catch {
    Debug-Log "Failed to create $objectType '$NewName': $($_.Exception.Message)" -Type "Problem"
    Show-Modal "Creation Failed" "Failed to create $objectType '$NewName':`n`n$($_.Exception.Message)"
  }
}

## Fail-safe Demo data
function Load-DefaultDemoData {
  <#
  .SYNOPSIS
  Stub function that prompts user to load demo data from a TDF file

  .DESCRIPTION
  This function has been replaced by TDF file loading. Previously contained 3000+
  lines of hardcoded demo data.
  Now automatically shows file picker to load a TDF/CSV file.
  #>

  Debug-Log "Load-DefaultDemoData called - prompting for TDF file" -Type "Insight"
  ## Set minimal scaffold so the app doesn't crash if user cancels
  if (-not $Script:ForestName) {
    $Script:ForestName    = "example.com"
    $Script:RootDomain    = "example.com"
    $Script:Domains       = @('example.com')
    $Script:CurrentDomain = "example.com"
    $Script:Sites         = @('Default-First-Site-Name')
  }

  ## Initialise empty arrays
  $Script:Users     = @()
  $Script:DCs       = @()
  $Script:Computers = @()
  $Script:Groups    = @()
  $Script:ADObjects = @()

  ## Set data source info
  $Script:DataSource = "No Data Loaded"
  $Script:DataSourceInfo = @{
    Source        = "None"
    CSVPath       = $null
    Server        = $null
    LoadedAt      = Get-Date
    IsReadOnly    = $true
    ObjectCounts  = @{
      Users       = 0
      Groups      = 0
      Computers   = 0
      DCs         = 0
    }
  }

  ## Automatically show the file picker
  Debug-Log "Showing file picker for demo data" -Type "Insight"
  $Script:selectedFile = $null
  Show-FileBrowserDialog -Mode 'Open'

  if ($Script:selectedFile -and (Test-Path -LiteralPath $Script:selectedFile)) {
    Debug-Log "User selected file: $($Script:selectedFile)" -Type "Insight"
    $result = Import-DataFile -FilePath $Script:selectedFile
    if ($result) {
      Debug-Log "Demo data loaded successfully from file" -Type "Success"
      Show-InfoPanel -UpdateOnly
      Refresh-Data -Domain $Script:CurrentDomain -RebuildTree
      return $true  ## Success!
    } else {
      Debug-Log "Failed to import data from file" -Type "Problem"
      Show-Modal "Import Failed" "Failed to load demo data from selected file"
      return $false
    }
  } else {
    Debug-Log "No file selected - demo mode has no data" -Type "Warning"
    Show-Modal "Demo Mode" "Demo mode requires a data file.`n`nPlease use File → Import Demo Data to load a TDF or CSV file."
    return $false
  }
}

## ----------{ Import Data (CSV, JSONC, or PS1) }----------
function Import-DataFile {
  param([string]$FilePath)

  Debug-Log "----------{Import Data File Called }----------" -Type "Insight"
  Debug-Log "FilePath parameter: $FilePath" -Type "Insight"
  if (-not $FilePath) {
    Debug-Log "ERROR - FilePath is empty!" -Type "Problem"
    return $false
  }
  if (-not (Test-Path $FilePath)) {
    Debug-Log "ERROR - File does not exist: $FilePath" -Type "Problem"
    return $false
  }

  ## Detect file type
  $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
  Debug-Log "Importing data from: $FilePath (Type: $extension)" -Type "Insight"

  if ($extension -eq '.ps1' -or $extension -eq '.psd1' -or $extension -eq '.tdf') {
    ## ----------{PowerShell Data File Import }----------
    try {
      Debug-Log "Loading PowerShell data file: $FilePath" -Type "Insight"

      $Script:rawUsers      = $null
      $Script:rawDemoGroups = $null
      $Script:rawComputers  = $null
      $Script:rawDCs        = $null

      Debug-Log "Reading and executing TDF file..." -Type "Tracing"

      ## Read the file content
      $fileContent = Get-Content -Path $FilePath -Raw

      ## Execute the content in the current scope to set Script: variables
      $null = Invoke-Expression $fileContent

      Debug-Log "File executed, checking variables..." -Type "Tracing"
      Debug-Log "rawUsers exists: $($null -ne $Script:rawUsers), count: $($Script:rawUsers.Count)" -Type "Tracing"
      Debug-Log "rawDemoGroups exists: $($null -ne $Script:rawDemoGroups), count: $($Script:rawDemoGroups.Count)" -Type "Tracing"
      Debug-Log "rawComputers exists: $($null -ne $Script:rawComputers), count: $($Script:rawComputers.Count)" -Type "Tracing"
      Debug-Log "rawDCs exists: $($null -ne $Script:rawDCs), count: $($Script:rawDCs.Count)" -Type "Tracing"

      $users     = if ($Script:rawUsers)      { $Script:rawUsers }      else { @() }
      $groups    = if ($Script:rawDemoGroups) { $Script:rawDemoGroups } else { @() }
      $computers = if ($Script:rawComputers)  { $Script:rawComputers }  else { @() }
      $dcs       = if ($Script:rawDCs)        { $Script:rawDCs }        else { @() }

      ## **Store the imported raw data for refresh **
      $Script:ImportedRawData = @{
        Users     = $users
        Groups    = $groups
        Computers = $computers
        DCs       = $dcs
      }
      $Script:ImportSource = 'File'
      Debug-Log "Loaded from PS1 - Users: $($users.Count), Groups: $($groups.Count), Computers: $($computers.Count), DCs: $($dcs.Count)" -Type "Insight"
      ## FIXED: Properly detect domain from hashtable keys
      $importedDomain = 'example.com'  # Default fallback

      if ($users.Count -gt 0) {
        Debug-Log "Checking user[0] for domain - Type: $($users[0].GetType().Name)" -Type "Tracing"
        if ($users[0] -is [hashtable]) {
          Debug-Log "User[0] is hashtable - Keys: $($users[0].Keys -join ', ')" -Type "Tracing"
          if ($users[0].ContainsKey('Domain') -and $users[0]['Domain']) {
            $importedDomain = $users[0]['Domain']
            Debug-Log "Found domain in user[0]: '$importedDomain'" -Type "Success"
          }
        }
      }
      if ($importedDomain -eq 'example.com' -and $dcs.Count -gt 0) {
        Debug-Log "Checking DC[0] for domain - Type: $($dcs[0].GetType().Name)" -Type "Tracing"
        if ($dcs[0] -is [hashtable] -and $dcs[0].ContainsKey('Domain') -and $dcs[0]['Domain']) {
          $importedDomain = $dcs[0]['Domain']
          Debug-Log "Found domain in DC[0]: '$importedDomain'" -Type "Success"
        }
      }
      if ($importedDomain -eq 'example.com' -and $computers.Count -gt 0) {
        Debug-Log "Checking computer[0] for domain - Type: $($computers[0].GetType().Name)" -Type "Tracing"
        if ($computers[0] -is [hashtable] -and $computers[0].ContainsKey('Domain') -and $computers[0]['Domain']) {
          $importedDomain = $computers[0]['Domain']
          Debug-Log "Found domain in computer[0]: '$importedDomain'" -Type "Success"
        }
      }
      Debug-Log "PS1 - Domain detected: '$importedDomain' | Users: $($users.Count), Groups: $($groups.Count), Computers: $($computers.Count), DCs: $($dcs.Count)" -Type "Insight"
    } catch {
      Debug-Log "Failed to import PowerShell data file: $($_.Exception.Message)" -Type "Problem"
      Debug-Log "Stack trace: $($_.ScriptStackTrace)" -Type "Problem"
      Show-Modal "PS1 Import Failed" "Could not import PowerShell data file:`n$($_.Exception.Message)"
      return $false
    }

  } elseif ($extension -eq '.jsonc' -or $extension -eq '.json') {
    ## ----------{JSONC Import }----------
    try {
      $rawContent = Get-Content -Path $FilePath -Raw
      $cleanedContent = $rawContent -split "`n" | ForEach-Object { $_ -replace '##.*$', '' } | Where-Object { $_.Trim() -ne '' } | Out-String
      $jsonData = $cleanedContent | ConvertFrom-Json
      $users = @()
      if ($jsonData.users) {
        $users = $jsonData.users | ForEach-Object {
          $h = @{}
          $_.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }
          $h
        }
      }
      $groups = @()
      if ($jsonData.groups) {
        $groups = $jsonData.groups | ForEach-Object {
          $h = @{}
          $_.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }
          $h
        }
      }
      $computers = @()
      if ($jsonData.computers) {
        $computers = $jsonData.computers | ForEach-Object {
          $h = @{}
          $_.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }
          $h
        }
      }
      $dcs = @()
      if ($jsonData.domainControllers) {
        $dcs = $jsonData.domainControllers | ForEach-Object {
          $h = @{}
          $_.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }
          $h
        }
      }
      ## **Store the imported raw data for refresh **
      $Script:ImportedRawData = @{
        Users     = $users
        Groups    = $groups
        Computers = $computers
        DCs       = $dcs
      }
      $Script:ImportSource = 'File'

      ## FIXED: Properly detect domain from hashtable keys
      $importedDomain = 'example.com'  ## Default fallback
      if ($users.Count -gt 0 -and $users[0] -is [hashtable] -and $users[0].ContainsKey('Domain') -and $users[0]['Domain']) { $importedDomain = $users[0]['Domain'] }
      elseif ($dcs.Count -gt 0 -and $dcs[0] -is [hashtable] -and $dcs[0].ContainsKey('Domain') -and $dcs[0]['Domain']) { $importedDomain = $dcs[0]['Domain'] }
      elseif ($computers.Count -gt 0 -and $computers[0] -is [hashtable] -and $computers[0].ContainsKey('Domain') -and $computers[0]['Domain']) { $importedDomain = $computers[0]['Domain'] }
      Debug-Log "JSONC - Domain detected: '$importedDomain' | Users: $($users.Count), Groups: $($groups.Count), Computers: $($computers.Count), DCs: $($dcs.Count)" -Type "Insight"
    } catch {
      Debug-Log "Failed to parse JSONC: $($_.Exception.Message)" -Type "Problem"
      Show-Modal "JSONC Import Failed" "Could not parse JSONC file:`n$($_.Exception.Message)"
      return $false
    }
  } elseif ($extension -eq '.csv') {
    ## ----------{CSV Import }----------
    try {
      $csvContent = Import-Csv -Path $FilePath -Encoding UTF8 -ErrorAction Stop
      Debug-Log "Loaded $($csvContent.Count) rows from CSV" -Type "Insight"

      ## Parse CSV rows into users/computers/etc based on objectClass column
      $users     = @()
      $groups    = @()
      $computers = @()
      $dcs       = @()

      foreach ($row in $csvContent) {
        ## Determine object type from objectClass field
        $objectType = if ($row.objectClass) { $row.objectClass
        } elseif ($row.Type) { $row.Type
        } else { 'user'
        }

        ## Convert CSV row to hashtable
        $hashRow = @{}
        $row.PSObject.Properties | ForEach-Object { if ($_.Value) { $hashRow[$_.Name] = $_.Value } }

        ## Extract OU path from DN or distinguishedName field
        $dnField = $null
        foreach ($field in @('DN', 'distinguishedName', 'dn')) {
          if ($hashRow.ContainsKey($field) -and $hashRow[$field]) {
            $dnField = $hashRow[$field]
            break
          }
        }

        if ($dnField) {
          ## Parse DN to extract OU path
          ## Example: "CN=John Smith,OU=Users,OU=IT,OU=Narnia,DC=jukebox,DC=corp"
          ## Extract: ['Cardiff', 'IT', 'Users'] (top-down hierarchy)
          $ouParts = @()
          $dnComponents = $dnField -split '(?<!\\),' | ForEach-Object { $_.Trim() }

          foreach ($component in $dnComponents) {
            if ($component -match '^OU=(.+)$') { $ouParts += $matches[1] }
          }
          ## Reverse to get top-down hierarchy
          if ($ouParts.Count -gt 0) {
            [array]::Reverse($ouParts)
            $hashRow['OU'] = $ouParts
          }
          ## Extract CN (common name) if not already present
          if (-not $hashRow.ContainsKey('Name') -or -not $hashRow['Name']) {
            if ($dnField -match '^CN=([^,]+)') { $hashRow['Name'] = $matches[1] }
          }
        }
        ## Categorize by object type
        switch -Regex ($objectType) {
          '^(user|person|organizationalPerson)$' { $users += $hashRow }
          '^(group|groupOfNames)$' { $groups += $hashRow }
          '^(computer|device)$' { $computers += $hashRow }
          ## Skip domain object itself
          '^(domainDNS)$' {}
          ## OUs are handled through DN parsing
          '^(organizationalUnit)$' {}
          default {
            # Try to infer type from other fields
            if ($hashRow.ContainsKey('sAMAccountName') -or $hashRow.ContainsKey('userPrincipalName')) { $users += $hashRow
            } elseif ($hashRow.ContainsKey('groupType')) { $groups += $hashRow
            }
          }
        }
      }

      ## **Store the imported raw data for refresh **
      $Script:ImportedRawData = @{
        Users     = $users
        Groups    = $groups
        Computers = $computers
        DCs       = $dcs
      }
      $Script:ImportSource = 'File'

      ## FIXED: Properly detect domain from hashtable keys OR from DN field
      $importedDomain = 'example.com'  # Default fallback

      ## First, try to get domain from explicit Domain field
      if ($users.Count -gt 0 -and $users[0] -is [hashtable] -and $users[0].ContainsKey('Domain') -and $users[0]['Domain']) { $importedDomain = $users[0]['Domain'] }
      elseif ($dcs.Count -gt 0 -and $dcs[0] -is [hashtable] -and $dcs[0].ContainsKey('Domain') -and $dcs[0]['Domain']) { $importedDomain = $dcs[0]['Domain'] }
      elseif ($computers.Count -gt 0 -and $computers[0] -is [hashtable] -and $computers[0].ContainsKey('Domain') -and $computers[0]['Domain']) { $importedDomain = $computers[0]['Domain'] }

      ## If still default, try extracting from DN or distinguishedName field
      if ($importedDomain -eq 'example.com') {
        $dnFields = @('DN', 'distinguishedName', 'dn')
        foreach ($obj in ($users + $dcs + $computers + $groups)) {
          if ($obj -is [hashtable]) {
            foreach ($dnField in $dnFields) {
              if ($obj.ContainsKey($dnField) -and $obj[$dnField]) {
                $dn = $obj[$dnField]
                ## Extract domain from DN like "DC=jukebox,DC=corp"
                if ($dn -match 'DC=([^,]+)') {
                  $dcParts = @()
                  $dn -split ',' | Where-Object { $_ -match '^DC=' } | ForEach-Object { $dcParts += ($_ -replace '^DC=', '') }
                  if ($dcParts.Count -gt 0) {
                    $importedDomain = $dcParts -join '.'
                    Debug-Log "Extracted domain from DN: '$importedDomain'" -Type "Success"
                    break
                  }
                }
              }
            }
            if ($importedDomain -ne 'example.com') { break }
          }
        }
      }
      Debug-Log "CSV - Domain detected: '$importedDomain' | Users: $($users.Count), Groups: $($groups.Count), Computers: $($computers.Count), DCs: $($dcs.Count)" -Type "Insight"
    } catch {
      Debug-Log "Failed to import CSV: $($_.Exception.Message)" -Type "Problem"
      Show-Modal "CSV Import Failed" "Could not import CSV file:`n$($_.Exception.Message)"
      return $false
    }
  } else {
    Debug-Log "Unsupported file type: $extension" -Type "Problem"
    Show-Modal "Unsupported File" "Please use .csv, .jsonc, or .tdf files only."
    return $false
  }

  ## ----------{Common Processing }----------
  ## Only set domain/forest info if not already set by the imported file
  if (-not $Script:ForestName) {
    $Script:ForestName = $importedDomain
    Debug-Log "Set ForestName to: '$importedDomain' (not in file)" -Type "Insight"
  } else {
    Debug-Log "ForestName already set by file: '$($Script:ForestName)'" -Type "Insight"
  }
  if (-not $Script:Domains -or $Script:Domains.Count -eq 0) {
    $Script:Domains = @($importedDomain)
    Debug-Log "Set Domains array to: $($Script:Domains -join ', ') (not in file)" -Type "Insight"
  } else {
    Debug-Log "Domains array already set by file: $($Script:Domains -join ', ')" -Type "Insight"
  }
  if (-not $Script:CurrentDomain) {
    $Script:CurrentDomain = if ($Script:RootDomain) { $Script:RootDomain } else { $importedDomain }
    Debug-Log "Set CurrentDomain to: '$($Script:CurrentDomain)' (not in file)" -Type "Insight"
  } else {
    Debug-Log "CurrentDomain already set by file: '$($Script:CurrentDomain)'" -Type "Insight"
  }

  $Script:Domain = $Script:CurrentDomain
  $baseDN = "DC=$($Script:CurrentDomain -replace '\.',',DC=')"
  Convert-DataToADObjects -Users $users -DCs $dcs -Computers $computers -Groups $groups -Domain $Script:CurrentDomain -BaseDN $baseDN

  ## Log final forest/domain configuration
  Debug-Log "========== Import Complete ==========" -Type "Success"
  Debug-Log "Forest Name:    $($Script:ForestName)" -Type "Insight"
  Debug-Log "Root Domain:    $($Script:RootDomain)" -Type "Insight"
  Debug-Log "Current Domain: $($Script:CurrentDomain)" -Type "Insight"
  Debug-Log "All Domains:    $($Script:Domains -join ', ')" -Type "Insight"
  Debug-Log "Domain Count:   $($Script:Domains.Count)" -Type "Insight"
  Debug-Log "Sites:          $($Script:Sites -join ', ')" -Type "Insight"
  return $true
}

function Load-ADData {
  param(
    [string]$Domain = $null
  )

  Debug-Log "Querying live Active Directory..." -Type "Insight"
  if (-not (Get-Module -ListAvailable ActiveDirectory)) {
    Debug-Log "ActiveDirectory module not available" -Type "Problem"
    return $false
  }
  Import-Module ActiveDirectory -ErrorAction Stop

  try {
    $rootDSE    = Get-ADRootDSE
    $baseDN     = $rootDSE.defaultNamingContext
    $domainName = ($baseDN -replace 'DC=','' -replace ',', '.')
    Debug-Log "Connected to domain: $domainName" -Type "Insight"

    ## ----------{ Users }----------
    $adUsers = Get-ADUser -LDAPFilter "(objectClass=user)" -SearchBase $baseDN -Properties * -ResultSetSize $null
    $users = foreach ($u in $adUsers) {
      $dn = $u.DistinguishedName
      $ouParts = @()
      if ($dn -match 'OU=') {
        $ouParts = ($dn -split '(?<!\\),' | Where-Object { $_ -match '^OU=' } | ForEach-Object { $_ -replace '^OU=', '' })
        [array]::Reverse($ouParts)
      }

      $groups = @()
      if ($u.MemberOf) { $groups = $u.MemberOf | ForEach-Object { if ($_ -match 'CN=([^,]+)') { $matches[1] }} }
      $uac = [int]$u.userAccountControl

      @{
        Name               = $u.Name
        SamAccountName     = $u.SamAccountName
        UserPrincipalName  = $u.UserPrincipalName
        Email              = $u.Mail
        Title              = $u.Title
        Department         = $u.Department
        Office             = $u.physicalDeliveryOfficeName
        Phone              = $u.telephoneNumber
        MobilePhone        = $u.Mobile
        Description        = $u.Description
        OU                 = $ouParts
        Groups             = $groups
        Domain             = $domainName
        Disabled           = ($uac -band 0x0002) -ne 0
        Locked             = ($uac -band 0x0010) -ne 0
        MustChangePassword = ($u.pwdLastSet -eq 0)
        Country            = $u.Country
        Manager            = if ($u.Manager -match 'CN=([^,]+)') { $matches[1] } else { '' }
        Company            = $u.Company
      }
    }

    ## ----------{ Groups }----------
    $adGroups = Get-ADGroup -LDAPFilter "(objectClass=group)" -SearchBase $baseDN -Properties * -ResultSetSize $null
    $groups = foreach ($g in $adGroups) {
      $gt = [int]$g.GroupType
      $isSecurity = ($gt -band 0x80000000) -ne 0
      $scope = switch ($gt -band 0xF) {
        1 { 'Global' }
        2 { 'DomainLocal' }
        4 { 'Universal' }
        default { 'Global' }
      }
      @{
        Name        = $g.Name
        Description = $g.Description
        Type        = if ($isSecurity) { 'Security' } else { 'Distribution' }
        Scope       = $scope
        Email       = $g.Mail
        Domain      = $domainName
        ManagedBy   = if ($g.ManagedBy -match 'CN=([^,]+)') { $matches[1] } else { '' }
      }
    }

    ## ----------[ Computers and DCs ]----------
    $adComputers = Get-ADComputer -LDAPFilter "(objectClass=computer)" -SearchBase $baseDN -Properties * -ResultSetSize $null
    $computers   = @()
    $dcs         = @()
    foreach ($c in $adComputers) {
      $dn      = $c.DistinguishedName
      $uac     = [int]$c.userAccountControl
      $ouParts = @()
      if ($dn -match 'OU=') { $ouParts = ($dn -split '(?<!\\),' | Where-Object { $_ -match '^OU=' } | ForEach-Object { $_ -replace '^OU=', '' })
        [array]::Reverse($ouParts)
      }
      $isDC = ($uac -band 8192) -or ($c.ServicePrincipalName -match 'E3514235') -or ($dn -match 'OU=Domain Controllers')
      if ($isDC) {
        $dcs += @{
          Name                   = $c.Name
          SamAccountName         = $c.SamAccountName
          DNSHostName            = $c.DNSHostName
          Site                   = $c.Site
          Location               = $c.Location
          Domain                 = $domainName
          Forest                 = $domainName
          OS                     = $c.OperatingSystem
          OperatingSystemVersion = $c.OperatingSystemVersion
          IPv4Address            = $c.IPv4Address
          Enabled                = $c.Enabled
          IsGlobalCatalog        = ($c.ServicePrincipalName -match 'GC/')
          FSMORoles              = @()
          LastReplication        = Get-Date
          ReplicationHealth      = 'Healthy'
          LastBoot               = $c.LastLogonDate
        }
      } else {
        $computers += @{
          Name                   = $c.Name
          SamAccountName         = $c.SamAccountName
          Type                   = 'Computer'
          Role                   = 'WKS'
          OU                     = $ouParts
          OS                     = $c.OperatingSystem
          OperatingSystemVersion = $c.OperatingSystemVersion
          DNSHostName            = $c.DNSHostName
          Description            = $c.Description
          Enabled                = $c.Enabled
          Domain                 = $domainName
        }
      }
    }
    Debug-Log "AD query complete: Users=$($users.Count), Groups=$($groups.Count), Computers=$($computers.Count), DCs=$($dcs.Count)" -Type "Success"
    ## **Store the AD raw data for refresh**
    $Script:ImportedRawData = @{
      Users     = $users
      Groups    = $groups
      Computers = $computers
      DCs       = $dcs
    }
    $Script:ImportSource = 'ActiveDirectory'

    ## ----------{ Final conversion }----------
    Convert-DataToADObjects -Users $users -Groups $groups -Computers $computers -DCs $dcs -Domain $domainName -BaseDN $baseDN
    $Script:DataSource = "ActiveDirectory"
    $Script:DataSourceInfo = @{
      Source        = "ActiveDirectory"
      Server        = $env:LOGONSERVER
      LoadedAt      = Get-Date
      IsReadOnly    = $false
      ObjectCounts  = @{
        Users       = $Script:Users.Count
        Groups      = $Script:Groups.Count
        Computers   = $Script:Computers.Count
        DCs         = $Script:DCs.Count
      }
    }
    return $true
  }
  catch {
    Debug-Log "Active Directory load failed: $($_.Exception.Message)" -Type "Problem"
    return $false
  }
}

function Set-UnifiedObject {
  <#
  .SYNOPSIS
  Update an AD-backed or in-memory object (User, Group, or Computer)
  #>

  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [ValidateSet('User','Group','Computer')]
    [string]$ObjectType,
    [Parameter(Mandatory)]
    [object]$Object,
    [Parameter(Mandatory)]
    [hashtable]$Properties
  )

  ## Resolve identity property once
  switch ($ObjectType) {
    'User'     { $identity = $Object.SamAccountName ; $label = 'user' }
    'Group'    { $identity = $Object.Name           ; $label = 'group' }
    'Computer' { $identity = $Object.Name           ; $label = 'computer' }
  }

  switch ($Script:DataSource) {
    'ActiveDirectory' {
      ## Call the correct AD cmdlet dynamically
      $cmdlet = "Set-Unified$ObjectType"
      & $cmdlet -Identity $identity @Properties -ErrorAction Stop
      Debug-Log "Updated ${label} in Active Directory: $($Object.Name)" -Type "Success"
    }
    default {
      ## In-memory update (CSV / Generated)
      foreach ($key in $Properties.Keys) { $Object.$key = $Properties[$key] }
      Debug-Log "Updated ${label} in memory (${Script:DataSource} source): $($Object.Name)" -Type "Insight"
    }
  }
}

function Apply-CombinedFilters {
  <#
  .SYNOPSIS
  Apply all active filters from FilterOptions

  .DESCRIPTION
  Central function that applies name filter + quick filter + checkboxes
  Returns filtered user collection
  #>
  param(
    [array]$Users = $Script:Users
  )

  $filtered = $Users
  ## Apply quick filter if set
  if ($Script:FilterOptions.QuickFilter -and $Script:FilterOptions.QuickFilter -ne 'All') { $filtered = Get-LDAPFilteredObject -ObjectType User -FilterType $Script:FilterOptions.QuickFilter -Users $filtered }
  ## Apply enabled/disabled checkboxes (if no quick filter overriding)
  if (-not $Script:FilterOptions.QuickFilter -or $Script:FilterOptions.QuickFilter -eq 'All') { $filtered = $filtered | Where-Object { ($_.Disabled -and $Script:FilterOptions.ShowDisabledUsers) -or (-not $_.Disabled -and $Script:FilterOptions.ShowEnabledUsers) } }
  ## Apply name filter with operator
  $nameFilter = $Script:FilterOptions.NameFilter.Trim()
  if ($nameFilter) {
    $operator = $Script:FilterOptions.NameOperator
    switch ($operator) {
      'Contains'   { $filtered = $filtered | Where-Object {$_.Name -like "*$nameFilter*" -or $_.EmailAddress -like "*$nameFilter*" -or $_.Title -like "*$nameFilter*"} }
      'StartsWith' { $filtered = $filtered | Where-Object {$_.Name -like "$nameFilter*" -or $_.EmailAddress -like "$nameFilter*" -or $_.Title -like "$nameFilter*"} }
      'EndsWith'   { $filtered = $filtered | Where-Object {$_.Name -like "*$nameFilter" -or $_.EmailAddress -like "*$nameFilter" -or $_.Title -like "*$nameFilter"} }
      'Equals'     { $filtered = $filtered | Where-Object {$_.Name -eq $nameFilter -or $_.EmailAddress -eq $nameFilter -or $_.SamAccountName -eq $nameFilter } }
    }
  }
  return $filtered
}

## ----------{ DNS Lookup Tool }---------
function Show-DNSLookupDialog {
  <#
  .SYNOPSIS
  DNS lookup tool with multi-server verification and expected result validation

  .DESCRIPTION
  Provides DNS query functionality with:
  - Multiple record types (A, AAAA, CNAME, MX, TXT, NS, SOA, PTR, SRV, ANY)
  - Custom DNS server specification
  - Expected result validation with visual indicators
  - Multi-server checking (like https://whatsmydns.net)
  - Clear result display with copy functionality

  .EXAMPLE
  Show-DNSLookupDialog
  #>

  Debug-Log "Opening DNS Lookup tool" -Type "Insight"

  ## Check for DnsClient module
  $hasDnsClient = $null -ne (Get-Module -ListAvailable -Name DnsClient)
  if (-not $hasDnsClient) {
    Show-Modal "DnsClient Module Missing" "The DnsClient PowerShell module is required for DNS lookups.`n`nThis module is included with Windows 8/Server 2012 and later.`n`nIt may not be available on this system."
    return
  }

  ## Import DnsClient module if not already loaded
  if (-not (Get-Module -Name DnsClient)) {
    try {
      Import-Module DnsClient -ErrorAction Stop
      Debug-Log "DnsClient module imported successfully" -Type "Success"
    } catch {
      Show-Modal "Module Import Failed" "Failed to import DnsClient module:`n`n$($_.Exception.Message)"
      return
    }
  }

  ## Capture functions and icons
  $debugLogFunc = ${function:Debug-Log}
  $showModalFunc = ${function:Show-Modal}
  $icons = if ($Script:Icons) { $Script:Icons } else { @{ Success = "✔"; Error = "✖" } }

  ## Public DNS Servers for Multi-Server Check
  $publicDNSServers = @(
    @{ Name = "Google Primary"       ; IP = "8.8.8.8"         ; Location = "Global" }
    @{ Name = "Google Secondary"     ; IP = "8.8.4.4"         ; Location = "Global" }
    @{ Name = "Cloudflare Primary"   ; IP = "1.1.1.1"         ; Location = "Global" }
    @{ Name = "Cloudflare Secondary" ; IP = "1.0.0.1"         ; Location = "Global" }
    @{ Name = "Quad9 Primary"        ; IP = "9.9.9.9"         ; Location = "Global" }
    @{ Name = "Quad9 Secondary"      ; IP = "149.112.112.112" ; Location = "Global" }
    @{ Name = "OpenDNS Primary"      ; IP = "208.67.222.222"  ; Location = "Global" }
    @{ Name = "OpenDNS Secondary"    ; IP = "208.67.220.220"  ; Location = "Global" }
  )

  ## ----------{ Lookup Function - Must be defined before buttons }---------
  function Local-DNSLookup {
    param(
      [string]$RecordName,
      [string]$RecordType,
      [string]$DNSServer,
      [string]$ExpectedResult,
      [bool]$Validate,
      [bool]$MultiServer,
      [array]$PublicServers,
      [hashtable]$Icons
    )

    if ([string]::IsNullOrWhiteSpace($RecordName)) { return "Error: Please enter a DNS record name" }

    $results = @()
    $results += "DNS Lookup Results"
    $results += "═════════════════════════════════════════════════════════════════"
    $results += "Record:     $RecordName"
    $results += "Type:       $RecordType"
    $results += "Timestamp:  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $results += ""

    ## Determine which servers to query
    $serversToQuery = @()

    if ($MultiServer) {
      $results += "Checking multiple DNS servers..."
      $results += ""
      $serversToQuery = $PublicServers
    } else {
      if ([string]::IsNullOrWhiteSpace($DNSServer)) {
        $results += "Using system default DNS resolver"
        $results += ""
        $serversToQuery = @(@{ Name = "System Default"; IP = $null; Location = "Local" })
      } else {
        $results += "Using DNS server: $DNSServer"
        $results += ""
        $serversToQuery = @(@{ Name = "Custom Server"; IP = $DNSServer; Location = "Custom" })
      }
    }

    ## Query each server
    foreach ($server in $serversToQuery) {
      if ($MultiServer) {
        $results += "─────────────────────────────────────────────────────────────────"
        $results += "$($server.Name) ($($server.IP)) - $($server.Location)"
        $results += ""
      }
      try {
        ## Build Resolve-DnsName parameters
        $dnsParams = @{
          Name = $RecordName
          ErrorAction = 'Stop'
        }
        ## Add type parameter (skip for ANY)
        if ($RecordType -ne "ANY") { $dnsParams['Type'] = $RecordType }
        ## Add server parameter if specified
        if ($server.IP) { $dnsParams['Server'] = $server.IP }
        ## Execute DNS query
        $dnsResult = Resolve-DnsName @dnsParams
        if ($dnsResult) {
          ## Extract relevant data based on record type
          $recordData = @()

          foreach ($record in $dnsResult) {
            switch ($RecordType) {
              "A"     { if ($record.IP4Address)    { $recordData += $record.IP4Address } }
              "AAAA"  { if ($record.IP6Address)    { $recordData += $record.IP6Address } }
              "CNAME" { if ($record.NameHost)      { $recordData += $record.NameHost } }
              "MX"    { if ($record.NameExchange)  { $recordData += "$($record.Preference) $($record.NameExchange)" } }
              "TXT"   { if ($record.Strings)       { $recordData += ($record.Strings -join ' ') } }
              "NS"    { if ($record.NameHost)      { $recordData += $record.NameHost } }
              "SOA"   { if ($record.PrimaryServer) { $recordData += "Primary: $($record.PrimaryServer), Serial: $($record.SerialNumber)" } }
              "PTR"   { if ($record.NameHost)      { $recordData += $record.NameHost } }
              "SRV"   { if ($record.NameTarget)    { $recordData += "$($record.Priority) $($record.Weight) $($record.Port) $($record.NameTarget)" } }
              "ANY"   {
                if ($record.IP4Address)   { $recordData += "A: $($record.IP4Address)" }
                if ($record.IP6Address)   { $recordData += "AAAA: $($record.IP6Address)" }
                if ($record.NameHost)     { $recordData += "CNAME/NS: $($record.NameHost)" }
                if ($record.NameExchange) { $recordData += "MX: $($record.Preference) $($record.NameExchange)" }
              }
            }
          }

          if ($recordData.Count -gt 0) {
            ## Validation check
            $validationResult = ""
            if ($Validate -and -not [string]::IsNullOrWhiteSpace($ExpectedResult)) {
              $matched = $false
              foreach ($data in $recordData) {
                if ($data -match [regex]::Escape($ExpectedResult)) {
                  $matched = $true
                  break
                }
              }
              if ($matched) {
                $validationResult = "  $($Icons.Success) MATCHES expected result"
              } else {
                $validationResult = "  $($Icons.Error) DOES NOT MATCH expected result (expected: $ExpectedResult)"
              }
            }

            foreach ($data in $recordData) { $results += "  $data" }
            if ($validationResult) {
              $results += ""
              $results += $validationResult
            }
          } else {
            $results += "  (No data returned for this record type)"
          }
        } else {
          $results += "  (No records found)"
        }
      } catch {
        $results += "  $($Icons.Error) Error: $($_.Exception.Message)"
      }
      $results += ""
    }
    $results += "═════════════════════════════════════════════════════════════════"
    $results += "Lookup completed"
    return ($results -join "`n")
  }

  ## Create dialog
  $dialog = [Terminal.Gui.Dialog]::new("DNS Lookup Tool", 120, 40)
  $y = 1

  ## ----------{ DNS Record Input }---------
  $lblRecord = [Terminal.Gui.Label]::new("DNS Record:")
  $lblRecord.X = 2
  $lblRecord.Y = $y
  $dialog.Add($lblRecord)

  $txtRecord = [Terminal.Gui.TextField]::new("")
  $txtRecord.X = 15
  $txtRecord.Y = $y
  $txtRecord.Width = 50
  $dialog.Add($txtRecord)
  $y += 2

  ## ----------{ Record Type Selection }---------
  $lblType = [Terminal.Gui.Label]::new("Record Type:")
  $lblType.X = 2
  $lblType.Y = $y
  $dialog.Add($lblType)

  $cmbType = [Terminal.Gui.ComboBox]::new()
  $cmbType.X = 15
  $cmbType.Y = $y
  $cmbType.Width = 20
  $cmbType.Height = 4
  $cmbType.SetSource([string[]]@("A", "AAAA", "CNAME", "MX", "TXT", "NS", "SOA", "PTR", "SRV", "ANY"))
  $cmbType.Text = "A"
  $dialog.Add($cmbType)
  $y += 2

  ## ----------{ DNS Server Input }---------
  $lblServer = [Terminal.Gui.Label]::new("DNS Server:")
  $lblServer.X = 2
  $lblServer.Y = $y
  $dialog.Add($lblServer)

  $txtServer = [Terminal.Gui.TextField]::new("")
  $txtServer.X = 15
  $txtServer.Y = $y
  $txtServer.Width = 30
  $dialog.Add($txtServer)

  $lblServerHint = [Terminal.Gui.Label]::new("(leave blank for system default)")
  $lblServerHint.X = 46
  $lblServerHint.Y = $y
  $dialog.Add($lblServerHint)
  $y += 2

  ## ----------{ Expected Result Validation }---------
  $chkValidate = [Terminal.Gui.CheckBox]::new("Validate against expected result:")
  $chkValidate.X = 2
  $chkValidate.Y = $y
  $chkValidate.Checked = $false
  $dialog.Add($chkValidate)
  $y += 1

  $txtExpected = [Terminal.Gui.TextField]::new("")
  $txtExpected.X = 4
  $txtExpected.Y = $y
  $txtExpected.Width = 50
  $txtExpected.Enabled = $false
  $dialog.Add($txtExpected)

  $lblExpectedHint = [Terminal.Gui.Label]::new("(e.g., 1.2.3.4 for A record)")
  $lblExpectedHint.X = 55
  $lblExpectedHint.Y = $y
  $dialog.Add($lblExpectedHint)
  $y += 2

  ## Enable/disable expected result field based on checkbox
  $chkValidate.add_Toggled({
    $txtExpected.Enabled = $chkValidate.Checked
  }.GetNewClosure())

  ## ----------{ Multi-Server Check }---------
  $chkMultiServer = [Terminal.Gui.CheckBox]::new("Check multiple DNS servers (like whatsmydns.net)")
  $chkMultiServer.X = 2
  $chkMultiServer.Y = $y
  $chkMultiServer.Checked = $false
  $dialog.Add($chkMultiServer)
  $y += 2

  ## ----------{ Results Display }---------
  $lblResults = [Terminal.Gui.Label]::new("Results:")
  $lblResults.X = 2
  $lblResults.Y = $y
  $dialog.Add($lblResults)
  $y += 1

  $txtResults = [Terminal.Gui.TextView]::new()
  $txtResults.X = 2
  $txtResults.Y = $y
  $txtResults.Width = [Terminal.Gui.Dim]::Fill(2)
  $txtResults.Height = [Terminal.Gui.Dim]::Fill(4)
  $txtResults.ReadOnly = $true
  $txtResults.Text = "Enter a DNS record and click Lookup to begin..."
  $dialog.Add($txtResults)

  ## ----------{ Capture for closures }---------
  $capturedLookupFunc = ${function:Local-DNSLookup}
  $capturedPublicServers = $publicDNSServers
  $capturedIcons = $icons

  ## ----------{ Buttons }---------
  ## Lookup button
  $btnLookup = [Terminal.Gui.Button]::new("Lookup")
  $btnLookup.X = 2
  $btnLookup.Y = [Terminal.Gui.Pos]::AnchorEnd(1)
  $btnLookup.add_Clicked({
    try {
      $txtResults.Text = "Performing DNS lookup..."
      [Terminal.Gui.Application]::Refresh()
      $record      = $txtRecord.Text.ToString().Trim()
      $type        = $cmbType.Text.ToString()
      $server      = $txtServer.Text.ToString().Trim()
      $expected    = $txtExpected.Text.ToString().Trim()
      $validate    = $chkValidate.Checked
      $multiServer = $chkMultiServer.Checked
      $result      = & $capturedLookupFunc -RecordName $record -RecordType $type -DNSServer $server -ExpectedResult $expected -Validate $validate -MultiServer $multiServer -PublicServers $capturedPublicServers -Icons $capturedIcons
      $txtResults.Text = $result
      & $debugLogFunc "DNS lookup completed for $record ($type)" -Type "Success"
    } catch {
      $errorMsg = "DNS lookup failed: $($_.Exception.Message)"
      $txtResults.Text = $errorMsg
      & $debugLogFunc $errorMsg -Type "Problem"
    }
  }.GetNewClosure())
  $dialog.Add($btnLookup)

  ## Clear button
  $btnClear = [Terminal.Gui.Button]::new("Clear")
  $btnClear.X = 14
  $btnClear.Y = [Terminal.Gui.Pos]::AnchorEnd(1)
  $btnClear.add_Clicked({
    $txtRecord.Text = ""
    $txtServer.Text = ""
    $txtExpected.Text = ""
    $cmbType.Text = "A"
    $chkValidate.Checked = $false
    $chkMultiServer.Checked = $false
    $txtExpected.Enabled = $false
    $txtResults.Text = "Enter a DNS record and click Lookup to begin..."
    & $debugLogFunc "DNS lookup form cleared" -Type "Insight"
  }.GetNewClosure())
  $dialog.Add($btnClear)

  ## Copy Results button
  $btnCopy = [Terminal.Gui.Button]::new("Copy Results")
  $btnCopy.X = 24
  $btnCopy.Y = [Terminal.Gui.Pos]::AnchorEnd(1)
  $btnCopy.add_Clicked({
    try {
      $results = $txtResults.Text.ToString()
      if (-not [string]::IsNullOrWhiteSpace($results)) {
        Set-Clipboard -Value $results
        & $showModalFunc "Copied" "Results copied to clipboard!"
        & $debugLogFunc "DNS results copied to clipboard" -Type "Success"
      } else {
        & $showModalFunc "No Results" "No results to copy"
      }
    } catch {
      & $showModalFunc "Copy Failed" "Failed to copy to clipboard: $($_.Exception.Message)"
      & $debugLogFunc "Failed to copy DNS results: $($_.Exception.Message)" -Type "Problem"
    }
  }.GetNewClosure())
  $dialog.Add($btnCopy)

  ## Close button
  $btnClose = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = [Terminal.Gui.Pos]::AnchorEnd(10)
  $btnClose.Y = [Terminal.Gui.Pos]::AnchorEnd(1)
  $btnClose.add_Clicked({
    & $debugLogFunc "DNS Lookup tool closed" -Type "Insight"
    [Terminal.Gui.Application]::RequestStop()
  }.GetNewClosure())
  $dialog.Add($btnClose)

  ## Run dialog
  [Terminal.Gui.Application]::Run($dialog)
}

## ----------{ File Browser }----------
function Show-FileBrowserDialog {
  param(
    [string]$StartDir = ".",
    [string]$Title = "Select File",
    [string[]]$Filter = @("*.*"),
    [ValidateSet('Open','Save')]
    [string]$Mode = 'Open'
  )

  $script:selectedFile = $null
  $script:currentPath = (Resolve-Path $StartDir).Path
  $dialog = [Terminal.Gui.Dialog]::new($Title, 80, 24)

  ## Current path label
  $labelPath = [Terminal.Gui.Label]::new(2, 1, "Path: $($script:currentPath)")
  $labelPath.Width = 74
  $dialog.Add($labelPath)

  ## ListView for files/folders
  $listView = [Terminal.Gui.ListView]::new()
  $listView.X = 2; $listView.Y = 3
  $listView.Width = 74; $listView.Height = 14
  $dialog.Add($listView)

  ## Selected file label
  $labelSelected = [Terminal.Gui.Label]::new(2, 18, "Selected: (none)")
  $labelSelected.Width = 74
  $dialog.Add($labelSelected)

  ## Update file list helper function
  function Update-FileList {
    param([string]$path)

    $script:currentPath = $path
    $labelPath.Text = [NStack.ustring]::Make("Path: $path")
    $items = [System.Collections.Generic.List[string]]::new()
    ## Parent directory
    if ($path -ne [System.IO.Path]::GetPathRoot($path)) { $items.Add("[..]") }
    ## Directories
    try {
      Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | Sort-Object Name | ForEach-Object { $items.Add("[DIR] $($_.Name)") }
      ## Files matching filter
      Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue | Where-Object { $Filter -contains "*.*" -or $Filter -contains "*$($_.Extension)" } | Sort-Object Name | ForEach-Object { $items.Add($_.Name) }
    } catch {
      Show-Modal "Error" "Cannot access directory: $path"
    }
    if ($items.Count -eq 0) { $items.Add("(empty directory)") }
    $listView.SetSource($items)
  }
  ## Shared selection logic (DRY principle)
  $handleSelection = {
    if ($listView.SelectedItem -lt 0) { return }
    $sel = $listView.Source.ToList()[$listView.SelectedItem]
    if ($sel -eq "[..]") {
      $parent = Split-Path -Parent $script:currentPath
      if ($parent) { Update-FileList -path $parent }
    }
    elseif ($sel -match '^\[DIR\] (.+)$') {
      $dirName = $Matches[1]
      $newPath = Join-Path $script:currentPath $dirName
      Update-FileList -path $newPath
    }
    elseif ($sel -ne "(empty directory)") {
      $script:selectedFile = Join-Path $script:currentPath $sel
      $labelSelected.Text = [NStack.ustring]::Make("Selected: $($script:selectedFile)")
    }
  }
  ## Double-click to select
  $listView.add_OpenSelectedItem($handleSelection)
  ## Enter key to select
  $listView.add_KeyPress({
    param($sender, $keyEvent)
    if ($keyEvent.KeyEvent.Key -eq [Terminal.Gui.Key]::Enter) {
      & $handleSelection
      $keyEvent.Handled = $true
    }
  })

  ## Select button
  $btnSelect = [Terminal.Gui.Button]::new(2, 20, "Select")
  $btnSelect.add_Clicked({
    if ($script:selectedFile) {
      Debug-Log "File selected: $($script:selectedFile)" -Type "Insight"
      [Terminal.Gui.Application]::RequestStop()
    } else {
      Show-Modal "No Selection" "Please select a file"
    }
  })
  $dialog.Add($btnSelect)

  ## Cancel button
  $btnCancel = [Terminal.Gui.Button]::new(15, 20, "Cancel")
  $btnCancel.add_Clicked({
    $script:selectedFile = $null
    [Terminal.Gui.Application]::RequestStop()
  })
  $dialog.Add($btnCancel)

  ## Initial population
  Update-FileList -path $script:currentPath

  ## Run dialog
  [Terminal.Gui.Application]::Run($dialog)
  return $script:selectedFile
}

function Initialise-UIFramework {
  <#
  .SYNOPSIS
  Initialise Terminal.Gui application and create main window

  .DESCRIPTION
  Sets up the Terminal.Gui framework, creates the top-level application
  and main window, and applies the selected theme. This should be called
  FIRST before any data loading to ensure the UI is visible.

  .PARAMETER Theme
  Theme name to apply (HighContrast, PanAm, Matrix, etc)

  .PARAMETER Title
  Window title to display

  .EXAMPLE
  $uiComponents = Initialise-UIFramework -Theme "PanAm" -Title "DSA-TUI v1.0"
  $top = $uiComponents.Top
  $win = $uiComponents.Window
  #>

  param(
    [string]$Theme = "HighContrast",
    [string]$Title = "DSA-TUI - Active Directory"
  )
  Debug-Log "Initializing Terminal.Gui framework..." -Type "Insight"

  ## ----------{ Step 1: Initialise Terminal.Gui }---------
  try {
    [Terminal.Gui.Application]::Init()
    Debug-Log "Terminal.Gui.Application Initialised" -Type "Success"
  } catch {
    Debug-Log "FATAL - Failed to Initialise Terminal.Gui: $($_.Exception.Message)" -Type "Problem"
    throw
  }
  ## Get top-level application
  $top = [Terminal.Gui.Application]::Top

  ## ----------{ Step 2: Create Main Window }---------
  $win = [Terminal.Gui.Window]::new($Title)
  $win.X = 0
  $win.Y = 0
  $win.Width  = [Terminal.Gui.Dim]::Fill()
  $win.Height = [Terminal.Gui.Dim]::Fill(1)  ## Leave room for status bar
  Debug-Log "Main window created with title: $Title" -Type "Insight"

  ## ----------{ Step 3: Apply Theme }---------
  Debug-Log "Applying theme: $Theme" -Type "Insight"
  $Script:ThemeMode = $Theme

  try {
    ## Get theme data
    $themeData = Get-Theme -mode $Theme
    if ($themeData) {
      ## Store theme data globally
      $Script:themeData = $themeData
      ## Use Apply-Theme to handle all components properly
      ## Note: Menu and StatusBar don't exist yet, so pass $null
      Apply-Theme -ThemeData $themeData -TopLevel $top -MainWindow $win -Menu $null -Status $null
      Debug-Log "Theme '$Theme' applied successfully" -Type "Success"
    } else { Debug-Log "WARNING - Theme data is null, using defaults" -Type "Warning" }
  } catch {
    Debug-Log "WARNING - Failed to apply theme: $($_.Exception.Message)" -Type "Warning"
  }

  ## ----------{ Step 4: Add Window to Top }---------
  $top.Add($win)
  Debug-Log "Main window added to top-level application" -Type "Success"

  ## ----------{ Step 5: Return Components }---------
  $result = @{
    Top    = $top
    Window = $win
    Theme  = $themeData
  }
  Debug-Log "UI Framework initialization complete" -Type "Success"
  return $result
}

## Show the F12 "right click" popup menu
## ----------[ Context Menu Handler ]----------
function Show-ObjectContextMenu {
  <#
  .SYNOPSIS
  Shows and handles a context menu for AD objects

  .PARAMETER Object
  The AD object (User, Group, Computer, DC, OU)

  .PARAMETER ObjectType
  Type of object: 'User', 'Group', 'Computer', 'DC', 'OU'
  #>
  param(
    [Parameter(Mandatory=$true)]
    [object]$Object,
    [Parameter(Mandatory=$true)]
    [ValidateSet('User', 'Group', 'Computer', 'DC', 'OU')]
    [string]$ObjectType
  )

  Debug-Log "Showing context menu for $($Object.Name) (Type: $ObjectType)" -Type "Insight"
  ## Build menu items based on object type
  $menuItems = [System.Collections.ArrayList]@()

  switch ($ObjectType) {
    'User' {
      [void]$menuItems.Add("Properties")
      [void]$menuItems.Add("Reset Password")
      if ($Object.Enabled) {
        [void]$menuItems.Add("Disable Account")
      } else {
        [void]$menuItems.Add("Enable Account")
      }
      if ($Object.LockedOut -or $Object.Locked) {
        [void]$menuItems.Add("Unlock Account")
      }
      [void]$menuItems.Add("Move to OU...")
      [void]$menuItems.Add("Add to Group...")
      [void]$menuItems.Add("Delete")
      [void]$menuItems.Add("Refresh")
    }
    'Group' {
      [void]$menuItems.Add("Properties")
      [void]$menuItems.Add("Add Member...")
      [void]$menuItems.Add("Remove Member...")
      [void]$menuItems.Add("Delete")
      [void]$menuItems.Add("Refresh")
    }
    'Computer' {
      [void]$menuItems.Add("Properties")
      if ($Object.Enabled) {
        [void]$menuItems.Add("Disable")
      } else {
        [void]$menuItems.Add("Enable")
      }
      [void]$menuItems.Add("Move to OU...")
      [void]$menuItems.Add("LAPS Password...")
      [void]$menuItems.Add("Delete")
      [void]$menuItems.Add("Refresh")
    }
    'DC' {
      [void]$menuItems.Add("Properties")
      [void]$menuItems.Add("Check Replication")
      [void]$menuItems.Add("View FSMO Roles")
      [void]$menuItems.Add("LAPS Password...")
      [void]$menuItems.Add("Refresh")
    }
    'OU' {
      [void]$menuItems.Add("Properties")
      [void]$menuItems.Add("New Object...")
      [void]$menuItems.Add("Delete")
      [void]$menuItems.Add("Refresh")
    }
  }

  Debug-Log "Built menu with $($menuItems.Count) items" -Type "Tracing"

  ## Capture variables AND functions for closure
  $capturedObj       = $Object
  $capturedObjType   = $ObjectType
  $capturedMenuItems = $menuItems

  ## Capture all the functions we'll need
  $debugLogFunc           = ${function:Debug-Log}
  $showUserPropsFunc      = ${function:Show-UserPropertiesDialog}
  $showGroupPropsFunc     = ${function:Show-GroupPropertiesDialog}
  $showComputerPropsFunc  = ${function:Show-ComputerPropertiesDialog}
  $showDCPropsFunc        = ${function:Show-DCPropertiesDialog}
  $showOUPropsFunc        = ${function:Show-OUPropertiesDialog}
  $showResetPwdFunc       = ${function:Show-ResetPasswordDialog}
  $toggleUserFunc         = ${function:Toggle-UserAccount}
  $unlockUserFunc         = ${function:Unlock-UserAccount}
  $toggleComputerFunc     = ${function:Toggle-ComputerAccount}
  $invokeObjOpFunc        = ${function:Invoke-ObjectOperation}
  $showEditGroupFunc      = ${function:Show-EditGroupMembershipDialog}
  $checkDCReplFunc        = ${function:Check-DCReplication}
  $showADHealthFunc       = ${function:Show-ADHealthDialog}
  $showNewObjFunc         = ${function:Show-NewObjectWizard}
  $refreshDataFunc        = ${function:Refresh-Data}
  $invokeBulkAddGroupFunc = ${function:Invoke-BulkAddToGroup}
  $showLAPSFunc           = ${function:Show-LAPSSearchModal}

  ## Create dialog
  $contextDialog   = [Terminal.Gui.Dialog]::new("Actions", 30, ($menuItems.Count + 4))
  $contextDialog.X = [Terminal.Gui.Pos]::Center()
  $contextDialog.Y = [Terminal.Gui.Pos]::Center()

  ## Create list view
  $listView = [Terminal.Gui.ListView]::new()
  $listView.SetSource($menuItems)
  $listView.X      = 0
  $listView.Y      = 0
  $listView.Width  = [Terminal.Gui.Dim]::Fill()
  $listView.Height = [Terminal.Gui.Dim]::Fill(2)
  $contextDialog.Add($listView)

  ## Handle selection
  $listView.add_OpenSelectedItem({
    $selected = $capturedMenuItems[$listView.SelectedItem]
    & $debugLogFunc "Menu item selected: $selected" -Type "Insight"
    [Terminal.Gui.Application]::RequestStop()

    switch ($selected) {
      "Properties" {
        switch ($capturedObjType) {
          'User' {
            & $debugLogFunc "Showing user properties for $($capturedObj.Name)" -Type "Insight"
            & $showUserPropsFunc -user $capturedObj
          }
          'Group' {
            & $debugLogFunc "Showing group properties for $($capturedObj.Name)" -Type "Insight"
            & $showGroupPropsFunc -group $capturedObj
          }
          'Computer' {
            & $debugLogFunc "Showing computer properties for $($capturedObj.Name)" -Type "Insight"
            & $showComputerPropsFunc -computerName $capturedObj.Name
          }
          'DC' {
            & $debugLogFunc "Showing DC properties for $($capturedObj.Name)" -Type "Insight"
            & $showDCPropsFunc -dc $capturedObj
          }
          'OU' {
            & $debugLogFunc "Showing OU properties for $($capturedObj.Name)" -Type "Insight"
            & $showOUPropsFunc -ouname $capturedObj.Name
          }
        }
      }
      "Reset Password"    { & $showResetPwdFunc -userName $capturedObj.Name }
      "Disable Account"   { & $toggleUserFunc -userName $capturedObj.Name -disable $true }
      "Enable Account"    { & $toggleUserFunc -userName $capturedObj.Name -disable $false }
      "Unlock Account"    { & $unlockUserFunc -userName $capturedObj.Name }
      "Disable"           { & $toggleComputerFunc -computerName $capturedObj.Name -disable $true }
      "Enable"            { & $toggleComputerFunc -computerName $capturedObj.Name -disable $false }
      "Move to OU..."     { & $invokeObjOpFunc -Objects @($capturedObj) -Operation 'Move' }
      "Delete"            { & $invokeObjOpFunc -Objects @($capturedObj) -Operation 'Delete' }
      "Add Member..."     { & $showEditGroupFunc -groupName $capturedObj.Name }
      "Remove Member..."  { & $showEditGroupFunc -groupName $capturedObj.Name }
      "Add to Group..."   {
        $Script:SelectedObjects = @($capturedObj.Name)
        & $invokeBulkAddGroupFunc
        $Script:SelectedObjects = @()
      }
      "Check Replication" { & $checkDCReplFunc -dcName $capturedObj.Name }
      "View FSMO Roles"   { & $showADHealthFunc -InitialTab "FSMO Roles" }
      "New Object..."     { & $showNewObjFunc }
      "LAPS Password..."  { & $showLAPSFunc }
      "Refresh"           { & $refreshDataFunc -Domain $Script:CurrentDomain -RebuildTree -ShowModal -ShowLoadingDialog }
    }
  }.GetNewClosure())

  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }.GetNewClosure())
  $contextDialog.AddButton($btnCancel)
  [Terminal.Gui.Application]::Run($contextDialog)
}

function Set-StatusBar {
  <#
  .SYNOPSIS
  Universal status bar management - Initialise, update, and animate

  .DESCRIPTION
  Single function that handles all status bar operations:
  - Initialise: Create status bar with F-key shortcuts and theme
  - Update: Set static messages
  - Icon: Show status with predefined icons
  - Progress: Optional ASCII progress bar [oooo----]
  - Final: Mark operation complete with checkmark

  .PARAMETER Initialise
  Create and configure the status bar (call once at startup)

  .PARAMETER ThemeData
  Theme data to apply (used with -Initialise)

  .PARAMETER Message
  Status message to display

  .PARAMETER Icon
  Predefined icon to show: 'Working', 'Success', 'Error', 'Warning', 'Info'

  .PARAMETER Percent
  Optional progress percentage (0–100). When set, shows ASCII progress bar.

  .PARAMETER BarWidth
  Width of ASCII progress bar (default 20)

  .PARAMETER Spinner
  (Deprecated) Mapped to -Icon 'Working'

  .PARAMETER Final
  Mark operation complete (same as -Icon 'Success')
  #>

  param(
    [Parameter(Position=0)]
    [string]$Message = "",
    [switch]$Initialise,
    [object]$ThemeData = $null,
    [ValidateSet('Working', 'Success', 'Error', 'Warning', 'Info')]
    [string]$Icon,
    [int]$Percent,
    [int]$BarWidth = 20,
    [switch]$Spinner,
    [switch]$Final
  )

  ## ----------{ Initialise mode }---------
  if ($Initialise) {
    Debug-Log "Initializing status bar..." -Type "Insight"

    $Script:StatusIcons = @{
      Working = "⏳"
      Success = "✓"
      Error   = "✗"
      Warning = "▲"
      Info    = "ℹ"
      Default = ">>"
    }
    $Script:StatusItem = [Terminal.Gui.StatusItem]::new(0, "Initializing...", $null)

    $shortcuts = @(
      @{ Key = [Terminal.Gui.Key]::F1;  Label = "~F1~ Help";         Action = { Show-Modal "Shortcuts" "F1 - Help`nF2 - Password Generator`nF3 - New`nF5 - Refresh`nF6 - Themes`nF7 - Search`nF8 - Focus Tree`nF9 - Show Menus`nF10 - Quit`nF11 - Full Screen`nF12 - Show Context Menu" } }
      @{ Key = [Terminal.Gui.Key]::F2;  Label = "~F2~ Password";     Action = { Generate-RandomPassword } }
      @{ Key = [Terminal.Gui.Key]::F3;  Label = "~F3~ New";          Action = { Show-NewObjectWizard } }
      @{ Key = [Terminal.Gui.Key]::F5;  Label = "~F5~ Refresh";      Action = { Refresh-Data -domain $Script:CurrentDomain -RebuildTree } }
      @{ Key = [Terminal.Gui.Key]::F6;  Label = "~F6~ Themes";       Action = { Show-ThemeSelector } }
      @{ Key = [Terminal.Gui.Key]::F7;  Label = "~F7~ Search";       Action = { Show-ADSearchDialog } }
      @{ Key = [Terminal.Gui.Key]::F8;  Label = "~F8~ Focus Tree";   Action = { } }
      @{ Key = [Terminal.Gui.Key]::F9;  Label = "~F9~ Menus";        Action = { } }
      @{ Key = [Terminal.Gui.Key]::F10; Label = "~F10~ Quit";        Action = { [Terminal.Gui.Application]::RequestStop() } }
      @{ Key = [Terminal.Gui.Key]::F11; Label = "~F11~ Full Screen"; Action = { } }
      @{ Key = ([Terminal.Gui.Key]::F12); Label = "~F12~ Context Menu"; Action = {
        $selectedNode = $Script:tree.SelectedObject
        if ($selectedNode -and $selectedNode.Tag -and $selectedNode.Tag.Object) { Show-ObjectContextMenu -Object $selectedNode.Tag.Object -ObjectType $selectedNode.Tag.Type }
      }}
       $handled = $true
    )

    $items = @()
    foreach ($sc in $shortcuts) { $items += [Terminal.Gui.StatusItem]::new($sc.Key, $sc.Label, $sc.Action) }
    $items += $Script:StatusItem
    $Script:StatusBar = [Terminal.Gui.StatusBar]::new($items)
    if ($ThemeData -and $ThemeData.StatusBar) {
      $Script:StatusBar.ColorScheme = $ThemeData.StatusBar
      Debug-Log "Status bar theme applied" -Type "Success"
    }
    Debug-Log "Status bar initialised" -Type "Success"
    return $Script:StatusBar
  }

  ## ----------{ Update mode }---------
  if (-not $Script:StatusItem -or -not $Script:StatusBar) {
    Debug-Log "StatusBar not initialised" -Type "Problem"
    return
  }
  ## Backwards compatibility
  if ($Spinner -and -not $Icon) { $Icon = 'Working' }
  if ($Final   -and -not $Icon) { $Icon = 'Success' }
  ## Prefix
  if ($Icon -and $Script:StatusIcons.ContainsKey($Icon)) {
    $prefix = $Script:StatusIcons[$Icon]
  }
  else {
    $prefix = $Script:StatusIcons['Default']
  }

  ## ----------{ ASCII Progress bar }---------
  $progressText = ""
  if ($PSBoundParameters.ContainsKey('Percent')) {
    if ($Percent -lt 0)   { $Percent = 0 }
    if ($Percent -gt 100) { $Percent = 100 }
    $filled = [Math]::Floor($BarWidth * ($Percent / 100))
    $empty  = $BarWidth - $filled
    $progressText = " [" + ('o' * $filled) + ('-' * $empty) + "] $Percent%"
  }

  ## Build final display
  $displayText = "$prefix | $Message$progressText"
  $Script:StatusItem.Title = $displayText

  ## Refresh
  try {
    $Script:StatusBar.SetNeedsDisplay()
  } catch {
    ## Ignore refresh errors
  }
}

##  We show so many pop-up dialogs, use this function to reduce clutter
function Show-Modal {
  param(
    [string]$title,
    [string]$msg,
    [switch]$YesNo,
    [switch]$EasterEgg
  )
  if ($YesNo) {
    ## Returns 0 for Yes, 1 for No
    ## Terminal.Gui 1.16: Query(int width, int height, string title, string message, params string[] buttons)
    $result = [Terminal.Gui.MessageBox]::Query(60, 9, $title, $msg, @("Yes", "No"))
    return $result
  }
  elseif ($EasterEgg) {
    ## Custom dialog with Easter egg support
    $dlg = [Terminal.Gui.Dialog]::new($title, 60, 12)
    $label = [Terminal.Gui.Label]::new(1, 1, $msg)
    $dlg.Add($label)
    ## Easter egg label (hidden)
    $eggMsg = "You get used to it, I don't even see the code, all I see`nis blond, brunette, redhead...`n`nJeg har det som blommen i et æg!"
    $eggLabel = [Terminal.Gui.Label]::new(1, 1, $eggMsg)
    $eggLabel.Visible = $false
    $dlg.Add($eggLabel)
    ## Add OK button
    $okBtn = [Terminal.Gui.Button]::new("OK")
    $okBtn.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }.GetNewClosure())
    $dlg.AddButton($okBtn)
    ## Key handler for ø
    $dlg.add_KeyPress({
      param($e)
      if ([char]$e.KeyEvent.Key -eq 'ø') {
        $eggLabel.Visible = $true
        $label.Visible    = $false
        $e.Handled        = $true
      }
    }.GetNewClosure())
    [Terminal.Gui.Application]::Run($dlg)
  }
  else {
    [Terminal.Gui.MessageBox]::Query($title, $msg, @("OK")) | Out-Null
  }
}

## ----------{ Initialise Data Source - SIMPLIFIED }----------
function Initialise-DataSource {
  param(
    [string]$FilePath = $null,
    [string]$Domain   = $null
  )

  Debug-Log "========== Initialise-DataSource ==========" -Type "Insight"
  Debug-Log "Parameters - FilePath: '$FilePath', Domain: '$Domain'" -Type "Tracing"
  Debug-Log "Flags - DemoMode: $Script:DemoMode, DataFileLoaded: $Script:DataFileLoaded, ImportDemoData: $Script:ImportDemoData" -Type "Tracing"

  ## If data was already loaded, don't reload
  if ($Script:DataFileLoaded) {
    Debug-Log "Data already loaded from file, skipping re-initialization" -Type "Insight"
    return $true
  }

  ## Priority 1: Active Directory (if available and not in DemoMode)
  if ($Script:HasActiveDirectory -and -not $Script:DemoMode) {
    Debug-Log "Loading from Active Directory..." -Type "Insight"
    if (Load-ADData -Domain $Domain) {
      Refresh-Data -domain $Script:CurrentDomain -RebuildTree
      Debug-Log "Active Directory data loaded successfully" -Type "Success"
      return $true
    }
  }

  ## Priority 2: CSV/JSONC/TDF file (if path provided OR ImportDemoData flag set)
  if ($FilePath -and (Test-Path $FilePath)) {
    Debug-Log "Loading from file: $FilePath" -Type "Insight"
    try {
      $importSuccess = Import-DataFile -FilePath $FilePath
      if ($importSuccess) {
        ## Set flags to prevent overwriting
        $Script:DataFileLoaded = $true
        $Script:DataFilePath   = $FilePath
        $Script:DataSource     = "File"
        Debug-Log "File data loaded successfully, DataFileLoaded flag set" -Type "Success"
        ## Set CurrentDC if we have DCs
        if ($Script:DCs.Count -gt 0 -and -not $Script:CurrentDC) {
          $Script:CurrentDC = $Script:DCs | Where-Object { $_.IsGlobalCatalog } | Select-Object -First 1
          if (-not $Script:CurrentDC) { $Script:CurrentDC = $Script:DCs | Select-Object -First 1 }
          $Script:CurrentDCName = $Script:CurrentDC.Name
          Debug-Log "Set current DC to: $($Script:CurrentDCName)" -Type "Insight"
        }
        return $true
      } else {
        Debug-Log "File import returned false" -Type "Problem"
      }
    } catch {
      Debug-Log "File import failed: $($_.Exception.Message)" -Type "Problem"
    }
  } elseif ($FilePath) {
    Debug-Log "File path provided but does not exist: $FilePath" -Type "Warning"
  }

  ## Priority 3: Demo data (ONLY if DemoMode enabled AND no file was loaded)
  if ($Script:DemoMode -and -not $Script:DataFileLoaded) {
    Debug-Log "DemoMode enabled and no file loaded - loading default demo data..." -Type "Insight"
    $demoSuccess = Load-DefaultDemoData
    if ($demoSuccess) {
      Debug-Log "Default demo data loaded successfully" -Type "Success"
      ## Set CurrentDC if we have DCs
      if ($Script:DCs.Count -gt 0 -and -not $Script:CurrentDC) {
        $Script:CurrentDC = $Script:DCs | Where-Object { $_.IsGlobalCatalog } | Select-Object -First 1
        if (-not $Script:CurrentDC) { $Script:CurrentDC = $Script:DCs | Select-Object -First 1 }
        $Script:CurrentDCName = $Script:CurrentDC.Name
        Debug-Log "Set current DC to: $($Script:CurrentDCName)" -Type "Insight"
      }
      return $true
    }
  }

  ## Set CurrentDC if we have DCs but no CurrentDC set yet
  if ($Script:DCs.Count -gt 0) {
    if (-not $Script:CurrentDC) {
      ## Prefer Global Catalog DC
      $Script:CurrentDC = $Script:DCs | Where-Object { $_.IsGlobalCatalog } | Select-Object -First 1
      if (-not $Script:CurrentDC) { $Script:CurrentDC = $Script:DCs | Select-Object -First 1 }
      $Script:CurrentDCName = $Script:CurrentDC.Name
      Debug-Log "Set current DC to: $($Script:CurrentDCName)" -Type "Insight"
    } else {
      $Script:CurrentDCName = $Script:CurrentDC.Name
      Debug-Log "CurrentDC already set to: $($Script:CurrentDCName), preserving selection" -Type "Tracing"
    }
  }

  ## Normalize display name
  $Script:CurrentDCName = if ($Script:CurrentDC -is [hashtable]) { $Script:CurrentDC['Name'] }
                      elseif ($Script:CurrentDC -is [string]) { $Script:CurrentDC }
                      elseif ($Script:CurrentDC) { $Script:CurrentDC.Name }
                      else { "(None)" }

  ## Nothing worked
  Debug-Log "No data source available!" -Type "Problem"
  return $false
}

## ----------{ Dual-pane list dialog helper }----------
function New-DualPaneListDialog {
  <#
  .SYNOPSIS
  Creates a dialog with two list panes and a details view

  .PARAMETER Title
  Dialog title

  .PARAMETER LeftPane
  Hashtable: @{ Title=""; Items=@(); FormatItem={} }

  .PARAMETER RightPane
  Hashtable: @{ Title=""; Items=@(); FormatItem={}; OnFilter={} }

  .PARAMETER DetailsPane
  Hashtable: @{ Title=""; OnSelectionChanged={} }

  .PARAMETER Buttons
  Array of button definitions: @{ Text=""; OnClick={} }

  .PARAMETER Width
  Dialog width

  .PARAMETER Height
  Dialog height
  #>

  param(
    [Parameter(Mandatory)]
    [string]$Title,
    [Parameter(Mandatory)]
    [hashtable]$LeftPane,
    [Parameter(Mandatory)]
    [hashtable]$RightPane,
    [hashtable]$DetailsPane,
    [array]$Buttons  = @(),
    [string]$Summary = "",
    [int]$Width      = 120,
    [int]$Height     = 40
  )

  $dialog        = [Terminal.Gui.Dialog]::new()
  $dialog.Title  = $Title
  $dialog.Width  = $Width
  $dialog.Height = $Height
  $y = 1

  ## Summary
  if ($Summary) {
    $lblSummary    = [Terminal.Gui.Label]::new($Summary)
    $lblSummary.X = 2; $lblSummary.Y = $y
    $dialog.Add($lblSummary)
    $y += 2
  }

  ## Left pane
  $lblLeft = [Terminal.Gui.Label]::new($LeftPane.Title)
  $lblLeft.X = 2; $lblLeft.Y = $y
  $dialog.Add($lblLeft)

  ## Right pane label
  $lblRight = [Terminal.Gui.Label]::new($RightPane.Title)
  $lblRight.X = 44; $lblRight.Y = $y
  $dialog.Add($lblRight)
  $y += 1

  ## Left list
  $lstLeft = [Terminal.Gui.ListView]::new()
  $lstLeft.X = 2; $lstLeft.Y = $y
  $lstLeft.Width = 40; $lstLeft.Height = 8

  $leftItems = if ($LeftPane.Items.Count -gt 0) {
    $LeftPane.Items | ForEach-Object { & $LeftPane.FormatItem $_ }
  } else {
    @("(No items)")
  }
  $lstLeft.SetSource($leftItems)
  $dialog.Add($lstLeft)

  ## Right list
  $lstRight = [Terminal.Gui.ListView]::new()
  $lstRight.X = 44; $lstRight.Y = $y
  $lstRight.Width = [Terminal.Gui.Dim]::Fill(2); $lstRight.Height = 8

  $rightItems = if ($RightPane.Items.Count -gt 0) {
    $RightPane.Items | ForEach-Object { & $RightPane.FormatItem $_ }
  } else {
    @("(No items)")
  }
  $lstRight.SetSource($rightItems)
  $dialog.Add($lstRight)
  $y += 9

  ## Details pane (if provided)
  if ($DetailsPane) {
    $lblDetails = [Terminal.Gui.Label]::new($DetailsPane.Title)
    $lblDetails.X = 2; $lblDetails.Y = $y
    $dialog.Add($lblDetails)
    $y += 1

    $txtDetails = [Terminal.Gui.TextView]::new()
    $txtDetails.X = 2; $txtDetails.Y = $y
    $txtDetails.Width = [Terminal.Gui.Dim]::Fill(2)
    $txtDetails.Height = 10
    $txtDetails.ReadOnly = $true
    $txtDetails.WordWrap = $false
    $dialog.Add($txtDetails)
    $y += 11

    ## Hook up selection changed
    if ($DetailsPane.OnSelectionChanged) {
      $lstRight.add_SelectedItemChanged({
        $selectedIndex = $lstRight.SelectedItem
        $selectedText = $lstRight.Source.ToList()[$selectedIndex]
        if ($selectedText -eq "(No items)" -or $selectedText -eq "(No matches)") {
          $txtDetails.Text = "(No item selected)"
          return
        }

        ## Find actual item
        $selectedItem = $null
        for ($i = 0; $i -lt $RightPane.Items.Count; $i++) {
          $formatted = & $RightPane.FormatItem $RightPane.Items[$i]
          if ($formatted -eq $selectedText) {
            $selectedItem = $RightPane.Items[$i]
            break
          }
        }
        if ($selectedItem) {
          $detailsText = & $DetailsPane.OnSelectionChanged $selectedItem
          $txtDetails.Text = $detailsText
        }
      }.GetNewClosure())
    }
  }

  ## Buttons
  $btnX = 2
  foreach ($btn in $Buttons) {
    $button = [Terminal.Gui.Button]::new($btn.Text)
    $button.X = $btnX; $button.Y = $y
    if ($btn.OnClick) {
      $button.add_Clicked({
        & $btn.OnClick @{
          LeftList   = $lstLeft
          LeftItems  = $LeftPane.Items
          RightList  = $lstRight
          RightItems = $RightPane.Items
          Dialog     = $dialog
        }
      }.GetNewClosure())
    }
    $dialog.Add($button)
    $btnX += $btn.Text.Length + 7
  }

  ## Close button
  $btnClose   = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = $btnX; $btnClose.Y = $y
  $btnClose.add_Clicked({ $dialog.RequestStop() }).GetNewClosure()
  $dialog.Add($btnClose)
  [Terminal.Gui.Application]::Run($dialog)
}

## ----------{ Get Theme }----------
function Get-Theme {
  <#
  .SYNOPSIS
  Get or dump theme colour schemes

  .PARAMETER Mode
  Theme name to load

  .PARAMETER Dump
  If specified, dumps the current theme colours to Debug-Log instead of loading

  .EXAMPLE
  Get-Theme -Mode "matrix"

  .EXAMPLE
  Get-Theme -Dump

  .EXAMPLE
  Get-Theme -Mode "british" -Dump
  #>

  param(
    [string]$Mode,
    [switch]$Dump
  )

  ## Dump mode
  if ($Dump) {
    $themeName = if ($Mode) { $Mode } else { $Script:ThemeMode }
    Debug-Log "Dumping colour scheme for theme: $themeName" -Type "Insight"
    ## Dump the ColorSchemes that are actually being used
    Debug-Log "=== Script ColorScheme ===" -Type "Insight"
      if ($Script:ScriptCs) {
        Debug-Log "Normal    : $($Script:ScriptCs.Normal)" -Type "Insight"
        Debug-Log "Focus     : $($Script:ScriptCs.Focus)" -Type "Insight"
        Debug-Log "HotNormal : $($Script:ScriptCs.HotNormal)" -Type "Insight"
        Debug-Log "HotFocus  : $($Script:ScriptCs.HotFocus)" -Type "Insight"
        Debug-Log "Disabled  : $($Script:ScriptCs.Disabled)" -Type "Insight"
      } else {
        Debug-Log "ScriptCs is null!" -Type "Warning"
      }

      Debug-Log "=== Main Window ColorScheme ===" -Type "Insight"
      if ($Script:mainWindowCs) {
        Debug-Log "Normal    : $($Script:mainWindowCs.Normal)" -Type "Insight"
        Debug-Log "Focus     : $($Script:mainWindowCs.Focus)" -Type "Insight"
        Debug-Log "HotNormal : $($Script:mainWindowCs.HotNormal)" -Type "Insight"
        Debug-Log "HotFocus  : $($Script:mainWindowCs.HotFocus)" -Type "Insight"
        Debug-Log "Disabled  : $($Script:mainWindowCs.Disabled)" -Type "Insight"
      } else {
        Debug-Log "mainWindowCs is null!" -Type "Warning"
      }
      return
    }

    ## Load Mode
    if (-not $Mode) { throw "Get-Theme called with empty mode" }
    ## Initialise colour schemes and Ensure ColorSchemes are instantiated
    if (-not $Script:ScriptCs) { $Script:ScriptCs = [Terminal.Gui.ColorScheme]::new() }
    if (-not $Script:mainWindowCs) { $Script:mainWindowCs = [Terminal.Gui.ColorScheme]::new() }
    ## Normalize theme string: lowercase + ASCII
    $Mode = $Mode.Trim().ToLower()

    <# Documentation for adding themes:

    Adding Themes:

    Add an option in the switch statement below and define:

    "faxekondi" {
        $Script:ScriptCs.Normal     <-- Foreground borders and background colour for all modals
        $Script:ScriptCs.Focus      <-- Foreground and background for menus
        $Script:mainWindowCs.Normal <-- Main opening dialog and foreground text colour
        $Script:mainWindowCs.Focus  <-- Main opening window focus colours foreground and background
    }

    Valid colours: Black, Blue, Green, Cyan, Red, Magenta, Brown, Gray, DarkGray, BrightBlue,
                   BrightGreen, BrightCyan, BrightRed, BrightMagenta, BrightYellow, White
    #>

    switch ($Mode) {

      "british" {
        $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
        $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
        $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
        $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::White)
      }
      "class91" {
        $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Black)
        $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::White)
        $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightRed,[Terminal.Gui.Color]::White)
        $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightBlue)
      }
      "dark" {
        $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Black)
        $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Green,[Terminal.Gui.Color]::White)
        $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Green)
        $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Red)
      }
      ## Deutsche Bundesbahn – Orientrot era (mid/late 80s)
      "db-1980s" {
        $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Gray)
        $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
        $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Gray)
        $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Red)
      }
      "default" {
        # fallback to dark
        $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Black)
        $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Gray)
        $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Black)
        $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::DarkGray)
      }
      "dsb" {
        $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
        $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
        $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
        $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::White)
      }
      "gemstones" {
        $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Green,[Terminal.Gui.Color]::White)
        $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::BrightMagenta)
        $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightGreen)
        $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Green,[Terminal.Gui.Color]::White)
      }
      ## BR Executive / Swallow
      "intercity-swallow" {
        $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::DarkGray)
        $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
        $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::DarkGray)
        $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Red)
      }
      ## Had to be done
      "irn-bru" {
        $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightRed)
        $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Blue,[Terminal.Gui.Color]::BrightYellow)
        $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightRed)
        $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::BrightYellow)
      }
     "light" {
        $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Cyan)
        $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
        $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Cyan)
        $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::Blue)
      }
    "matrix" {
      $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Green,[Terminal.Gui.Color]::Black)
      $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Green)
      $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::Black)
      $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::Gray)
    }
    ## Dutch Railways (Nederlandse Spoorwegen)
    "ns" {
      $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::BrightYellow)
      $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
      $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::BrightYellow)
      $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Yellow,[Terminal.Gui.Color]::Blue)
    }
    ## NSE "Toothpaste" livery
    "network-southeast" {
      $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::White)
      $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
      $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::White)
      $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Red)
    }
    "panam" {
      $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightBlue,[Terminal.Gui.Color]::White)
      $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightBlue)
      $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::BrightBlue)
      $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightBlue,[Terminal.Gui.Color]::White)
    }
    "procomm" {
      $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::Red)
      $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::BrightBlue)
      $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Black)
      $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightBlue,[Terminal.Gui.Color]::White)
    }
    "scotrail" {
      $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::Blue)
      $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Blue,[Terminal.Gui.Color]::White)
      $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::BrightBlue)
      $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
    }
    ## Trans World Airways not two (if you're Scottish)
    "twa" {
      $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
      $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::DarkGray)
      $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Red)
      $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightRed,[Terminal.Gui.Color]::DarkGray)
    }
    "viarail" {
      $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::BrightYellow)
      $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
      $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::BrightYellow)
      $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::Blue)
    }
    "viarail-soft" {
      $Script:ScriptCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
      $Script:ScriptCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Blue,[Terminal.Gui.Color]::BrightYellow)
      $Script:mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
      $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Gray)
    }
    default {
      throw "Unknown theme: $Mode"
    }
  }

  ## Ensure HotNormal/HotFocus
  $Script:ScriptCs.HotNormal     = $Script:ScriptCs.Normal
  $Script:ScriptCs.HotFocus      = $Script:ScriptCs.Focus
  $Script:mainWindowCs.HotNormal = $Script:mainWindowCs.Normal
  $Script:mainWindowCs.HotFocus  = $Script:mainWindowCs.Focus

  return @{
    Global     = $Script:ScriptCs
    MainWindow = $Script:mainWindowCs
  }
}

## ----------{ Apply Colours }----------
function Apply-Theme {
  param(
    [hashtable]$ThemeData,
    [object]$TopLevel,
    [object]$MainWindow,
    [object]$Menu,
    [object]$StatusBar
  )
  if ($null -eq $ThemeData) { return }

  ## Terminal.Gui base colours FIRST
  [Terminal.Gui.Colors]::Base     = $ThemeData.Global
  [Terminal.Gui.Colors]::Dialog   = $ThemeData.Global
  [Terminal.Gui.Colors]::Menu     = $ThemeData.Global
  [Terminal.Gui.Colors]::Error    = $ThemeData.Global
  [Terminal.Gui.Colors]::TopLevel = $ThemeData.Global

  ## Global / TopLevel
  if ($TopLevel -and $TopLevel.PSObject.Properties.Name -contains 'ColorScheme') {
    $TopLevel.ColorScheme = $ThemeData.Global
    $TopLevel.SetNeedsDisplay()
  }
  ## Main window
  if ($MainWindow -and $MainWindow.PSObject.Properties.Name -contains 'ColorScheme') {
    $MainWindow.ColorScheme = $ThemeData.MainWindow
    $MainWindow.SetNeedsDisplay()
  }
  ## Menu
  if ($Menu -and $Menu.PSObject.Properties.Name -contains 'ColorScheme') {
    $Menu.ColorScheme = $ThemeData.Global
    $Menu.SetNeedsDisplay()
  }
  ## StatusBar
  if ($StatusBar -and $StatusBar.PSObject.Properties.Name -contains 'ColorScheme') {
    $StatusBar.ColorScheme = $ThemeData.Global
    $StatusBar.SetNeedsDisplay()
  }
  ## Tree - Use main window scheme to match parent
  if ($Script:tree -and $Script:tree.PSObject.Properties.Name -contains 'ColorScheme') {
    $Script:tree.ColorScheme = $ThemeData.MainWindow  # ← CHANGED FROM Global
    $Script:tree.SetNeedsDisplay()
  }
  ## Filter Panel
  if ($Script:filterPanel -and $Script:filterPanel.PSObject.Properties.Name -contains 'ColorScheme') {
    $Script:filterPanel.ColorScheme = $ThemeData.MainWindow  # ← Match parent
    $Script:filterPanel.SetNeedsDisplay()
  }
  ## Force complete refresh
  [Terminal.Gui.Application]::Refresh()
}

## ----------{ Show progress bar }----------
## Helper: Show a simple loading/progress dialog with spinner
function Show-LoadingDialog {
  param(
    [string]$Message = "Loading, please wait..."
  )

  ## Create the dialog
  $dlg = [Terminal.Gui.Dialog]::new("", 40, 7)
  $dlg.X = 2
  $dlg.Y = 2

  ## Message label
  $lbl = [Terminal.Gui.Label]::new($Message)
  $lbl.X = 2
  $lbl.Y = 2
  $dlg.Add($lbl)

  ## Spinner label
  $spinner = [Terminal.Gui.Label]::new("|")
  $spinner.X = [Terminal.Gui.Pos]::Right($lbl) + 1
  $spinner.Y = 2
  $dlg.Add($spinner)

  ## Spinner globals
  $Script:spinnerFrames = @("|","/","-","\\")
  $Script:spinnerFrameIndex = 0

  ## Add a timeout handler to update the spinner
  $timeout = [Terminal.Gui.Application]::AddTimeout([TimeSpan]::FromMilliseconds(150), {
    $Script:spinnerFrameIndex = ($Script:spinnerFrameIndex + 1) % $Script:spinnerFrames.Count
    $spinner.Text = $Script:spinnerFrames[$Script:spinnerFrameIndex]

    ## Return true to keep the timeout running
    return $true
  })

  ## Show dialog in modal-safe, non-blocking way
  [Terminal.Gui.Application]::BeginModal($dlg)

  ## Return dialog and timeout for clean-up
  return [PSCustomObject]@{
    Dialog  = $dlg
    Timeout = $timeout
  }
}

## ----------{ Danske Soda vand }----------
## This is a theme now. Danish Fruit soda based fun. Method to the madness
function Show-BlaabaerInfo {
  $message = @"
$($Script:ProjectName) is codenamed $Script:FruitName because:
- I was drinking blueberry soda when writing the code
- $($Script:FruitName) is Danish for blueberry
- Føtex sells a rather nice $($Script:FruitName) soda
- Every great project needs a forest-fruit mascot!
"@
  Show-Modal "Why $($Script:FruitName)...? 🫐" $message -EasterEgg
}

## ~~~~~~~~~~~~~~~~~~~~~~~~~{ AD FUNCTIONS BELOW HERE }~~~~~~~~~~~~~~~~~~~~~~~~~

## ----------{ Cleaned up DNS Dialog }----------
function Show-DNSDialog {
  param([string]$Domain)

  if (-not $Domain) { $Domain = $Script:CurrentDomain }
  Debug-Log "Opening DNS viewer for domain: $Domain" -Type "Insight"

  ## Get DNS data
  $dnsZones = @()
  $dnsRecords = @()

  if ($Script:rawDNSZones)   { $dnsZones = $Script:rawDNSZones | Where-Object { $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')") }}
  if ($Script:rawDNSRecords) { $dnsRecords = $Script:rawDNSRecords | Where-Object { $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')") }}
  Debug-Log "Found $($dnsZones.Count) zones and $($dnsRecords.Count) records" -Type "Insight"

  ## Format functions
  $formatZone = { param($zone) $zone.Name }
  $formatRecord = {
    param($record)
    $recordType = "Unknown"
    if ($record.Name -match '^_') { $recordType = "SRV" }
    elseif ($record.Name -match '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}') { $recordType = "PTR" }
    else { $recordType = "A/Other" }
    "$($record.Name) [$recordType]"
  }

  ## Details handler
  $showRecordDetails = {
    param($record)
    @"
DNS Record Details
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Name: $($record.Name)

Distinguished Name:
$($record.DN)

DNS Record Data:
$(if ($record.DNSRecord) { $record.DNSRecord } else { "(No data)" })

Tombstoned: $(if ($record.DNSTombstoned) { $record.DNSTombstoned } else { "False" })

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Note: This is raw DNS record data from AD.
Use DNS Manager for detailed record information.
"@
  }

  ## Button definitions
  $buttons = @(
    @{
      Text = "View Zone Details..."
      OnClick = {
        param($state)
        $selectedIndex = $state.LeftList.SelectedItem
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $dnsZones.Count) {
          $zone = $dnsZones[$selectedIndex]
          Show-DNSZoneDetailsDialog -Zone $zone
        } else {
          Show-Modal "Info" "Please select a DNS zone to view"
        }
      }
    },
    @{
      Text = "Export Records..."
      OnClick = {
        param($state)
        try {
          $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
          $filename   = "dns_records_${Domain}_$timestamp.csv"
          $exportData = $dnsRecords | ForEach-Object {
            [PSCustomObject]@{
              Name = $_.Name
              RecordData = $_.DNSRecord
              Tombstoned = $_.DNSTombstoned
              DN = $_.DN
            }
          }
          $exportData | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
          Show-Modal "Success" "Exported $($dnsRecords.Count) DNS records to:`n$filename"
          Debug-Log "Exported DNS records to $filename" -Type "Success"
        } catch {
          Show-Modal "Error" "Failed to export DNS records:`n$($_.Exception.Message)"
          Debug-Log "Export failed: $($_.Exception.Message)" -Type "Problem"
        }
      }
    },
    @{
      Text = "Open DNS Manager"
      OnClick = {
        param($state)
        try {
          if ($IsWindows) {
            Debug-Log "Launching dnsmgmt.msc" -Type "Insight"
            Start-Process "dnsmgmt.msc" -ErrorAction Stop
            Show-Modal "DNS Manager" "Launching DNS Manager (dnsmgmt.msc)...`n`nNote: Requires administrative privileges and DNS tools installed."
          } else {
            Show-Modal "Not Available" "DNS Manager (dnsmgmt.msc) is only available on Windows.`n`nOn Linux/macOS, use 'nsupdate' or web-based DNS management tools."
          }
        } catch {
          Show-Modal "Error" "Failed to launch DNS Manager:`n$($_.Exception.Message)`n`nEnsure DNS management tools are installed."
          Debug-Log "Failed to launch DNS Manager: $($_.Exception.Message)" -Type "Problem"
        }
      }
    }
  )
  ## Show dialog
  New-DualPaneListDialog -Title "DNS Records - $Domain" -Summary "DNS Zones: $($dnsZones.Count)  |  DNS Records: $($dnsRecords.Count)" -LeftPane @{ Title = "DNS Zones:"; Items = $dnsZones; FormatItem = $formatZone } -RightPane @{ Title = "DNS Records:"; Items = $dnsRecords; FormatItem = $formatRecord } -DetailsPane @{ Title = "Record Details:"; OnSelectionChanged = $showRecordDetails } -Buttons $buttons -Width 120 -Height 40
}

## ----------{ DNS Zone Details Dialog }----------
## Part of the DNS manager dialog
function Show-DNSZoneDetailsDialog {
  param($Zone)

  Debug-Log "Showing details for DNS zone: $($Zone.Name)" -Type "Insight"

  $details = @"
DNS Zone Details
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Zone Name: $($Zone.Name)

Distinguished Name:
$($Zone.DN)

Description:
$(if ($Zone.Description) { $Zone.Description } else { "(No description)" })

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

This is an Active Directory-integrated DNS zone.

For detailed zone configuration, use DNS Manager
or PowerShell DNS cmdlets (Get-DnsServerZone).
"@

  $dialog = [Terminal.Gui.Dialog]::new()
  $dialog.Title = "DNS Zone - $($Zone.Name)"
  $dialog.Width = 80
  $dialog.Height = 20
  $txtDetails = [Terminal.Gui.TextView]::new()
  $txtDetails.X = 1
  $txtDetails.Y = 1
  $txtDetails.Width = [Terminal.Gui.Dim]::Fill(1)
  $txtDetails.Height = [Terminal.Gui.Dim]::Fill(3)
  $txtDetails.ReadOnly = $true
  $txtDetails.Text = $details
  $dialog.Add($txtDetails)
  $btnClose   = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = [Terminal.Gui.Pos]::Center()
  $btnClose.Y = [Terminal.Gui.Pos]::Bottom($txtDetails) + 1
  $btnClose.add_Clicked({ $dialog.RequestStop() }).GetNewClosure()
  $dialog.Add($btnClose)
  [Terminal.Gui.Application]::Run($dialog)
}

function Show-IPSecPoliciesDialog {
  param([string]$Domain)
  if (-not $Domain) { $Domain = $Script:CurrentDomain }
  Debug-Log "Opening IPSec policies viewer for domain: $Domain" -Type "Insight"
  ## Get IPSec data
  $ipsecPolicies = @()
  if ($Script:rawIPSecPolicies) {
    $ipsecPolicies = $Script:rawIPSecPolicies | Where-Object { $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')") }
    if ($ipsecPolicies.Count -eq 0) { $ipsecPolicies = $Script:rawIPSecPolicies }
  }
  Debug-Log "Found $($ipsecPolicies.Count) IPSec object(s)" -Type "Insight"
  if ($ipsecPolicies.Count -eq 0) {
    Show-Modal "No IPSec Policies" "No IPSec policies found for domain: $Domain`n`nThis domain may not have any IPSec policies configured in Active Directory."
    return
  }

  ## Format function
  $formatPolicy = {
    param($policy)
    $name = if ($policy.IPSecName) { $policy.IPSecName } elseif ($policy.Name) { $policy.Name } else { "(Unnamed)" }
    $type = $policy.Type
    "$name [$type]"
  }

  ## View handler
  $onView = {
    param($policy)

    ## Decode IPSec object type
    $typeDesc = switch ($policy.Type) {
      "ipsecPolicy"             { "IPSec Policy (main policy object)" }
      "ipsecNFA"                { "IPSec Negotiation Policy Association (links filters to policies)" }
      "ipsecISAKMPPolicy"       { "ISAKMP Policy (Phase 1 - IKE negotiation)" }
      "ipsecNegotiationPolicy"  { "IPSec Negotiation Policy (Phase 2 - IPSec SA)" }
      "ipsecFilter"             { "IPSec Filter (traffic matching rules)" }
      default                   { $policy.Type }
    }

    $details = @"
IPSec Policy Details
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Name: $(if ($policy.IPSecName) { $policy.IPSecName } else { $policy.Name })

Type: $($policy.Type)
$typeDesc

IPSec ID: $(if ($policy.IPSecID) { $policy.IPSecID } else { "N/A" })

Description:
$(if ($policy.Description) { $policy.Description } else { "(No description)" })

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Distinguished Name:
$($policy.DN)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

About IPSec Policies:

IPSec policies in Active Directory define how network
traffic is secured using Internet Protocol Security.

Policy Components:
- ipsecPolicy: Main policy object containing rules
- ipsecNFA: Links filters to negotiation policies
- ipsecISAKMPPolicy: IKE (Phase 1) key exchange
- ipsecNegotiationPolicy: IPSec SA (Phase 2) settings
- ipsecFilter: Traffic matching criteria

To manage IPSec policies, use:
- IP Security Policy Management (secpol.msc)
- Group Policy Management Console
- Windows Firewall with Advanced Security
- PowerShell cmdlets (Get-NetIPsecRule, etc.)
"@

    Show-Modal "IPSec Policy Details" $details
  }

  ## Export handler
  $onExport = {
    param($items)
    try {
      $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
      $filename   = "ipsec_policies_${Domain}_$timestamp.csv"
      $exportData = $items | ForEach-Object {
        [PSCustomObject]@{
          Name        = if ($_.IPSecName) { $_.IPSecName } else { $_.Name }
          Type        = $_.Type
          IPSecID     = $_.IPSecID
          Description = $_.Description
          DN          = $_.DN
        }
      }

      $exportData | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
      Show-Modal "Success" "Exported $($items.Count) IPSec object(s) to:`n$filename"
      Debug-Log "Exported IPSec policies to $filename" -Type "Success"
    } catch {
      Show-Modal "Error" "Failed to export IPSec policies:`n$($_.Exception.Message)"
      Debug-Log "Export failed: $($_.Exception.Message)" -Type "Problem"
    }
  }

  ## Show dialog using existing helper
  New-ListDialog -Title "IPSec Policies - $Domain" -Items $ipsecPolicies -FormatItem $formatPolicy -OnView $onView -OnExport $onExport -FilterHelp "(Filter by policy name or type)"
}

function Show-IPSecHelpDialog {
  $helpText = @"
IPSec Policies in Active Directory
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Overview:
IPSec (Internet Protocol Security) policies define how
network traffic is authenticated and encrypted between
computers in an Active Directory domain.

Policy Architecture:

1. ipsecPolicy
   The main policy object containing one or more rules.
   Example: "Secure Server Policy"

2. ipsecNFA (Negotiation Policy Association)
   Links traffic filters to negotiation settings.
   Defines WHAT traffic to secure and HOW to secure it.

3. ipsecISAKMPPolicy (Phase 1 - IKE)
   Internet Key Exchange settings for establishing
   secure channels. Includes encryption algorithms
   (DES, 3DES, AES) and authentication methods.

4. ipsecNegotiationPolicy (Phase 2)
   IPSec Security Association settings for actual
   data protection. Defines AH/ESP protocols.

5. ipsecFilter
   Traffic matching rules (source/dest IP, ports,
   protocols). Determines which packets are secured.

Common Use Cases:
• Server isolation (only allow authenticated clients)
• Domain isolation (encrypt all domain traffic)
• Require encryption for sensitive servers
• Block specific traffic patterns

Management Tools:

Windows:
• Local Security Policy (secpol.msc)
• Group Policy Management (gpmc.msc)
• IP Security Monitor (ipsecmon.exe)
• PowerShell: Get-NetIPsecRule, New-NetIPsecRule

Command Line:
• netsh ipsec show all
• netsh ipsec static show policy all

Linux:
• ipsec status (strongSwan/libreswan)
• ip xfrm state/policy

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

For detailed IPSec configuration and troubleshooting,
refer to Microsoft IPSec documentation.
"@

  $helpDialog        = [Terminal.Gui.Dialog]::new()
  $helpDialog.Title  = "IPSec Policies - Help"
  $helpDialog.Width  = 90
  $helpDialog.Height = 35
  $txtHelp           = [Terminal.Gui.TextView]::new()
  $txtHelp.X         = 1
  $txtHelp.Y         = 1
  $txtHelp.Width     =  [Terminal.Gui.Dim]::Fill(1)
  $txtHelp.Height    = [Terminal.Gui.Dim]::Fill(3)
  $txtHelp.ReadOnly  = $true
  $txtHelp.Text      = $helpText
  $helpDialog.Add($txtHelp)
  $btnClose   = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = [Terminal.Gui.Pos]::Center()
  $btnClose.Y = [Terminal.Gui.Pos]::AnchorEnd(2)
  $btnClose.add_Clicked({ $helpDialog.RequestStop() }).GetNewClosure()
  $helpDialog.Add($btnClose)
  [Terminal.Gui.Application]::Run($helpDialog)
}

function Show-PrintQueuesDialog {
  param([string]$Domain)

  if (-not $Domain) { $Domain = $Script:CurrentDomain }
  Debug-Log "Opening Print Queues viewer for domain: $Domain" -Type "Insight"

  ## Get print queue data
  $printQueues = @()

  if ($Script:rawPrintQueues) {
    $printQueues = $Script:rawPrintQueues | Where-Object { $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')") }
    if ($printQueues.Count -eq 0) { $printQueues = $Script:rawPrintQueues }
  }
  Debug-Log "Found $($printQueues.Count) print queue(s)" -Type "Insight"
  if ($printQueues.Count -eq 0) {
    Show-Modal "No Print Queues" "No shared print queues found for domain: $Domain`n`nThis domain may not have any printers published in Active Directory."
    return
  }

  ## Group by server
  $serverGroups = $printQueues | Where-Object { $_.ServerName } | Group-Object ServerName
  $servers = @()
  if ($serverGroups) {
    $servers = $serverGroups | ForEach-Object {
      [PSCustomObject]@{
        ServerName   = $_.Name
        PrinterCount = $_.Count
      }
    }
  }

  ## Format functions
  $formatServer = {
    param($server)
    "$($server.ServerName) ($($server.PrinterCount) printers)"
  }

  $formatPrinter = {
    param($printer)
    $name     = $printer.PrinterName ?? $printer.Name ?? "(Unnamed)"
    $server   = if ($printer.ServerName) { " on $($printer.ServerName)" } else { "" }
    $location = if ($printer.Location) { " [$($printer.Location)]" } else { "" }
    "$name$server$location"
  }

  ## Details handler
  $showPrinterDetails = {
    param($printer)

    $uncPath = if ($printer.ServerName -and $printer.ShareName) {
      "\\$($printer.ServerName)\$($printer.ShareName)"
    } elseif ($printer.PrinterName) {
      $printer.PrinterName
    } else {
      "N/A"
    }

    @"
Print Queue Details
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Printer Name:       $($printer.PrinterName ?? $printer.Name)
UNC Path:           $uncPath
Server:             $(if ($printer.ServerName) { $printer.ServerName } else { "N/A" })
Share Name:         $(if ($printer.ShareName) { $printer.ShareName } else { "N/A" })
Location:           $(if ($printer.Location) { $printer.Location } else { "(Not specified)" })

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Distinguished Name: $($printer.DN)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

About Print Queues in Active Directory:

Print queues published in AD allow users to easily find and connect to shared printers across the domain.

Benefits:
• Users can search for printers by location
• Centralized printer management
• Automatic driver deployment via Group Policy
• Follow-me printing across locations

To Connect:
Windows: $uncPath
Add via: Devices and Printers → Add Printer → Network

To Manage:
• Print Management Console (printmanagement.msc)
• Server Manager → Print Services
• PowerShell: Get-Printer, Add-Printer
"@
  }

  ## Button definitions
  $buttons = @(
    @{
      Text = "Copy UNC Path"
      OnClick = {
        param($state)
        $selectedIndex = $state.RightList.SelectedItem
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $printQueues.Count) {
          $printer = $printQueues[$selectedIndex]
          if ($printer.ServerName -and $printer.ShareName) {
            $uncPath = "\\$($printer.ServerName)\$($printer.ShareName)"
            try {
              if ($IsWindows) {
                Set-Clipboard -Value $uncPath
              } elseif ($IsLinux) {
                $uncPath | xclip -selection clipboard
              } elseif ($IsMacOS) {
                $uncPath | pbcopy
              }
              Show-Modal "Success" "Copied to clipboard:`n$uncPath`n`nTo connect:`nControl Panel → Devices and Printers → Add Printer`nOr run: rundll32 printui.dll,PrintUIEntry /in /n $uncPath"
            } catch {
              Show-Modal "Error" "Failed to copy to clipboard:`n$($_.Exception.Message)"
            }
          } else {
            Show-Modal "Error" "Unable to determine UNC path for this printer"
          }
        } else {
          Show-Modal "Info" "Please select a printer"
        }
      }
    },
    @{
      Text = "Export List..."
      OnClick = {
        param($state)
        try {
          $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
          $filename   = "print_queues_${Domain}_$timestamp.csv"
          $exportData = $printQueues | ForEach-Object {
            [PSCustomObject]@{
              PrinterName = $_.PrinterName ?? $_.Name
              ServerName  = $_.ServerName
              ShareName   = $_.ShareName
              UNCPath     = if ($_.ServerName -and $_.ShareName) { "\\$($_.ServerName)\$($_.ShareName)" } else { "" }
              Location    = $_.Location
              DN          = $_.DN
            }
          }
          $exportData | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
          Show-Modal "Success" "Exported $($printQueues.Count) print queue(s) to:`n$filename"
          Debug-Log "Exported print queues to $filename" -Type "Success"
        } catch {
          Show-Modal "Error" "Failed to export print queues:`n$($_.Exception.Message)"
          Debug-Log "Export failed: $($_.Exception.Message)" -Type "Problem"
        }
      }
    },
    @{
      Text = "Open Print Management"
      OnClick = {
        param($state)
        try {
          if ($IsWindows) {
            Debug-Log "Launching printmanagement.msc" -Type "Insight"
            Start-Process "printmanagement.msc" -ErrorAction Stop
            Show-Modal "Print Management" "Launching Print Management Console...`n`nNote: Requires Print Services management tools installed."
          } else {
            Show-Modal "Not Available" "Print Management Console (printmanagement.msc) is only available on Windows.`n`nOn Linux, use CUPS web interface (http://localhost:631) or cupsd commands."
          }
        } catch {
          Show-Modal "Error" "Failed to launch Print Management:`n$($_.Exception.Message)`n`nInstall Print Services management tools via Server Manager."
          Debug-Log "Failed to launch printmanagement.msc: $($_.Exception.Message)" -Type "Problem"
        }
      }
    }
  )
  ## Show dialog using existing helper
  New-DualPaneListDialog -Title "Shared Print Queues - $Domain" -Summary "Shared Printers: $($printQueues.Count)  |  Print Servers: $($serverGroups.Count)" -LeftPane @{ Title = "Print Servers:"; Items = $servers; FormatItem = $formatServer } -RightPane @{ Title = "Print Queues:"; Items = $printQueues; FormatItem = $formatPrinter } -DetailsPane @{ Title = "Printer Details:"; OnSelectionChanged = $showPrinterDetails } -Buttons $buttons -Width 120 -Height 40
}

## show Domain trusts window
function Show-TrustsDialog {
  param([string]$Domain)

  if (-not $Domain) { $Domain = $Script:CurrentDomain }
  Debug-Log "Opening Trusts viewer for domain: $Domain" -Type "Insight"
  ## Get trust data
  $trusts = @()
  if ($Script:rawTrusts) {
    $trusts = $Script:rawTrusts | Where-Object { $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')") }
    if ($trusts.Count -eq 0) { $trusts = $Script:rawTrusts }
  }
  Debug-Log "Found $($trusts.Count) trust(s)" -Type "Insight"
  if ($trusts.Count -eq 0) {
    Show-Modal "No Trusts" "No trust relationships found for domain: $Domain`n`nThis domain may not have any external trust relationships configured."
    return
  }
  ## Format function
  $formatTrust = {
    param($trust)
    $partner   = if ($trust.Partner) { $trust.Partner } else { $trust.Name }
    $trustType = if ($trust.Type) { $trust.Type } else { "Unknown" }
    $direction = if ($trust.Direction) { $trust.Direction } else { "Unknown" }
    "$partner [$trustType - $direction]"
  }
  ## View handler
  $onView = {
    param($trust)

    $partner   = if ($trust.Partner) { $trust.Partner } else { $trust.Name }
    $trustType = if ($trust.Type) { $trust.Type } else { "Unknown" }
    $direction = if ($trust.Direction) { $trust.Direction } else { "Unknown" }

    ## Decode direction
    $directionDesc = switch ($direction) {
      "Inbound"       { "Inbound (users in $partner can access resources in $Domain)" }
      "Outbound"      { "Outbound (users in $Domain can access resources in $partner)" }
      "Bidirectional" { "Bidirectional (two-way trust, mutual access)" }
      default         { $direction }
    }

    ## Decode trust type
    $typeDesc = switch ($trustType) {
      "External"    { "External Trust (domain-to-domain in different forests)" }
      "Forest"      { "Forest Trust (all domains in both forests)" }
      "Shortcut"    { "Shortcut Trust (improves authentication performance)" }
      "Realm"       { "Realm Trust (cross-platform Kerberos trust)" }
      "ParentChild" { "Parent-Child Trust (automatic within forest)" }
      "TreeRoot"    { "Tree Root Trust (automatic within forest)" }
      default       { $trustType }
    }

    $details = @"
Trust Relationship Details
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Trusted Partner: $partner
Trust Type:      $trustType
$typeDesc

Direction:       $direction
$directionDesc

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Distinguished Name:
$(if ($trust.DN) { $trust.DN } else { "N/A" })

Created:  $(if ($trust.Created) { $trust.Created } else { "Unknown" })
Modified: $(if ($trust.Modified) { $trust.Modified } else { "Unknown" })

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Trust Status:
Use 'nltest /trusted_domains' or 'Get-ADTrust' to verify
trust health and test connectivity.

For trust management, use Active Directory Domains and Trusts
console (domain.msc) or PowerShell trust cmdlets.
"@

    Show-Modal "Trust Details" $details
  }

  ## Export handler
  $onExport = {
    param($items)
    try {
      $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
      $filename   = "trusts_${Domain}_$timestamp.csv"
      $exportData = $items | ForEach-Object {
        [PSCustomObject]@{
          Partner   = if ($_.Partner) { $_.Partner } else { $_.Name }
          Type      = $_.Type
          Direction = $_.Direction
          Created   = $_.Created
          Modified  = $_.Modified
          DN        = $_.DN
        }
      }
      $exportData | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
      Show-Modal "Success" "Exported $($items.Count) trust(s) to:`n$filename"
      Debug-Log "Exported trusts to $filename" -Type "Success"
    } catch {
      Show-Modal "Error" "Failed to export trusts:`n$($_.Exception.Message)"
      Debug-Log "Export failed: $($_.Exception.Message)" -Type "Problem"
    }
  }
  ## Show dialog using existing helper
  New-ListDialog -Title "Trust Relationships - $Domain" -Items $trusts -FormatItem $formatTrust -OnView $onView -OnExport $onExport -FilterHelp "(Filter by partner name or trust type)"
}

## TODO: This isn't called. but I feel ADHealth could be using it...?
function Test-TrustConnection {
  param($Trust)

  $partner = if ($Trust.Partner) { $Trust.Partner } else { $Trust.Name }
  Debug-Log "Testing trust connection to $partner" -Type "Insight"

  ## Show testing dialog
  $testDialog = [Terminal.Gui.Dialog]::new()
  $testDialog.Title = "Testing Trust - $partner"
  $testDialog.Width = 80
  $testDialog.Height = 25
  $y = 1

  $lblTesting = [Terminal.Gui.Label]::new("Testing trust relationship with: $partner")
  $lblTesting.X = 2
  $lblTesting.Y = $y
  $testDialog.Add($lblTesting)
  $y += 2

  $txtResults = [Terminal.Gui.TextView]::new()
  $txtResults.X = 2
  $txtResults.Y = $y
  $txtResults.Width = [Terminal.Gui.Dim]::Fill(2)
  $txtResults.Height = [Terminal.Gui.Dim]::Fill(4)
  $txtResults.ReadOnly = $true
  $testDialog.Add($txtResults)

  $btnClose = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = [Terminal.Gui.Pos]::Center()
  $btnClose.Y = [Terminal.Gui.Pos]::AnchorEnd(2)
  $btnClose.add_Clicked({ $testDialog.RequestStop() }).GetNewClosure()
  $testDialog.Add($btnClose)

  ## Run test in background
  [Terminal.Gui.Application]::MainLoop.Invoke({
    $results = @()
    $results += "Testing trust relationship..."
    $results += ""

    try {
      if ($Script:DemoMode -or (-not (Get-Command nltest -ErrorAction SilentlyContinue))) {
        # Demo mode or nltest not available
        $results += "⚠ Demo Mode or nltest not available"
        $results += ""
        $results += "To test trust in production, use:"
        $results += "  nltest /server:$env:COMPUTERNAME /sc_query:$partner"
        $results += ""
        $results += "Or use PowerShell:"
        $results += "  Test-ComputerSecureChannel -Credential (Get-Credential) -Server $partner"
        $results += ""
        $results += "Or use netdom:"
        $results += "  netdom trust $($Script:CurrentDomain) /domain:$partner /verify"
        $results += ""
        $results += "Simulated Result: ✓ Trust appears healthy"
      } else {
        ## Production Mode - actually test
        $results += "Running: nltest /sc_query:$partner"
        $results += ""

        $nlTestOutput = nltest /sc_query:$partner 2>&1
        $results += $nlTestOutput -join "`n"

        if ($LASTEXITCODE -eq 0) {
          $results += ""
          $results += "✓ Trust test successful"
        } else {
          $results += ""
          $results += "✗ Trust test failed (exit code: $LASTEXITCODE)"
        }
      }
    } catch {
      $results += "✗ Error testing trust:"
      $results += $_.Exception.Message
    }

    $txtResults.Text = $results -join "`n"
  }.GetNewClosure())
  [Terminal.Gui.Application]::Run($testDialog)
}

## Show Audit Log dialog
function Show-AuditLogDialog {
  <#
  .SYNOPSIS
  Show audit log for an AD object (User, Group, Computer)
  #>

  param(
    [Parameter(Mandatory=$true)]
    [object]$Object,

    [Parameter(Mandatory=$true)]
    [ValidateSet('User', 'Group', 'Computer')]
    [string]$ObjectType
  )

  $objectName = $Object.Name
  Debug-Log "Showing audit log for $ObjectType '$objectName'" -Type "Insight"

  ## ----------{ Enhanced Diagnostics }---------
  Debug-Log "========== Object Diagnostics ==========" -Type "Insight"
  Debug-Log "DemoMode: $($Script:DemoMode)" -Type "Tracing"
  Debug-Log "Object type: $($Object.GetType().Name)" -Type "Tracing"
  Debug-Log "Object has AuditLog property: $($null -ne $Object.AuditLog)" -Type "Tracing"

  ## List ALL properties
  if ($Object.PSObject.Properties) {
    Debug-Log "Object properties via PSObject:" -Type "Tracing"
    $propNames = $Object.PSObject.Properties | ForEach-Object { $_.Name }
    Debug-Log "Properties: $($propNames -join ', ')" -Type "Tracing"
    Debug-Log "Total properties: $($propNames.Count)" -Type "Tracing"

    ## Specifically check for AuditLog
    $auditLogProp = $Object.PSObject.Properties | Where-Object { $_.Name -eq 'AuditLog' }
    if ($auditLogProp) {
      Debug-Log "Found AuditLog property!" -Type "Insight"
      if ($null -ne $auditLogProp.Value) {
        Debug-Log "AuditLog type: $($auditLogProp.Value.GetType().Name)" -Type "Tracing"
        if ($auditLogProp.Value -is [array]) { Debug-Log "AuditLog count: $($auditLogProp.Value.Count)" -Type "Tracing" }
      } else {
        Debug-Log "AuditLog value is NULL" -Type "Warning"
      }
    } else {
      Debug-Log "AuditLog property NOT FOUND in PSObject.Properties" -Type "Warning"
    }
  }

  ## ----------{ Generate Audit log entries }---------
  $logEntries = @()

  if ($Script:DemoMode) {
    ## Check for audit log data in the object
    $hasAuditLog  = $false
    $auditLogData = $null

    ## Try different ways to access AuditLog
    if ($Object.PSObject.Properties['AuditLog']) {
      $auditLogData = $Object.PSObject.Properties['AuditLog'].Value
      if ($null -ne $auditLogData -and $auditLogData.Count -gt 0) {
        $hasAuditLog = $true
        Debug-Log "Found AuditLog via PSObject.Properties" -Type "Tracing"
      }
    } elseif ($Object.AuditLog) {
      $auditLogData = $Object.AuditLog
      if ($null -ne $auditLogData -and $auditLogData.Count -gt 0) {
        $hasAuditLog = $true
        Debug-Log "Found AuditLog via direct property access" -Type "Tracing"
      }
    } elseif ($Object -is [hashtable] -and $Object.ContainsKey('AuditLog')) {
      $auditLogData = $Object['AuditLog']
      if ($null -ne $auditLogData -and $auditLogData.Count -gt 0) {
        $hasAuditLog = $true
        Debug-Log "Found AuditLog in hashtable" -Type "Tracing"
      }
    }

    if ($hasAuditLog -and $auditLogData -and $auditLogData.Count -gt 0) {
      Debug-Log "Using object's audit log ($($auditLogData.Count) entries)" -Type "Insight"
      $logEntries = $auditLogData | ForEach-Object {
        [PSCustomObject]@{
          Timestamp  = if ($_.Timestamp -is [datetime]) {
            $_.Timestamp
          } else {
            try {
              [datetime]::Parse($_.Timestamp)
            } catch {
              Get-Date
            }
          }
          Action     = $_.Action
          Details    = $_.Details
          By         = $_.By
          ObjectType = $ObjectType
          ObjectName = $objectName
        }
      }
      $logEntries = $logEntries | Sort-Object Timestamp -Descending
    } else {
      ## No audit log in object - show modal and return
      Debug-Log "Object has no audit log data" -Type "Insight"
      Show-Modal "No Audit Log" "No audit log data is available for this $ObjectType.`n`nAudit logging may not be configured for this object."
      return
    }
  } else {
    ## Production Mode - Query Windows Event Log
    Debug-Log "Production Mode - querying Windows Event Log" -Type "Insight"
    try {
      $filter = @{
        LogName   = 'Security'
        ID        = 4720,4722,4723,4724,4725,4726,4738,4740,4767
        StartTime = (Get-Date).AddDays(-30)
      }
      $events = Get-WinEvent -FilterHashtable $filter -ErrorAction SilentlyContinue | Where-Object { $_.Message -match [regex]::Escape($objectName) } | Select-Object -First 50
      foreach ($event in $events) {
        $action = switch ($event.Id) {
          4720 { "Created" }
          4722 { "Account Status" }
          4723 { "Password Reset" }
          4724 { "Password Reset" }
          4725 { "Account Status" }
          4726 { "Deleted" }
          4738 { "Modified" }
          4740 { "Account Status" }
          4767 { "Account Status" }
          default { "Event" }
        }
        $details = switch ($event.Id) {
          4720 { "Account created" }
          4722 { "Account enabled" }
          4723 { "Password change attempted" }
          4724 { "Password reset attempted" }
          4725 { "Account disabled" }
          4726 { "Account deleted" }
          4738 { "Account modified" }
          4740 { "Account locked out" }
          4767 { "Account unlocked" }
          default { $event.Message.Split("`n")[0] }
        }
        $logEntries += [PSCustomObject]@{
          Timestamp  = $event.TimeCreated
          Action     = $action
          Details    = $details
          By         = "System"
          ObjectType = $ObjectType
          ObjectName = $objectName
        }
      }
      if ($logEntries.Count -eq 0) {
        Debug-Log "No audit events found in Windows Event Log" -Type "Insight"
        Show-Modal "No Audit Data" "No audit events found in the last 30 days for this $ObjectType."
        return
      } else {
        Debug-Log "Found $($logEntries.Count) audit events in Windows Event Log" -Type "Insight"
      }
    } catch {
      Debug-Log "Error querying Windows Event Log: $($_.Exception.Message)" -Type "Problem"
      Show-Modal "Error" "Failed to query audit log:`n`n$($_.Exception.Message)"
      return
    }
  }
  Debug-Log "Total log entries to display: $($logEntries.Count)" -Type "Insight"

  ## ----------{ Create Audit Log tab }---------
  $auditTab = @{
    Name    = "Audit Log"
    Builder = {
      param($view, $data, $state)

      ## Defensive NULL check
      if ($null -eq $view) {
        Debug-Log "ERROR: view parameter is null in Audit Log builder!" -Type "Problem"
        throw "View parameter is null"
      }

      ## Capture necessary variables and functions
      $entries        = $logEntries
      $objName        = $objectName
      $debugLogFunc   = ${function:Debug-Log}
      $showModalFunc  = ${function:Show-Modal}
      $formatDateFunc = ${function:Format-DateSafe}

      & $debugLogFunc "Building Audit Log tab with $($entries.Count) entries" -Type "Tracing"
      & $debugLogFunc "View is: $($view.GetType().Name)" -Type "Tracing"

      ## Format helper function - FIXED to return proper IList and audit log single-item bug
      function Local-FormatLogItems {
        param([array]$items)

        ## Create List<string> (implements IList) - prevents PowerShell unwrapping
        $list = [System.Collections.Generic.List[string]]::new()

        if (-not $items -or $items.Count -eq 0) {
          $list.Add("(No entries to display)")
          return $list  ## ← CRITICAL: Must return the list!
        }

        foreach ($entry in $items) {
          try {
            $dateStr = & $formatDateFunc $entry.Timestamp 'yyyy-MM-dd HH:mm'
            $formatted = "$dateStr | $($entry.Action.PadRight(15)) | $($entry.Details.PadRight(35)) | By: $($entry.By)"
            $list.Add($formatted)
          } catch {
            $list.Add("Error formatting entry")
          }
        }
        return $list  ## ← CRITICAL: Must return the list!
      }

      ## Header
      $lblHeader = [Terminal.Gui.Label]::new("Recent audit events for $objName ($($entries.Count))")
      $lblHeader.X = 1
      $lblHeader.Y = 0
      if ($null -eq $lblHeader) {
        & $debugLogFunc "ERROR: Failed to create header label!" -Type "Problem"
        throw "Failed to create label"
      }
      $view.Add($lblHeader)
      & $debugLogFunc "Added header label" -Type "Tracing"

      ## Filter label
      $lblFilter = [Terminal.Gui.Label]::new("Filter:")
      $lblFilter.X = 1
      $lblFilter.Y = 2
      $view.Add($lblFilter)

      ## Filter radio group
      $rdoFilter = [Terminal.Gui.RadioGroup]::new()
      $rdoFilter.X = 10
      $rdoFilter.Y = 2
      $rdoFilter.RadioLabels = [NStack.ustring[]]@(
        "All",
        "Modifications",
        "Password Resets",
        "Group Changes",
        "Account Status"
      )
      $view.Add($rdoFilter)
      & $debugLogFunc "Added filter controls" -Type "Tracing"

      ## ----------{ List View }---------
      $lstLog = [Terminal.Gui.ListView]::new()
      if ($null -eq $lstLog) {
        & $debugLogFunc "ERROR: Failed to create ListView!" -Type "Problem"
        throw "Failed to create ListView"
      }
      $lstLog.X = 1
      $lstLog.Y = 5
      $lstLog.Width  = [Terminal.Gui.Dim]::Fill(3)  ## Make room for vertical scrollbar
      $lstLog.Height = [Terminal.Gui.Dim]::Fill(6)  ## Make room for horizontal scrollbar
      $lstLog.CanFocus = $true

      ## Set initial source
      try {
        $formattedItems = Local-FormatLogItems $entries
        & $debugLogFunc "Formatted $($formattedItems.Count) items as IList" -Type "Tracing"
        if ($null -eq $formattedItems) {
          & $debugLogFunc "ERROR: formattedItems is null!" -Type "Problem"
          $fallbackList = [System.Collections.Generic.List[string]]::new()
          $fallbackList.Add("Error: formatted items is null")
          $formattedItems = $fallbackList
        }
        $lstLog.SetSource($formattedItems)
        & $debugLogFunc "Set ListView source successfully" -Type "Tracing"
      } catch {
        & $debugLogFunc "Error setting ListView source: $($_.Exception.Message)" -Type "Problem"
        $errorList = [System.Collections.Generic.List[string]]::new()
        $errorList.Add("Error displaying log entries: $($_.Exception.Message)")
        $lstLog.SetSource($errorList)
      }
      & $debugLogFunc "About to add ListView to view" -Type "Tracing"
      if ($null -eq $view) {
        & $debugLogFunc "ERROR: View became null before Add!" -Type "Problem"
        throw "View is null before Add"
      }
      if ($null -eq $lstLog) {
        & $debugLogFunc "ERROR: lstLog became null before Add!" -Type "Problem"
        throw "lstLog is null before Add"
      }
      $view.Add($lstLog)
      & $debugLogFunc "Added ListView to view" -Type "Tracing"

      ## ----------{ Scrollbars }---------
      try {
        ## Vertical scrollbar
        $vScrollBar = [Terminal.Gui.ScrollBarView]::new($lstLog, $true, $true)
        $vScrollBar.X = [Terminal.Gui.Pos]::AnchorEnd(1)
        $vScrollBar.Y = 5
        $vScrollBar.Width = 1
        $vScrollBar.Height = [Terminal.Gui.Dim]::Fill(6)
        $vScrollBar.add_ChangedPosition({
          $lstLog.TopItem = $vScrollBar.Position
          $lstLog.SetNeedsDisplay()
        }.GetNewClosure())
        $lstLog.add_DrawContent({
          if ($lstLog.Source) {
            $vScrollBar.Size = $lstLog.Source.Count
            $vScrollBar.Position = $lstLog.TopItem
            $vScrollBar.SetNeedsDisplay()
          }
        }.GetNewClosure())
        $view.Add($vScrollBar)
        & $debugLogFunc "Added vertical scrollbar" -Type "Tracing"

        ## Horizontal scrollbar
        $hScrollBar = [Terminal.Gui.ScrollBarView]::new($lstLog, $false, $true)
        $hScrollBar.X = 1
        $hScrollBar.Y = [Terminal.Gui.Pos]::AnchorEnd(4)
        $hScrollBar.Width = [Terminal.Gui.Dim]::Fill(3)
        $hScrollBar.Height = 1
        $hScrollBar.add_ChangedPosition({
          $lstLog.LeftItem = $hScrollBar.Position
          $lstLog.SetNeedsDisplay()
        }.GetNewClosure())

        $lstLog.add_DrawContent({
          if ($lstLog.Source -and $lstLog.Source.Count -gt 0) {
            $maxWidth = 0
            foreach ($item in $lstLog.Source) { if ($item -and $item.Length -gt $maxWidth) { $maxWidth = $item.Length }}
            $hScrollBar.Size = $maxWidth
            $hScrollBar.Position = $lstLog.LeftItem
            $hScrollBar.SetNeedsDisplay()
          }
        }.GetNewClosure())
        $view.Add($hScrollBar)
        & $debugLogFunc "Added horizontal scrollbar" -Type "Tracing"
      } catch {
        & $debugLogFunc "Error adding scrollbars: $($_.Exception.Message)" -Type "Warning"
      }

      ## ----------{ Filter Handler }---------
      $rdoFilter.add_SelectedItemChanged({
        try {
          $filtered = switch ($rdoFilter.SelectedItem) {
            0 { $entries }
            1 { $entries | Where-Object { $_.Action -match "Modified" } }
            2 { $entries | Where-Object { $_.Action -match "Password" } }
            3 { $entries | Where-Object { $_.Action -match "Group" } }
            4 { $entries | Where-Object { $_.Action -match "Account|Status" } }
          }
          if (-not $filtered -or $filtered.Count -eq 0) {
            $emptyList = [System.Collections.Generic.List[string]]::new()
            $emptyList.Add("(No entries match filter)")
            $lstLog.SetSource($emptyList)
          } else {
            $formattedFiltered = Local-FormatLogItems $filtered
            $lstLog.SetSource($formattedFiltered)
          }
        } catch {
          & $debugLogFunc "Filter error: $($_.Exception.Message)" -Type "Problem"
          $errorList = [System.Collections.Generic.List[string]]::new()
          $errorList.Add("Error filtering entries")
          $lstLog.SetSource($errorList)
        }
      }.GetNewClosure())

      ## ----------{ Export Button }---------
      $btnExport = [Terminal.Gui.Button]::new("Export to CSV")
      $btnExport.X = 1
      $btnExport.Y = [Terminal.Gui.Pos]::AnchorEnd(2)
      $btnExport.add_Clicked({
        try {
          $fn = "audit_log_${objName}_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
          $entries | Export-Csv $fn -NoTypeInformation -Encoding UTF8
          & $showModalFunc "Export Complete" "Exported to: $fn"
        } catch {
          & $showModalFunc "Export Failed" $_.Exception.Message
        }
      }.GetNewClosure())
      $view.Add($btnExport)
      & $debugLogFunc "Audit Log tab builder completed" -Type "Tracing"
    }
  }

  ## ----------{ Show dialog using New-PropertiesDialog }---------
  $dialogData = @{
    ObjectName = $objectName
    ObjectType = $ObjectType
    LogEntries = $logEntries
  }
  New-PropertiesDialog -Title "Audit Log - ${ObjectType}: $objectName" -Width 100 -Height 30 -Tabs @($auditTab) -Data $dialogData -IncludeSearchTab $false
}

## ----------{ Bulk attribute editor }---------
function Set-BulkAttribute {
  <#
  .SYNOPSIS
  Set a single attribute value across multiple AD objects

  .PARAMETER Objects
  Array of objects to modify

  .PARAMETER Attribute
  Attribute name to change (e.g., 'Company', 'Department', 'Description')

  .PARAMETER Value
  New value for the attribute

  .PARAMETER ShowDialog
  If specified, shows interactive dialog to select attribute and enter value

  .EXAMPLE
  Set-BulkAttribute -Objects $users -Attribute 'Company' -Value 'Acme Corp'

  .EXAMPLE
  Set-BulkAttribute -Objects $selectedUsers -ShowDialog
  #>

  param(
    [Parameter(Mandatory=$true)]
    [object[]]$Objects,
    [string]$Attribute,
    [string]$Value,
    [switch]$ShowDialog
  )

  if (-not $Objects -or $Objects.Count -eq 0) {
    Show-Modal "No Objects" "No objects provided"
    return
  }

  ## Detect object types
  $hasUsers     = $false
  $hasGroups    = $false
  $hasComputers = $false

  foreach ($obj in $Objects) {
    if ($obj.PSObject.Properties.Match('SamAccountName') -and -not $obj.PSObject.Properties.Match('ComputerType')) { $hasUsers = $true }
    elseif ($obj.PSObject.Properties.Match('Members')) { $hasGroups = $true }
    elseif ($obj.PSObject.Properties.Match('ComputerType')) { $hasComputers = $true }
  }

  ## Define available attributes by object type
  $userAttributes     = @( 'DisplayName', 'Description', 'EmailAddress', 'Title', 'Department', 'Company', 'Manager', 'OfficePhone', 'MobilePhone', 'StreetAddress', 'City', 'PostalCode', 'Country', '--- PASSWORD OPTIONS ---', 'ChangePasswordAtLogon', 'ResetPassword' )
  $groupAttributes    = @( 'Description', 'Email', 'ManagedBy' )
  $computerAttributes = @( 'Description', 'Location' )

  ## Build combined attribute list based on selected objects
  $availableAttributes = @()
  if ($hasUsers)     { $availableAttributes += $userAttributes     }
  if ($hasGroups)    { $availableAttributes += $groupAttributes    }
  if ($hasComputers) { $availableAttributes += $computerAttributes }
  $availableAttributes = $availableAttributes | Select-Object -Unique | Sort-Object

  ## ----------{ Interactive dialog }---------
  if ($ShowDialog) {
    $dlg = [Terminal.Gui.Dialog]::new("Bulk Attribute Editor", 70, 20)

    ## Info label
    $lblInfo = [Terminal.Gui.Label]::new("Editing $($Objects.Count) object(s)")
    $lblInfo.X = 2
    $lblInfo.Y = 1
    $dlg.Add($lblInfo)

    ## Object type info
    $typeInfo = @()
    if ($hasUsers) { $typeInfo += "Users" }
    if ($hasGroups) { $typeInfo += "Groups" }
    if ($hasComputers) { $typeInfo += "Computers" }
    $lblTypes = [Terminal.Gui.Label]::new("Types: $($typeInfo -join ', ')")
    $lblTypes.X = 2
    $lblTypes.Y = 2
    $dlg.Add($lblTypes)

    ## Attribute selection
    $lblAttribute = [Terminal.Gui.Label]::new("Select Attribute:")
    $lblAttribute.X = 2
    $lblAttribute.Y = 4
    $dlg.Add($lblAttribute)
    $lstAttributes = [Terminal.Gui.ListView]::new($availableAttributes)
    $lstAttributes.X = 2
    $lstAttributes.Y = 5
    $lstAttributes.Width = 66
    $lstAttributes.Height = 6
    $dlg.Add($lstAttributes)

    ## Value entry
    $lblValue = [Terminal.Gui.Label]::new("Enter New Value:")
    $lblValue.X = 2
    $lblValue.Y = 12
    $dlg.Add($lblValue)
    $txtValue = [Terminal.Gui.TextField]::new("")
    $txtValue.X = 2
    $txtValue.Y = 13
    $txtValue.Width = 66
    $dlg.Add($txtValue)

    ## Help label for password options
    $lblHelp = [Terminal.Gui.Label]::new("")
    $lblHelp.X = 2
    $lblHelp.Y = 14
    $lblHelp.Width = 66
    $dlg.Add($lblHelp)

    ## Update help text when attribute selection changes
    $lstAttributes.add_SelectedItemChanged({
      if ($lstAttributes.SelectedItem -ge 0) {
        $attr = $availableAttributes[$lstAttributes.SelectedItem]
        if ($attr -eq 'ChangePasswordAtLogon') {
          $lblHelp.Text  = [NStack.ustring]::Make("Value: 'true' or 'false'")
          $txtValue.Text = [NStack.ustring]::Make("true")
        } elseif ($attr -eq 'ResetPassword') {
          $lblHelp.Text  = [NStack.ustring]::Make("Enter new password (leave blank for random)")
          $txtValue.Text = [NStack.ustring]::Make("")
        } else {
          $lblHelp.Text = [NStack.ustring]::Make("")
        }
      }
    }.GetNewClosure())

    ## Apply button
    $btnApply = [Terminal.Gui.Button]::new("Apply")
    $btnApply.add_Clicked({
      if ($lstAttributes.SelectedItem -lt 0) {
        Show-Modal "No Selection" "Please select an attribute"
        return
      }
      $selectedAttr = $availableAttributes[$lstAttributes.SelectedItem]
      $newValue = $txtValue.Text.ToString()
      if ([string]::IsNullOrWhiteSpace($newValue)) {
        $confirm = Show-Modal "Empty Value" "Set attribute to empty/blank?`n`nAttribute: $selectedAttr" -YesNo
        if ($confirm -ne 0) { return }
      }
      [Terminal.Gui.Application]::RequestStop()

      ## Perform bulk update
      Set-BulkAttribute -Objects $Objects -Attribute $selectedAttr -Value $newValue
    }.GetNewClosure())

    $dlg.AddButton($btnApply)
    ## Cancel button
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
    $dlg.AddButton($btnCancel)
    [Terminal.Gui.Application]::Run($dlg)
    return
  }

  ## ----------{ Direct attribute update }---------
  if (-not $Attribute) {
    Show-Modal "Missing Attribute" "Attribute parameter is required when not using -ShowDialog"
    return
  }

  ## Validate attribute is available for these object types
  if ($Attribute -notin $availableAttributes) {
    Show-Modal "Invalid Attribute" "Attribute '$Attribute' is not valid for the selected object types.`n`nAvailable: $($availableAttributes -join ', ')"
    return
  }

  ## Confirmation
  $confirmMsg = "Set '$Attribute' to '$Value' for $($Objects.Count) object(s)?"
  $confirm = Show-Modal "Confirm Bulk Update" $confirmMsg -YesNo
  if ($confirm -ne 0) {
    Debug-Log "Bulk attribute update cancelled by user" -Type "Insight"
    return
  }
  Debug-Log "Starting bulk attribute update: $Attribute = '$Value' on $($Objects.Count) objects" -Type "Insight"

  ## Process objects
  $successCount = 0
  $failCount = 0
  $errors = @()

  foreach ($obj in $Objects) {
    $name = $obj.Name
    ## Determine object type
    $objectType = if ($obj.PSObject.Properties.Match('ComputerType')) { 'Computer' }
                  elseif ($obj.PSObject.Properties.Match('Members')) { 'Group' }
                  else { 'User' }
    ## Validate attribute
    $validForObject = $false
    if ($objectType -eq 'User' -and $Attribute -in $userAttributes)    { $validForObject = $true }
    if ($objectType -eq 'Group' -and $Attribute -in $groupAttributes)  { $validForObject = $true }
    if ($objectType -eq 'Computer' -and $Attribute -in $computerAttributes) { $validForObject = $true }
    if (-not $validForObject) {
        $failCount++
        $errors += "${name}: Attribute '$Attribute' not valid for $objectType"
        Debug-Log "Skipped $name - attribute not valid for $objectType" -Type "Warning"
        continue
    }

    try {
      if ($Script:DemoMode) {
        ## Demo mode
        if ($Attribute -eq 'ChangePasswordAtLogon') {
          $boolValue = [bool]::Parse($Value)
          if ($obj.PSObject.Properties.Match('ChangePasswordAtLogon')) { $obj.ChangePasswordAtLogon = $boolValue
          } else { $obj | Add-Member -NotePropertyName 'ChangePasswordAtLogon' -NotePropertyValue $boolValue -Force
          }
          $successCount++
          Debug-Log "Updated $objectType '$name': ChangePasswordAtLogon = $boolValue (demo mode)" -Type "Insight"
        } elseif ($Attribute -eq 'ResetPassword') {
          $passwordValue = if ([string]::IsNullOrWhiteSpace($Value)) {
            -join ((65..90) + (97..122) + (48..57) + (33,35,36,37,38,42,63) | Get-Random -Count 12 | ForEach-Object {[char]$_})
          } else {
            $Value
          }
          if ($obj.PSObject.Properties.Match('Password')) { $obj.Password = $passwordValue
          } else { $obj | Add-Member -NotePropertyName 'Password' -NotePropertyValue $passwordValue -Force
          }
          $successCount++
          Debug-Log "Reset password for $objectType '$name' (demo mode, password: $passwordValue)" -Type "Insight"
        } else {
          ## Standard attribute
          if ($obj.PSObject.Properties.Match($Attribute)) { $obj.$Attribute = $Value
          } else { $obj | Add-Member -NotePropertyName $Attribute -NotePropertyValue $Value -Force
          }
          $successCount++
          Debug-Log "Updated $objectType '$name': $Attribute = '$Value' (demo mode)" -Type "Insight"
        }
      } else {
        ## Production Mode
        if ($Attribute -eq 'ChangePasswordAtLogon') {
          $boolValue = [bool]::Parse($Value)
          Set-UnifiedObject -ObjectType User -Object $obj -Properties @{ ChangePasswordAtLogon = $boolValue }
          $successCount++
          Debug-Log "Updated $objectType '$name': ChangePasswordAtLogon = $boolValue in AD" -Type "Success"
        } elseif ($Attribute -eq 'ResetPassword') {
          $passwordValue = if ([string]::IsNullOrWhiteSpace($Value)) {
            $pw = -join ((65..90) + (97..122) + (48..57) + (33,35,36,37,38,42,63) | Get-Random -Count 12 | ForEach-Object { [char]$_ } )
            ConvertTo-SecureString -String $pw -AsPlainText -Force
            } else { ConvertTo-SecureString -String $Value -AsPlainText -Force
            }
            Set-ADAccountPassword -Identity $obj.SamAccountName -NewPassword $passwordValue -Reset -ErrorAction Stop
            $successCount++
            Debug-Log "Reset password for $objectType '$name' in AD" -Type "Success"
        } else {
          ## Standard attribute update via unified setter
          $properties = @{}

          switch ($Attribute) {
            'EmailAddress'  { $properties['EmailAddress'] = $Value }
            'Email'         { $properties['mail']  = $Value ; $properties['Email'] = $Value }
            'DisplayName'   { $properties['DisplayName']   = $Value }
            'Description'   { $properties['Description']   = $Value }
            'Title'         { $properties['Title']         = $Value }
            'Department'    { $properties['Department']    = $Value }
            'Company'       { $properties['Company']       = $Value }
            'Manager'       { $properties['Manager']       = $Value }
            'ManagedBy'     { $properties['ManagedBy']     = $Value }
            'OfficePhone'   { $properties['OfficePhone']   = $Value }
            'MobilePhone'   { $properties['MobilePhone']   = $Value }
            'StreetAddress' { $properties['StreetAddress'] = $Value }
            'City'          { $properties['City']          = $Value }
            'PostalCode'    { $properties['PostalCode']    = $Value }
            'Country'       { $properties['Country']       = $Value }
            'Location'      { $properties['Location']      = $Value }
          }
          Set-UnifiedObject -ObjectType $objectType -Object $obj -Properties $properties
          $successCount++
          Debug-Log "Updated $objectType '$name': $Attribute = '$Value' in AD" -Type "Success"
        }
      } ## End Production Mode
    } catch {
      $failCount++
      $errors += "${name}: $($_.Exception.Message)"
      Debug-Log "Failed to update $objectType '$name': $($_.Exception.Message)" -Type "Problem"
    }
  } ## End foreach
}

## ----------{ Unified Refresh Domain Data }----------
function Refresh-Data {
  <#
  .SYNOPSIS
  Refreshes Active Directory data and optionally rebuilds the tree view

  .PARAMETER Domain
  Domain to refresh (defaults to current domain)

  .PARAMETER RebuildTree
  Rebuild the tree view after refreshing data

  .PARAMETER ShowModal
  Show success modal dialog when complete

  .PARAMETER ShowLoadingDialog
  Show loading dialog during refresh
  #>
  param(
    [string]$Domain,
    [switch]$RebuildTree,
    [switch]$ShowModal,
    [switch]$ShowLoadingDialog
  )

  if (-not $Domain) { $Domain = $Script:CurrentDomain }
  Debug-Log "Refreshing data for domain: $Domain" -Type "Insight"

  ## Show loading dialog if requested
  $loadingDlg = $null
  if ($ShowLoadingDialog) { $loadingDlg = Show-LoadingDialog -Message "Refreshing Active Directory data..." }
  try {
    ## Demo mode - reload from raw data
    if ($Script:DemoMode) {
      Set-StatusBar "Refreshing demo data..." -Icon 'Working'
      ## Check if we have ImportedRawData to refresh from
      if ($null -eq $Script:ImportedRawData) {
        Debug-Log "No ImportedRawData available - loading default demo data" -Type "Warning"
        Load-DefaultDemoData
      } else {
        ## Determine source label for status messages
        $sourceLabel = if ($Script:ImportSource) { $Script:ImportSource } else { "demo data" }
        Debug-Log "Refreshing from $sourceLabel" -Type "Insight"
        ## RE-CONVERT using the preserved raw data (File, AD, or Demo)
        $baseDN = "DC=$($Domain -replace '\.',',DC=')"
        Convert-DataToADObjects -Users $Script:ImportedRawData.Users -DCs $Script:ImportedRawData.DCs -Computers $Script:ImportedRawData.Computers -Groups $Script:ImportedRawData.Groups -Domain $Domain -BaseDN $baseDN
      }
      if ($RebuildTree) {
        Set-StatusBar "Rebuilding tree..." -Icon 'Working'
        [Terminal.Gui.Application]::MainLoop.Invoke({
          try {
            $Script:tree.ClearObjects()
            Build-Tree -domain $Domain
            Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
            [Terminal.Gui.Application]::Refresh()
          } catch {
            Debug-Log "Tree rebuild error: $_" -Type "Problem"
          }
        })
      }
      Set-StatusBar "Demo data refreshed" -Icon 'Success'
      if ($ShowModal) { Show-Modal "Refreshed" "Demo data refreshed successfully" }
      return $true
    }

    ## Production Mode - query AD
    Set-StatusBar "Loading domain controllers..." -Icon 'Working'
    $dcs       = Get-ADDomainController -Filter * -Server $Domain -ErrorAction Stop
    Set-StatusBar "Loading users..." -Icon 'Working'
    $users     = Get-ADUser -Filter * -Server $Domain -ErrorAction Stop
    Set-StatusBar "Loading groups..." -Icon 'Working'
    $groups    = Get-ADGroup -Filter * -Server $Domain -ErrorAction Stop
    Set-StatusBar "Loading computers..." -Icon 'Working'
    $computers = Get-ADComputer -Filter * -Server $Domain -ErrorAction Stop

    ## Validate results
    if ($null -eq $dcs -or $null -eq $users -or $null -eq $groups -or $null -eq $computers) {
      Debug-Log "Refresh failed - one or more queries returned null" -Type "Warning"
      Set-StatusBar "Refresh failed - check logs" -Icon 'Success'
      return $false
    }
    Debug-Log "Loaded $($users.Count) users, $($groups.Count) groups, $($computers.Count) computers, $($dcs.Count) DCs" -Type "Insight"

    ## Convert to standardized format
    Set-StatusBar "Converting data..." -Icon 'Working'
    $baseDN = "DC=$($Domain -replace '\.',',DC=')"

    ## Convert raw AD objects to our format
    $userList = $users | ForEach-Object {
      @{
        Name           = $_.Name
        SamAccountName = $_.SamAccountName
        Email          = $_.EmailAddress
        Title          = $_.Title
        Department     = $_.Department
        Disabled       = -not $_.Enabled
        Locked         = $_.LockedOut
        Domain         = $Domain
      }
    }
    $groupList = $groups | ForEach-Object {
      @{
        Name        = $_.Name
        Description = $_.Description
        Type        = $_.GroupCategory
        Scope       = $_.GroupScope
        Domain      = $Domain
      }
    }
    $computerList = $computers | ForEach-Object {
      @{
        Name      = $_.Name
        OS        = $_.OperatingSystem
        OSVersion = $_.OperatingSystemVersion
        Enabled   = $_.Enabled
        Domain    = $Domain
      }
    }
    $dcList = $dcs | ForEach-Object {
      @{
        Name            = $_.Name
        Domain          = $_.Domain
        Site            = $_.Site
        IsGlobalCatalog = $_.IsGlobalCatalog
        OS              = $_.OperatingSystem
      }
    }
    ## Store the raw AD data for future refreshes
    $Script:ImportedRawData = @{
      Users     = $userList
      Groups    = $groupList
      Computers = $computerList
      DCs       = $dcList
    }
    $Script:ImportSource = 'ActiveDirectory'
    Convert-DataToADObjects -Users $userList -DCs $dcList -Computers $computerList -Groups $groupList -Domain $Domain -BaseDN $baseDN

    ## Update data source tracking
    $Script:DataSource = "AD"
    $Script:DataSourceInfo = @{
      Source       = "AD"
      Server       = $Domain
      LoadedAt     = Get-Date
      IsReadOnly   = $false
      ObjectCounts = @{
        Users      = $Script:Users.Count
        Groups     = $Script:Groups.Count
        Computers  = $Script:Computers.Count
        DCs        = $Script:DCs.Count
      }
    }

    ## Rebuild tree if requested
    if ($RebuildTree) {
      Set-StatusBar "Rebuilding tree..." -Icon 'Working'
      [Terminal.Gui.Application]::MainLoop.Invoke({
        try {
          $Script:tree.ClearObjects()
          Build-Tree -domain $Domain
          Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
          Show-InfoPanel -UpdateOnly
          [Terminal.Gui.Application]::Refresh()
        } catch {
          Debug-Log "Tree rebuild error: $_" -Type "Problem"
        }
      })
    }

    Set-StatusBar "Refresh complete" -Icon 'Success'
    if ($ShowModal) { Show-Modal "Refreshed" "Active Directory data refreshed successfully" }
    return $true
  } catch {
    Debug-Log "Refresh error: $_" -Type "Problem"
    Set-StatusBar "Refresh error" -Icon 'Success'
    if ($ShowModal) { Show-Modal "Error" "Failed to refresh data:`n`n$($_.Exception.Message)" }
    return $false
  } finally {
    ## Close loading dialog if it was shown
    if ($loadingDlg) { Close-LoadingDialog $loadingDlg }
  }
}

## Is it a special day...?
## Determines which unicode emoji to use for the AD window title.
function Initialise-DirectoryEmoji {
  param(
    [DateTime]$Date = (Get-Date)
  )

  $month = $Date.Month
  $day   = $Date.Day
  ## Default: card index
  $emoji = "🗂️"

  ## ----------{ Icon Initialisation }----------
  if ($Script:HasTerminalIcons) {

    ## Nerd Font / Symbol icons (explicit Unicode escapes) & Safe for Terminal.Gui when font supports Nerd Fonts
    $Script:DirectoryEmoji = "`u{F115}"   ## nf-fa-folder
    $Script:FileEmoji      = "`u{F016}"   ## nf-fa-file
    $Script:UserEmoji      = "`u{F007}"   ## nf-fa-user

    ## Additional commonly used icons
    $Script:ComputerEmoji  = "`u{F108}"   ## nf-fa-desktop
    $Script:GroupEmoji     = "`u{F0C0}"   ## nf-fa-users
    $Script:DomainEmoji    = "`u{F233}"   ## nf-fa-server
    $Script:WarningEmoji   = "`u{F071}"   ## nf-fa-exclamation_triangle
    $Script:OkEmoji        = "`u{F00C}"   ## nf-fa-check
    $Script:ErrorEmoji     = "`u{F00D}"   ## nf-fa-times
  } else {
    ## ASCII fallbacks (alignment-safe)
    $Script:DirectoryEmoji = "[D]"
    $Script:FileEmoji      = "[F]"
    $Script:UserEmoji      = "[U]"
    $Script:ComputerEmoji  = "[PC]"
    $Script:GroupEmoji     = "[G]"
    $Script:DomainEmoji    = "[DC]"
    $Script:WarningEmoji   = "!"
    $Script:OkEmoji        = "OK"
    $Script:ErrorEmoji     = "X"
  }

  switch ($true) {
    { $month -eq 1  -and $day -eq 1  }                   { $emoji = "📅" ; break }                    ## 1st Jan New Year’s Day
    { $month -eq 1  -and $day -eq 2  }                   { $emoji = "🦄" ; break }                    ## Wild haggis Hunting
    { $month -eq 4  -and $day -eq 9  }                   { $emoji = "`u{1F1E9}`u{1F1F0}" ; break }     ## 9th Apr Danmark's besættelse (liberation day)
    { $month -eq 5  -and $day -eq 4  }                   { $emoji = "🕯️" ; break }                    ## 4th May Candle for Besættelsen
    { $month -eq 6  -and $day -eq 5  }                   { $emoji = "`u{1F1E9}`u{1F1F0}" ; break }     ## 5th June Constitution day in Danmark
    { $month -eq 5  -and $day -eq 21 }                   { $emoji = "`u{1F1EC}`u{1F1F1}" ; break }     ## 21st May Grønland Day
    { $month -eq 7  -and $day -eq 1  }                   { $emoji = "🇨🇦" ; break }                     ## 1st Jul Canada Day
    { $month -eq 7  -and $day -eq 4  }                   { $emoji = "🫖" ; break }                    ## 4th Jul Teapot (to annoy Americans)
    { $month -eq 7  -and $day -eq 29 }                   { $emoji = "`u{1F1EB}`u{1F1F4}" ; break }     ## 29th Jul Faroe Islands
    { $month -eq 11 -and $day -eq 9  }                   { $emoji = "`u{1F1E9}`u{1F1EA}" ; break }     ## 9th Nov Erich Honecker leck mich am Arsch!
    { $month -eq 11 -and $day -eq 24 }                   { $emoji = "👑" ; break }                    ## 24th Nov (Så ta'r vi den en gang til for Prins Knud B-))
    { $month -eq 11 -and $day -eq 30 }                   { $emoji = "🏴󠁧󠁢󠁳󠁣󠁴󠁿" ; break }                    ## 30th Nov St Andrew’s Day (Saltire)
    { $month -eq 12 -and ($day -eq 24 -or $day -eq 25) } { $emoji = "🎄" ; break }                    ## 24th/25th Dec Tree for Jul / Christmas
  }

  if ($Script:HasTerminalIcons) {
    ## Nerd Font glyphs (Terminal.Gui-safe)
    $Script:DirectoryEmoji = ""  ## nf-fa-folder (U+F115)
    $Script:FileEmoji      = ""  ## nf-fa-file (U+F016)
    $Script:UserEmoji      = ""  ## nf-fa-user (U+F007)
  } else {
    # ASCII fallback
    $Script:DirectoryEmoji = "[D]"
    $Script:FileEmoji      = "[F]"
    $Script:UserEmoji      = "[U]"
  }
  Debug-Log ("The emoji Today is: $emoji") -Type "Insight"
  $Script:DirectoryEmoji = $emoji
}

## ----------{Icon initialisation function }----------
## Call this AFTER module checks, before building the tree

function Initialise-Icons {
  <#
  .SYNOPSIS
  Initialise icon set based on available modules

  .DESCRIPTION
  Uses NerdFonts glyphs when available, falls back to emoji/text
  #>

  if ($Script:hasNerdFonts -and $Script:HasTerminalIcons) {
    Debug-Log "NerdFonts available - using glyph icons" -Type "Insight"

    ## NerdFonts glyphs for AD object types
    $Script:Icons = @{
      ## Core AD object types (NerdFonts)
      User      = [char]0xf007  # nf-fa-user
      Group     = [char]0xf0c0  # nf-fa-users
      Computer  = [char]0xf108  # nf-fa-desktop
      Laptop    = [char]0xf109  # nf-fa-laptop
      Server    = [char]0xf233  # nf-fa-server
      OU        = [char]0xf07b  # nf-fa-folder
      DC        = [char]0xf233  # nf-fa-server (domain controller)
      Forest    = [char]0xf1bb  # nf-fa-sitemap
      Domain    = [char]0xf0ac  # nf-fa-globe

      ## Status indicators (NerdFonts)
      Locked    = [char]0xf023  # nf-fa-lock
      Unlocked  = [char]0xf09c  # nf-fa-unlock
      Disabled  = [char]0xf05e  # nf-fa-ban
      Enabled   = [char]0xf058  # nf-fa-check_circle
      Warning   = [char]0xf071  # nf-fa-exclamation_triangle
      Error     = [char]0xf06a  # nf-fa-exclamation_circle
      Success   = [char]0xf00c  # nf-fa-check
      Info      = [char]0xf05a  # nf-fa-info_circle

      ## Additional useful icons
      Password  = [char]0xf084  # nf-fa-key
      Email     = [char]0xf0e0  # nf-fa-envelope
      Phone     = [char]0xf095  # nf-fa-phone
      Calendar  = [char]0xf073  # nf-fa-calendar
      Clock     = [char]0xf017  # nf-fa-clock_o

      ## Device types
      Printer   = [char]0xf02f  # nf-fa-print
      Mobile    = [char]0xf10b  # nf-fa-mobile
      Tablet    = [char]0xf10a  # nf-fa-tablet
    }
  } else {
    Debug-Log "NerdFonts not available - using emoji/text icons" -Type "Insight"

    ## Emoji fallback
    $Script:Icons = @{
      ## Core AD object types (Emoji)
      User      = "👤"
      Group     = "👥"
      Computer  = "💻"
      Laptop    = "💻"
      Server    = "🖥️"
      OU        = "📁"
      DC        = "🖥️"
      Forest    = "🌲"
      Domain    = "🌐"

      ## Status indicators (Emoji/Text)
      Locked    = "🔒"
      Unlocked  = "🔓"
      Disabled  = "⊗"
      Enabled   = "●"
      Warning   = "⚠️"
      Error     = "✖"
      Success   = "✔"
      Info      = "ℹ"

      ## Additional useful icons
      Password  = "🔑"
      Email     = "📧"
      Phone     = "📞"
      Calendar  = "📅"
      Clock     = "⏰"

      ## Device types
      Printer   = "🖨️"
      Mobile    = "📱"
      Tablet    = "📱"
    }
  }
  Debug-Log "Icon set Initialised with $($Script:Icons.Count) icons" -Type "Success"
}

## ----------{Helper function to get icon }----------
function Get-Icon {
  param(
    [string]$IconName,
    [string]$Fallback = ""
  )

  if ($Script:Icons.ContainsKey($IconName)) {
    return $Script:Icons[$IconName]
  } else {
    return $Fallback
  }
}

## ----------{ Convert Domain Data }----------
function Convert-DataToADObjects {
  param(
    [array]$Users,
    [array]$DCs = @(),
    [array]$Computers = @(),
    [array]$Groups = @(),
    [string]$Domain = "example.com",
    [string]$BaseDN = "DC=example,DC=com"
  )

  Debug-Log "Converting demo data to AD-like objects..." -Type "Tracing"

  ## Helper functions
  function New-FakeGuid { [guid]::NewGuid().ToString() }
  function New-FakeSid {
    $rid = Get-Random -Minimum 1000 -Maximum 65535
    "S-1-5-21-{0}-{1}-{2}-{3}" -f (Get-Random -Max 999999999),(Get-Random -Max 999999999),(Get-Random -Max 999999999),$rid
  }

  ## Helper to convert Manager name to DN
  function Convert-ManagerToDN {
    param($ManagerName, $AllUsers, $BaseDN)

    if ([string]::IsNullOrWhiteSpace($ManagerName)) { return "" }

    ## Try to find manager in user list
    $manager = $AllUsers | Where-Object { $_.Name -eq $ManagerName }
    if ($manager) {
      ## Use manager's email to determine domain
      $mgrDomain = if ($manager.Email -and $manager.Email -match '@(.+)$') {
        $matches[1]
      } else {
        $Domain
      }

      ## Build DN from manager's OU if available
      if ($manager.OU) {
        $ouChain = $manager.OU | ForEach-Object { "OU=$_" }
        [array]::Reverse($ouChain)
        return "CN=$ManagerName," + ($ouChain -join ',') + ",$BaseDN"
      }
    }
    ## Fallback: simple DN
    return "CN=$ManagerName,$BaseDN"
  }

  ## Users
  $convertedUsers = @()
  foreach ($user in $Users) {
    $sam = if ($user.SamAccountName) {
      $user.SamAccountName
    } else {
      ($user.Name -replace '\s+', '.').ToLower()
    }
    $upn = if ($user.UserPrincipalName) {
      $user.UserPrincipalName
    } elseif ($user.Email) {
      $user.Email
    } else {
      "$sam@$Domain"
    }
    $userDomain = if ($user.Email -and $user.Email -match '@(.+)$') {
      $matches[1]
    } else {
      $Domain
    }
    if ($user.OU) {
      $ouChain = $user.OU | ForEach-Object { "OU=$_" }
      [array]::Reverse($ouChain)
      $dn = "CN=$($user.Name)," + ($ouChain -join ',') + ",$BaseDN"
    } else {
      $dn = "CN=$($user.Name),$BaseDN"
    }

    ## Build base user object with standard AD properties
    $adUser = [PSCustomObject]@{
      ## Core Identity (using standard AD property names)
      ObjectClass                   = 'user'
      Name                          = $user.Name
      cn                            = $user.Name
      Domain                        = $userDomain
      sAMAccountName                = $sam
      userPrincipalName             = $upn
      displayName                   = $user.Name
      givenName                     = ($user.Name -split '\s+')[0]
      sn                            = ($user.Name -split '\s+')[-1]  # surname
      distinguishedName             = $dn

      ## GUIDs and SIDs
      objectGUID                    = New-FakeGuid
      objectSid                     = New-FakeSid

      ## AD Metadata
      instanceType                  = 4
      objectCategory                = "CN=Person,CN=Schema,CN=Configuration,$BaseDN"
      primaryGroupID                = 513  # Domain Users

      ## Status Properties
      Enabled                       = (-not $user.Disabled)
      Disabled                      = [bool]$user.Disabled
      LockedOut                     = [bool]$user.Locked
      PasswordExpired               = [bool]$user.MustChangePassword
      userAccountControl            = if ($user.Disabled) { 514 } else { 512 }  # 512=Enabled, 514=Disabled

      ## Timestamps
      whenCreated                   = (Get-Date).AddDays(-90)
      whenChanged                   = (Get-Date).AddDays(-5)
      PasswordLastSet               = (Get-Date).AddDays(-30)
      pwdLastSet                    = if ($user.MustChangePassword) { 0 } else { 134091109521968821 }
      LastLogonDate                 = if ($user.LastLogonDate) {
        if ($user.LastLogonDate -is [DateTime]) {
          $user.LastLogonDate
        } elseif ($user.LastLogonDate -is [string]) {
          try { [DateTime]::Parse($user.LastLogonDate) } catch { (Get-Date).AddHours(-(Get-Random -Minimum 1 -Maximum 72)) }
        } else {
          $user.LastLogonDate
        }
      } else {
        (Get-Date).AddHours(-(Get-Random -Minimum 1 -Maximum 72))
      }
      lastLogon                     = if ($user.LastLogonDate) {
        $dt = if ($user.LastLogonDate -is [DateTime]) {
          $user.LastLogonDate
        } elseif ($user.LastLogonDate -is [string]) {
          try { [DateTime]::Parse($user.LastLogonDate) } catch { (Get-Date).AddHours(-(Get-Random -Minimum 1 -Maximum 72)) }
        } else {
          (Get-Date).AddHours(-(Get-Random -Minimum 1 -Maximum 72))
        }
        $dt.ToFileTime()
      } else {
        (Get-Date).AddHours(-(Get-Random -Minimum 1 -Maximum 72)).ToFileTime()
      }
      lastLogonTimestamp            = if ($user.LastLogonDate) {
        $dt = if ($user.LastLogonDate -is [DateTime]) {
          $user.LastLogonDate
        } elseif ($user.LastLogonDate -is [string]) {
          try { [DateTime]::Parse($user.LastLogonDate) } catch { (Get-Date).AddHours(-(Get-Random -Minimum 1 -Maximum 72)) }
        } else {
          (Get-Date).AddHours(-(Get-Random -Minimum 1 -Maximum 72))
        }
        $dt.ToFileTime()
      } else {
        (Get-Date).AddHours(-(Get-Random -Minimum 1 -Maximum 72)).ToFileTime()
      }

      ## Password & Logon
      badPwdCount                   = if ($user.Locked) { (Get-Random -Minimum 3 -Maximum 10) } else { 0 }
      logonCount                    = Get-Random -Minimum 0 -Maximum 500
      accountExpires                = if ($user.AccountExpirationDate) {
        $dt = if ($user.AccountExpirationDate -is [DateTime]) {
          $user.AccountExpirationDate
        } elseif ($user.AccountExpirationDate -is [string]) {
          try { [DateTime]::Parse($user.AccountExpirationDate) } catch { $null }
        } else {
          $null
        }
        if ($dt) { $dt.ToFileTime() } else { 0 }
      } else {
        0
      }
      lockoutTime                   = if ($user.Locked) { (Get-Date).ToFileTime() } else { 0 }

      ## User Details
      title                         = $user.Title
      description                   = $user.Description

      ## Contact Info (standard AD property names)
      mail                          = $user.Email
      telephoneNumber               = if ($user.telephoneNumber) {
        $user.telephoneNumber
      } elseif ($user.Phone) {
        $user.Phone
      } else {
        ""
      }
      mobile                        = if ($user.mobile) {
        $user.mobile
      } elseif ($user.MobilePhone) {
        $user.MobilePhone
      } else {
        ""
      }
      physicalDeliveryOfficeName    = if ($user.physicalDeliveryOfficeName) {
        $user.physicalDeliveryOfficeName
      } elseif ($user.Office) {
        $user.Office
      } else {
        ""
      }

      ## Address Properties
      streetAddress                 = if ($user.Street) { $user.Street } else { "" }
      l                             = if ($user.City) { $user.City } else { "" }  # locality
      postalCode                    = if ($user.PostalCode) { $user.PostalCode } else { "" }
      c                             = if ($user.Country) { $user.Country } else { "" }  # country
      co                            = if ($user.Country) { $user.Country } else { "" }  # country name

      ## Manager (will be set to DN in second pass)
      manager                       = ""

      ## Extension Attributes for Department/Company
      extensionAttribute1           = if ($user.Department) { $user.Department } else { "" }
      extensionAttribute2           = if ($user.Company) { $user.Company } else { "" }

      ## Keep friendly names for tool compatibility
      Department                    = $user.Department
      Company                       = $user.Company

      ## Group Membership
      memberOf                      = if ($user.Groups) { $user.Groups } else { @() }

      ## OU Path and Audit
      OU                            = $user.OU
      AuditLog                      = $user.AuditLog
    }

    ## Preserve all extra user properties (LAPS, TPM, BitLocker, etc.)
    if ($user -is [hashtable]) {
      $standardProps = @(
        'Name','cn','Domain','sAMAccountName','userPrincipalName', 'Email','mail','title','Department','Company',
        'manager','Phone','telephoneNumber','MobilePhone','mobile','Office','physicalDeliveryOfficeName','Street',
        'streetAddress','City','PostalCode','Country','description','Groups','memberOf','OU','Disabled','Locked',
        'MustChangePassword','displayName','givenName','sn','distinguishedName','AuditLog','LastLogonDate'
      )

      foreach ($prop in $user.Keys) {
        if ($prop -in $standardProps) { continue }
        if ($null -ne $user[$prop] -and $user[$prop] -ne '') { $adUser | Add-Member -NotePropertyName $prop -NotePropertyValue $user[$prop] -Force }
      }
    }
    $adUser.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADUser')
    $convertedUsers += $adUser
  }

  ## Second pass: Convert Manager names to DNs now that we have all users
  foreach ($adUser in $convertedUsers) {
    $originalUser = $Users | Where-Object { $_.Name -eq $adUser.Name } | Select-Object -First 1
    if ($originalUser -and $originalUser.Manager) {
      $managerDN = Convert-ManagerToDN -ManagerName $originalUser.Manager -AllUsers $Users -BaseDN $BaseDN
      $adUser.manager = $managerDN
    }
  }

  ## Domain Controllers
  $convertedDCs = @()

  foreach ($dc in $DCs) {
    $dcDomain = if ($dc.Domain) { $dc.Domain }
                elseif ($dc.Location -match 'Germany') { 'example.net' }
                else { $Domain }
    $dn   = "CN=$($dc.Name),OU=Domain Controllers,$BaseDN"
    $adDC = [PSCustomObject]@{
      ObjectClass            = 'computer'
      Name                   = $dc.Name
      cn                     = $dc.Name
      Domain                 = $dcDomain
      dNSHostName            = "$($dc.Name).$dcDomain"
      distinguishedName      = $dn
      objectGUID             = New-FakeGuid
      objectSid              = New-FakeSid
      instanceType           = 4
      objectCategory         = "CN=Computer,CN=Schema,CN=Configuration,$BaseDN"
      Enabled                = $true
      userAccountControl     = 4096  # Workstation/Server trust account
      Site                   = $dc.Site
      Location               = $dc.Location
      operatingSystem        = $dc.OS
      operatingSystemVersion = $dc.OSVersion
      IPv4Address            = $dc.IPv4Address
      IsGlobalCatalog        = $dc.IsGlobalCatalog
      FSMORoles              = $dc.FSMORoles
      Services               = $dc.Services
      whenCreated            = (Get-Date).AddDays(-180)
      whenChanged            = (Get-Date).AddDays(-10)
      OU                     = 'Domain Controllers'
      AuditLog               = $dc.AuditLog
    }

    ## Preserve all extra DC properties
    if ($dc -is [hashtable]) {
      $standardProps = @(
        'Name','cn','Domain','dNSHostName','distinguishedName','Site','Location','operatingSystem',
        'operatingSystemVersion','IPv4Address','IsGlobalCatalog','FSMORoles','Services','OU','AuditLog'
      )

      foreach ($prop in $dc.Keys) {
        if ($prop -in $standardProps) { continue }
        if ($null -ne $dc[$prop] -and $dc[$prop] -ne '') { $adDC | Add-Member -NotePropertyName $prop -NotePropertyValue $dc[$prop] -Force }
      }
    }
    $adDC.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADComputer')
    $convertedDCs += $adDC
  }

  ## Groups
  $convertedGroups = @()
  foreach ($group in $Groups) {
    $sam = if ($group.SamAccountName) {
      $group.SamAccountName
    } else {
      ($group.Name -replace '\s+', '.').ToLower()
    }
    $dn  = "CN=$($group.Name),OU=Groups,$BaseDN"
    $groupDomain = if ($group.Email -match '@(.+)$') { $matches[1] } else { $Domain }

    $adGroup = [PSCustomObject]@{
      ObjectClass       = 'group'
      Name              = $group.Name
      cn                = $group.Name
      Domain            = $groupDomain
      sAMAccountName    = $sam
      distinguishedName = $dn
      objectGUID        = New-FakeGuid
      objectSid         = New-FakeSid
      instanceType      = 4
      objectCategory    = "CN=Group,CN=Schema,CN=Configuration,$BaseDN"
      GroupCategory     = $group.Type
      GroupScope        = $group.Scope
      groupType         = -2147483646  # Global Security Group (default)
      description       = $group.Description
      managedBy         = $group.ManagedBy
      mail              = $group.Email
      whenCreated       = (Get-Date).AddDays(-180)
      whenChanged       = (Get-Date).AddDays(-10)
      AuditLog          = $group.AuditLog
    }

    ## Preserve all extra group properties
    if ($group -is [hashtable]) {
      $standardProps = @(
        'Name','cn','Domain','sAMAccountName','description','GroupCategory','GroupScope','managedBy',
        'Members','MemberOf','AuditLog','Email','mail','Type','Scope','distinguishedName'
      )

      foreach ($prop in $group.Keys) {
        if ($prop -in $standardProps) { continue }
        if ($null -ne $group[$prop] -and $group[$prop] -ne '') {
          $adGroup | Add-Member -NotePropertyName $prop -NotePropertyValue $group[$prop] -Force
        }
      }
    }
    $adGroup.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADGroup')
    $convertedGroups += $adGroup
  }

  ## Computers
  $convertedComputers = @()

  foreach ($computer in $Computers) {
    $dn = if ($computer.OU) {
      $ouChain = $computer.OU | ForEach-Object { "OU=$_" }
      "CN=$($computer.Name)," + ($ouChain[-1..0] -join ',') + ",$BaseDN"
    } else {
      "CN=$($computer.Name),CN=Computers,$BaseDN"
    }
    $compDomain = if ($computer.Domain) { $computer.Domain }
                  elseif ($computer.Location -match 'Germany') { 'example.net' }
                  else { $Domain }
    $adComputer = [PSCustomObject]@{
      ObjectClass            = 'computer'
      Name                   = $computer.Name
      cn                     = $computer.Name
      Domain                 = $compDomain
      dNSHostName            = "$($computer.Name).$compDomain"
      distinguishedName      = $dn
      objectGUID             = New-FakeGuid
      objectSid              = New-FakeSid
      instanceType           = 4
      objectCategory         = "CN=Computer,CN=Schema,CN=Configuration,$BaseDN"
      Enabled                = $computer.Enabled
      userAccountControl     = if ($computer.Enabled -eq $false) { 4098 } else { 4096 }
      operatingSystem        = $computer.OS
      operatingSystemVersion = $computer.OSVersion
      IPv4Address            = $computer.IPv4Address
      LastLogonDate          = $computer.LastLogon
      lastLogon              = if ($computer.LastLogon) {
        $dt = if ($computer.LastLogon -is [DateTime]) {
          $computer.LastLogon
        } elseif ($computer.LastLogon -is [string]) {
          try { [DateTime]::Parse($computer.LastLogon) } catch { $null }
        } else {
          $null
        }
        if ($dt) { $dt.ToFileTime() } else { 0 }
      } else {
        0
      }
      Location               = $computer.Location
      description            = $computer.Description
      ComputerType           = $computer.Type
      whenCreated            = (Get-Date).AddDays(-120)
      whenChanged            = (Get-Date).AddDays(-7)
      OU                     = $computer.OU
      AuditLog               = $computer.AuditLog
    }

    ## Preserve all extra computer properties
    if ($computer -is [hashtable]) {
      $standardProps = @(
        'Name','cn','Domain','operatingSystem','operatingSystemVersion','IPv4Address','LastLogon','lastLogon',
        'Location','description','Type','OU','Enabled','AuditLog','dNSHostName','distinguishedName'
      )
      foreach ($prop in $computer.Keys) {
        if ($prop -in $standardProps) { continue }
        if ($null -ne $computer[$prop] -and $computer[$prop] -ne '') { $adComputer | Add-Member -NotePropertyName $prop -NotePropertyValue $computer[$prop] -Force }
      }
    }
    $adComputer.PSObject.TypeNames.Insert(0, 'Microsoft.ActiveDirectory.Management.ADComputer')
    $convertedComputers += $adComputer
  }

  ## Globals
  $Script:Users     = $convertedUsers
  $Script:Computers = $convertedComputers
  $Script:DCs       = $convertedDCs
  $Script:Groups    = $convertedGroups
  $Script:ADObjects = $convertedUsers + $convertedDCs + $convertedGroups + $convertedComputers

  Debug-Log "Converted $($convertedUsers.Count) users, $($convertedDCs.Count) DCs, $($convertedComputers.Count) computers, and $($convertedGroups.Count) groups" -Type "Tracing"

  return @{
    Users     = $convertedUsers
    DCs       = $convertedDCs
    Groups    = $convertedGroups
    Computers = $convertedComputers
  }
}

##----------{ Convwert Object to Tree Items }----------
function Convert-ToTreeNode {
  param($node)

  ## Determine node type from the backing object
  if ($node -is [OUNode]) {
    $type = 'ou'
    $text = "$($Script:Icons.OU) $($node.Name)"
    $tag  = @{ Type='ou'; Object=$node }
  }
  elseif ($node.Type -eq 'group') {
    $type = 'group'
    $text = "$($Script:Icons.Group) $($node.Name)"
    $tag  = @{ Type='group'; Object=$node }
  }
  elseif ($node.Type -eq 'user') {
    $type = 'user'
    $text = "$($Script:Icons.User) $($node.Name)"
    $tag  = @{ Type='user'; Object=$node }
  }
  elseif ($node.Type -eq 'computer') {
    $type = 'computer'
    $text = "$($Script:Icons.Computer) $($node.Name)"
    $tag  = @{ Type='computer'; Object=$node }
  }
  else {
    $type = 'unknown'
    $text = $node.Name
    $tag  = @{ Type='unknown'; Object=$node }
  }

  $treenode = [Terminal.Gui.Trees.TreeNode]::new($text)
  $treenode.Tag = $tag
  foreach ($child in $node.Children) { $treenode.Children.Add((Convert-ToTreeNode $child)) }
  return $treenode
}

## ----------{ helper function for safe date formatting }----------
function Format-DateSafe {
  <#
  .SYNOPSIS
  Safely formats a date value that might be DateTime or String

  .DESCRIPTION
  Handles date values from both CSV (strings) and TDF (DateTime objects)

  .PARAMETER DateValue
  The date value to format (DateTime object or string)

  .PARAMETER Format
  The format string (default: 'yyyy-MM-dd HH:mm')

  .EXAMPLE
  Format-DateSafe $user.LastLogonDate

  .EXAMPLE
  Format-DateSafe $computer.whenCreated 'yyyy-MM-dd'
  #>
  param(
    $DateValue,
    [string]$Format = 'yyyy-MM-dd HH:mm'
  )

  if ($null -eq $DateValue -or $DateValue -eq '') { return 'Never' }
  ## If it's already a DateTime, format it
  if ($DateValue -is [DateTime]) { return $DateValue.ToString($Format) }

  ## If it's a string, try to parse it
  if ($DateValue -is [string]) {
    $parsedDate = $null
    if ([DateTime]::TryParse($DateValue, [ref]$parsedDate)) { return $parsedDate.ToString($Format) }
    ## Couldn't parse, return as-is
    return $DateValue
  }

  ## Unknown type, convert to string
  return $DateValue.ToString()
}

## ----------{ build domain content }----------
function Build-DomainContent {
  param(
    [Parameter(Mandatory)]
    [Terminal.Gui.Trees.TreeNode]$domainNode,
    [string]$domain,
    [bool]$ShowGroups = $true,
    [bool]$ShowDCs = $true,
    [bool]$ShowComputers = $true,
    [bool]$ShowOUs = $true
  )

  ## Get icons with fallback
  $userIcon     = Get-Icon 'User'
  $groupIcon    = Get-Icon 'Group'
  $computerIcon = Get-Icon 'Computer'
  $dcIcon       = Get-Icon 'DC'
  $ouIcon       = Get-Icon 'OU'
  $disabledIcon = Get-Icon 'Disabled'
  $lockedIcon   = Get-Icon 'Locked'

  Debug-Log "Building content for domain: $domain" -Type "Insight"

  ## ----------{Users & OUs section }----------
  ## Step 1: Get users for this domain
  $domainUsers = $Script:Users | Where-Object { $_.Domain -eq $domain }
  Debug-Log "Filtered to $($domainUsers.Count) users for domain $domain" -Type "Insight"

  ## Step 2: Apply combined filters (name, quick filter, checkboxes)
  $domainUsers = Apply-CombinedFilters -Users $domainUsers
  Debug-Log "After applying filters: $($domainUsers.Count) users" -Type "Insight"

  if ($domainUsers.Count -gt 0) {
    ## Diagnostics
    $firstUser = $domainUsers[0]
    Debug-Log "=== First User Diagnostic ===" -Type "Insight"
    Debug-Log "Name: $($firstUser.Name)" -Type "Tracing"
    Debug-Log "Has OU: $($null -ne $firstUser.OU)" -Type "Tracing"
    if ($firstUser.OU) { Debug-Log "OU values: $($firstUser.OU -join ' > ')" -Type "Tracing" }

    ## Count users with/without OUs
    $usersWithOU = $domainUsers | Where-Object { $_.OU -and $_.OU.Count -gt 0 }
    $usersWithoutOU = $domainUsers | Where-Object { -not $_.OU -or $_.OU.Count -eq 0 }
    Debug-Log "Users WITH OU: $($usersWithOU.Count)" -Type "Insight"
    Debug-Log "Users WITHOUT OU: $($usersWithoutOU.Count)" -Type "Insight"

    ## Build OU tree structure using standalone function
    $ouTree = Build-OUTree -users $domainUsers

    ## Add OU nodes to tree
    Debug-Log "=== Adding OU nodes to tree ===" -Type "Insight"
    $addedOUs = 0

    foreach ($ouName in ($ouTree.Keys | Sort-Object)) {
      if ($ouName -eq '_ROOT_USERS') {
        ## Add root users directly to domain (sorted alphabetically)
        Debug-Log "Adding $($ouTree['_ROOT_USERS'].Users.Count) root users" -Type "Tracing"
        foreach ($user in ($ouTree[$ouName].Users | Sort-Object -Property Name)) {
          $statusIcons = ""
          if (-not $user.Enabled -or $user.Disabled) { $statusIcons += " $disabledIcon" }
          if ($user.LockedOut -or $user.Locked) { $statusIcons += " $lockedIcon" }

          $userNodeText = "$userIcon $($user.Name)$statusIcons"
          $userNode = [Terminal.Gui.Trees.TreeNode]::new($userNodeText)
          $userNode.Tag = @{ Type = 'User'; Object = $user; Name = $user.Name }
          $domainNode.Children.Add($userNode)
        }
      } else {
        ## Add OU tree using standalone function
        $addedOUs++
        Debug-Log "Calling Add-OUNode for: $ouName" -Type "Tracing"
        Add-OUNode -parentNode $domainNode -ouData $ouTree[$ouName] -ouName $ouName -depth 0
      }
    }
    Debug-Log "Total OUs added: $addedOUs" -Type "Insight"
  }

  ## ----------{Groups section }----------
  if ($ShowGroups) {
    $domainGroups = $Script:Groups | Where-Object { $_.Domain -eq $domain }

    if ($domainGroups.Count -gt 0) {
      Debug-Log "Building Groups container with $($domainGroups.Count) groups" -Type "Insight"
      $groupsNode = [Terminal.Gui.Trees.TreeNode]::new("$groupIcon Groups")
      $groupsNode.Tag = @{ Type = 'Container'; Object = $null; Name = 'Groups' }
      foreach ($group in ($domainGroups | Sort-Object -Property Name)) {
        $groupNodeText = "$groupIcon $($group.Name)"
        $groupNode = [Terminal.Gui.Trees.TreeNode]::new($groupNodeText)
        $groupNode.Tag = @{ Type = 'Group'; Object = $group; Name = $group.Name }
        $groupsNode.Children.Add($groupNode)
      }
      $domainNode.Children.Add($groupsNode)
      Debug-Log "Added Groups container with $($domainGroups.Count) groups" -Type "Success"
    }
  }

  ## ----------{Domain controllers section }----------
  if ($ShowDCs) {
    $domainDCs = $Script:DCs | Where-Object { $_.Domain -eq $domain }
    if ($domainDCs.Count -gt 0) {
      Debug-Log "Building Domain Controllers container with $($domainDCs.Count) DCs" -Type "Insight"
      $dcsNode = [Terminal.Gui.Trees.TreeNode]::new("$dcIcon Domain Controllers")
      $dcsNode.Tag = @{ Type = 'Container'; Object = $null; Name = 'Domain Controllers' }
      foreach ($dc in ($domainDCs | Sort-Object -Property Name)) {
        $dcNodeText = "$dcIcon $($dc.Name)"
        $dcNode = [Terminal.Gui.Trees.TreeNode]::new($dcNodeText)
        $dcNode.Tag = @{ Type = 'DC'; Object = $dc; Name = $dc.Name }
        $dcsNode.Children.Add($dcNode)
      }
      $domainNode.Children.Add($dcsNode)
      Debug-Log "Added Domain Controllers container with $($domainDCs.Count) DCs" -Type "Success"
    }
  }

  ## ----------{Computers/devices section (Grouped by Type) }----------
  if ($ShowComputers) {
    $domainComputers = $Script:Computers | Where-Object { $_.Domain -eq $domain }
    if ($domainComputers.Count -gt 0) {
      Debug-Log "Building Computers container with $($domainComputers.Count) devices" -Type "Insight"
      $computersNode = [Terminal.Gui.Trees.TreeNode]::new("$computerIcon Computers & Devices")
      $computersNode.Tag = @{ Type = 'Container'; Object = $null; Name = 'Computers' }

      ## Group by ComputerType (converted AD objects use ComputerType, not Type)
      $devicesByType = $domainComputers | Group-Object -Property { if ($_.ComputerType) { $_.ComputerType } else { "Unknown" }} | Sort-Object -Property Name
      Debug-Log "Found $($devicesByType.Count) device type groups" -Type "Insight"

      foreach ($typeGroup in $devicesByType) {
        $typeName = $typeGroup.Name
        $devices  = $typeGroup.Group
        Debug-Log "Adding device type group: '$typeName' with $($devices.Count) devices" -Type "Tracing"

        ## Create type container node
        $pluralName   = if ($typeName -eq 'Unknown') { 'Unknown Devices' } else { "${typeName}s" }
        $typeNodeText = "$computerIcon $pluralName"
        $typeNode     = [Terminal.Gui.Trees.TreeNode]::new($typeNodeText)
        $typeNode.Tag = @{ Type = 'Container'; Object = $null; Name = $typeName }

        ## Add devices in this type (sorted alphabetically)
        foreach ($computer in ($devices | Sort-Object -Property Name)) {
          $statusIcons = ""
          if (-not $computer.Enabled -or $computer.Disabled) { $statusIcons += " $disabledIcon" }
          $computerNodeText = "$computerIcon $($computer.Name)$statusIcons"
          $computerNode     = [Terminal.Gui.Trees.TreeNode]::new($computerNodeText)
          $computerNode.Tag = @{ Type = 'Computer'; Object = $computer; Name = $computer.Name }
          $typeNode.Children.Add($computerNode)
        }
        $computersNode.Children.Add($typeNode)
      }
      $domainNode.Children.Add($computersNode)
      Debug-Log "Added Computers container with $($devicesByType.Count) device types" -Type "Success"
    }
  }
  Debug-Log "Finished building content for domain $domain" -Type "Success"
}

function Add-OUNode {
  <#
  .SYNOPSIS
  Recursively adds OU nodes and their users to a TreeView
  .DESCRIPTION
  Creates a TreeView node for an OU, adds all users in that OU,
  then recursively processes child OUs
  .PARAMETER parentNode
  The parent TreeNode to add this OU to
  .PARAMETER ouData
  Hashtable containing Users array and Children hashtable
  .PARAMETER ouName
  Name of this OU
  .PARAMETER depth
  Current recursion depth (for logging indentation)
  #>
  param(
    [Parameter(Mandatory)]
    $parentNode,
    [Parameter(Mandatory)]
    [hashtable]$ouData,
    [Parameter(Mandatory)]
    [string]$ouName,
    [int]$depth = 0
  )

  ## Get icons from script scope with fallback
  $userIcon     = Get-Icon 'User'
  $ouIcon       = Get-Icon 'OU'
  $disabledIcon = Get-Icon 'Disabled'
  $lockedIcon   = Get-Icon 'Locked'

  ## Logging with indentation - Uncomment this if you'd like Debug-Logs to "mirror" the tree view
  #$indent = "  " * $depth
  Debug-Log "$indent[Add-OUNode] Processing: $ouName" -Type "Tracing"
  Debug-Log "$indent[Add-OUNode] Users: $($ouData.Users.Count), Child OUs: $($ouData.Children.Keys.Count)" -Type "Tracing"

  if ($ouData.Children.Keys.Count -gt 0) { Debug-Log "$indent[Add-OUNode] Child OU names: $($ouData.Children.Keys -join ', ')" -Type "Tracing" }

  ## Create OU node
  $ouNodeText = "$ouIcon $ouName"
  $ouNode     = [Terminal.Gui.Trees.TreeNode]::new($ouNodeText)
  $ouNode.Tag = @{ Type = 'OU'; Object = @{ Name = $ouName }; Name = $ouName }
  $parentNode.Children.Add($ouNode)
  Debug-Log "$indent[Add-OUNode] OU node created and added to parent" -Type "Tracing"

  ## Add users in this OU (sorted alphabetically)
  $addedUsers = 0
  foreach ($user in ($ouData.Users | Sort-Object -Property Name)) {
    $statusIcons = ""
    if (-not $user.Enabled -or $user.Disabled) { $statusIcons += " $disabledIcon" }
    if ($user.LockedOut -or $user.Locked) { $statusIcons += " $lockedIcon" }
    $userNodeText = "$userIcon $($user.Name)$statusIcons"
    $userNode     = [Terminal.Gui.Trees.TreeNode]::new($userNodeText)
    $userNode.Tag = @{ Type = 'User'; Object = $user; Name = $user.Name }
    $ouNode.Children.Add($userNode)
    $addedUsers++
  }
  if ($addedUsers -gt 0) { Debug-Log "$indent[Add-OUNode] Added $addedUsers users" -Type "Tracing" }

  ## Recursively add child OUs (sorted alphabetically)
  foreach ($childOUName in ($ouData.Children.Keys | Sort-Object)) {
    Debug-Log "$indent[Add-OUNode] Recursing into: $childOUName" -Type "Tracing"
    Add-OUNode -parentNode $ouNode -ouData $ouData.Children[$childOUName] -ouName $childOUName -depth ($depth + 1)
  }
  Debug-Log "$indent[Add-OUNode] Completed: $ouName (total children: $($ouNode.Children.Count))" -Type "Tracing"
}

## ----------{Build-OuTree - Fixed Version }----------
function Build-OUTree {
  param(
    [Parameter(Mandatory)]
    [array]$users
  )

  Debug-Log "Building OU tree from $($users.Count) users..." -Type "Tracing"
  $ouTree = @{}
  $processedUsers = 0

  foreach ($user in $users) {
    if ($user.OU -and $user.OU.Count -gt 0) {
      $processedUsers++
      if ($processedUsers -le 5) { Debug-Log "Processing user $($processedUsers): $($user.Name) - OU: $($user.OU -join ' > ')" -Type "Tracing" }
      ## Build hierarchy - navigate down creating nodes as needed
      $currentLevel = $ouTree
      foreach ($ouName in $user.OU) {
        if (-not $currentLevel.ContainsKey($ouName)) {
          $currentLevel[$ouName] = @{
            Users    = @()
            Children = @{}
          }
          if ($processedUsers -le 5) { Debug-Log "Created OU structure: $ouName" -Type "Tracing" }
        }
        ## KEY FIX: Navigate through .Children to get to next level
        $currentLevel = $currentLevel[$ouName].Children
      }

      ## FIXED: Add user to the deepest OU by navigating correctly. We need to go back up a level because we went into .Children
      $deepestOU = $ouTree
      for ($i = 0; $i -lt $user.OU.Count; $i++) {
        $ouName = $user.OU[$i]
        if ($i -eq ($user.OU.Count - 1)) {
          ## Last OU - this is where the user goes
          $deepestOU[$ouName].Users += $user
          if ($processedUsers -le 5) { Debug-Log "Added user to OU '$ouName' (now has $($deepestOU[$ouName].Users.Count) users)" -Type "Tracing" }
        } else {
          ## Not the last - keep navigating down through Children
          $deepestOU = $deepestOU[$ouName].Children
        }
      }
    } else {
      ## No OU path - add to root
      if (-not $ouTree.ContainsKey('_ROOT_USERS')) {
        $ouTree['_ROOT_USERS'] = @{
          Users    = @()
          Children = @{}
        }
      }
      $ouTree['_ROOT_USERS'].Users += $user
    }
  }
  Debug-Log "OU tree built - processed $processedUsers users with OUs" -Type "Tracing"
  Debug-Log "OU tree has $($ouTree.Keys.Count) top-level keys" -Type "Tracing"
  Debug-Log "Top-level OU names: $($ouTree.Keys -join ', ')" -Type "Insight"
  return $ouTree
}

##----------{ Build The Tree }----------
function Build-Tree {
  param([string]$domain)

  ## Capture functions at top
  $debugLogFunc           = ${function:Debug-Log}
  $buildDomainContentFunc = ${function:Build-DomainContent}

  if (-not $domain) { $domain = $Script:CurrentDomain }
  & $debugLogFunc "Building tree..." -Type "Insight"

  ## TreeView MUST already exist in 1.16
  if ($null -eq $Script:tree) {
    throw "Build-Tree failed: TreeView does not exist"
  }

  ## Log current filter state for debugging
  & $debugLogFunc "Filter check - ShowGroups: $($Script:FilterOptions.ShowGroups), ShowDCs: $($Script:FilterOptions.ShowDCs), ShowComputers: $($Script:FilterOptions.ShowComputers), ShowOUs: $($Script:FilterOptions.ShowOUs)" -Type "Tracing"

  ## Clear existing objects safely
  try {
    $Script:tree.ClearObjects()
  } catch {
    & $debugLogFunc "WARNING - ClearObjects failed: $_" -Type "Warning"
  }
  $root = $null

  ## ----------{ Multi-domain forest }----------
  if ($Script:Domains.Count -gt 1) {
    & $debugLogFunc "Creating multi-domain forest tree with root: $($Script:ForestName)" -Type "Insight"
    $root = [Terminal.Gui.Trees.TreeNode]::new($Script:ForestName)

    foreach ($dom in $Script:Domains) {
      & $debugLogFunc "Adding domain node: $($dom)" -Type "Insight"
      $domainNode = [Terminal.Gui.Trees.TreeNode]::new($dom)
      $root.Children.Add($domainNode)

      ## Build content with visibility filters from FilterOptions
      & $buildDomainContentFunc -domainNode $domainNode -domain $dom -ShowGroups $Script:FilterOptions.ShowGroups -ShowDCs $Script:FilterOptions.ShowDCs -ShowComputers $Script:FilterOptions.ShowComputers -ShowOUs $Script:FilterOptions.ShowOUs
    }

    if (-not $Script:CurrentDomain) {
      $Script:CurrentDomain = $Script:Domains[0]
      & $debugLogFunc "CurrentDomain was empty. Auto-selected: $($Script:CurrentDomain)" -Type "Insight"
    }
  }
  ## ----------{Single-domain }----------
  else {
    & $debugLogFunc "Creating single-domain tree: $($Script:Domains[0])" -Type "Insight"
    $root = [Terminal.Gui.Trees.TreeNode]::new($Script:Domains[0])

    # Build content with visibility filters from FilterOptions
    & $buildDomainContentFunc -domainNode $root -domain $Script:Domains[0] -ShowGroups $Script:FilterOptions.ShowGroups -ShowDCs $Script:FilterOptions.ShowDCs -ShowComputers $Script:FilterOptions.ShowComputers -ShowOUs $Script:FilterOptions.ShowOUs
    if (-not $Script:CurrentDomain) {
      $Script:CurrentDomain = $Script:Domains[0]
      & $debugLogFunc "CurrentDomain set (single-domain): $($Script:CurrentDomain)" -Type "Insight"
    }
  }

  if ($null -eq $root) {
    throw "Build-Tree failed: root node is null"
  }

  ## ----------{Attach root to TreeView }----------
  try {
    $Script:tree.AddObject($root)
    $Script:tree.SelectedObject = $root
    & $debugLogFunc "Root node added to TreeView" -Type "Success"
  } catch {
    throw "Build-Tree failed while attaching root to TreeView: $_"
  }
  & $debugLogFunc "Build-Tree completed successfully" -Type "Success"
  return $root
}

## ----------{ Filter Label Function }----------
function Manage-FilterStatusLabel {
  <#
  .SYNOPSIS
  Create or update the filter status label

  .PARAMETER Action
  'Create' to make a new label, 'Update' to refresh existing label text

  .PARAMETER Label
  Required for 'Update' action - the label to update

  .PARAMETER X
  X position for new label (default: 80)

  .PARAMETER Y
  Y position for new label (default: 6)

  .PARAMETER Width
  Width for new label (default: 40). Use [Terminal.Gui.Dim]::Fill(1) for dynamic width

  .PARAMETER InPanel
  Set to $true when creating label inside FilterPanel (uses Fill width)

  .EXAMPLE
  ## Create standalone label
  $lblStatus = Manage-FilterStatusLabel -Action 'Create'

  ## Create label inside panel
  $lblStatus = Manage-FilterStatusLabel -Action 'Create' -X 1 -Y 15 -InPanel

  ## Update existing label
    Manage-FilterStatusLabel -Action 'Update' -Label $lblStatus
  #>

  param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Create', 'Update')]
    [string]$Action,
    [Parameter(Mandatory=$false)]
    [Terminal.Gui.Label]$Label,
    [int]$X = 80,
    [int]$Y = 6,
    [object]$Width = 40,  ## Can be int or Dim object
    [switch]$InPanel
  )

  if ($Action -eq 'Create') {
    Debug-Log "Creating filter status label" -Type "Insight"
    try {
      $lblStatus = [Terminal.Gui.Label]::new("")
      $lblStatus.X = $X
      $lblStatus.Y = $Y
      ## Handle Width - can be int or Dim object
      if ($InPanel) {
        $lblStatus.Width = [Terminal.Gui.Dim]::Fill(1)
      } else {
        $lblStatus.Width = $Width
      }
      Debug-Log "Filter status label created successfully" -Type "Success"
      return $lblStatus
    } catch {
      Debug-Log "ERROR creating filter status label: $($_.Exception.Message)" -Type "Problem"
      return $null
    }
  }
  if ($Action -eq 'Update') {
    if (-not $Label) {
      Debug-Log "Label parameter is null in Update action" -Type "Warning"
      return
    }
    Debug-Log "Updating filter status label" -Type "Insight"

    ## Build active filters list
    $activeFilters = @()

    if (-not $Script:FilterOptions.ShowEnabledUsers)  { $activeFilters += "No Enabled"   }
    if (-not $Script:FilterOptions.ShowDisabledUsers) { $activeFilters += "No Disabled"  }
    if (-not $Script:FilterOptions.ShowGroups)        { $activeFilters += "No Groups"    }
    if (-not $Script:FilterOptions.ShowDCs)           { $activeFilters += "No DCs"       }
    if (-not $Script:FilterOptions.ShowComputers)     { $activeFilters += "No Computers" }
    if ($Script:FilterOptions.NameFilter)             { $activeFilters += "Name:$($Script:FilterOptions.NameFilter)" }

    ## Update label text
    if ($activeFilters.Count -gt 0) {
      $Label.Text = "Active Filters: " + ($activeFilters -join ", ")
    } else {
      $Label.Text = "No filters active (showing all)"
    }
    Debug-Log "Filter status label updated: $($activeFilters.Count) filters active" -Type "Insight"
  }
}

## ----------{ LDAP filter helper }----------
function Get-LDAPFilteredObject {
  <#
  .SYNOPSIS
  Apply LDAP-style filters to AD users or groups

  .DESCRIPTION
  Unified filter dispatcher for user and group collections using parameter sets

  .EXAMPLE
  Get-LDAPFilteredObject -ObjectType User -FilterType PasswordExpired -Users $Script:Users

  .EXAMPLE
  Get-LDAPFilteredObject -ObjectType Group -FilterType EmptyGroups -Groups $Script:Groups
  #>

  [CmdletBinding(DefaultParameterSetName='User')]
  param(

    [Parameter(Mandatory=$true)]
    [ValidateSet('User','Group')]
    [string]$ObjectType,

    ## ----------{ Users filters }----------
    [Parameter(Mandatory=$true, ParameterSetName='User')]
    [ValidateSet(
      'All', 'LockedOnly', 'DisabledOnly', 'EnabledOnly', 'NeverLoggedIn', 'NoManager', 'PasswordExpired', 'PasswordExpiring72h',
      'PasswordNeverExpires', 'AccountExpired', 'AccountExpiring30d', 'StaleAccounts90d', 'EmptyEmail', 'EmptyDepartment' )]
    [string]$FilterType,

    [Parameter(ParameterSetName='User')]
    [array]$Users = $Script:Users,

    ## ----------{ Group filters }----------
    [Parameter(Mandatory=$true, ParameterSetName='Group')]
    [ValidateSet('All','EmptyGroups','SecurityGroups','DistributionGroups')]
    [string]$GroupFilterType,

    [Parameter(ParameterSetName='Group')]
    [array]$Groups = $Script:Groups
  )

  $now = Get-Date

  switch ($ObjectType) {

    'User' {
      switch ($FilterType) {
        'All'                  { return $Users }
        'LockedOnly'           { return $Users | Where-Object { $_.LockedOut -or $_.Locked }}
        'DisabledOnly'         { return $Users | Where-Object { $_.Disabled -or -not $_.Enabled }}
        'EnabledOnly'          { return $Users | Where-Object { -not $_.Disabled -and $_.Enabled }}
        'NeverLoggedIn'        { return $Users | Where-Object { -not $_.LastLogonDate -or $_.LastLogonDate -eq [DateTime]::MinValue }}
        'NoManager'            { return $Users | Where-Object { [string]::IsNullOrWhiteSpace($_.Manager) }}
        'PasswordExpired'      { return $Users | Where-Object { $_.PasswordExpired -eq $true }}
        'PasswordNeverExpires' { return $Users | Where-Object { $_.PasswordNeverExpires -eq $true }}
        'AccountExpired'       { return $Users | Where-Object { $_.AccountExpirationDate -and $_.AccountExpirationDate -lt $now }}
        'AccountExpiring30d'   {
          $window = $now.AddDays(30)
          return $Users | Where-Object {
            $_.AccountExpirationDate -and
            $_.AccountExpirationDate -gt $now -and
            $_.AccountExpirationDate -le $window
          }
        }
        'StaleAccounts90d'     {
          $staleDate = $now.AddDays(-90)
          return $Users | Where-Object { $_.LastLogonDate -and $_.LastLogonDate -lt $staleDate }
        }
        'EmptyEmail'           {
          return $Users | Where-Object {
            [string]::IsNullOrWhiteSpace($_.EmailAddress) -and
            [string]::IsNullOrWhiteSpace($_.mail)
          }
        }
        'EmptyDepartment'      {
          return $Users | Where-Object {
            [string]::IsNullOrWhiteSpace($_.Department)
          }
        }
        'PasswordExpiring72h'  {
          $expiryDays    = 90
          $warningWindow = $now.AddHours(72)

          return $Users | Where-Object {
            if ($_.PasswordNeverExpires -or -not $_.PasswordLastSet) { return $false }
            $expiryDate = $_.PasswordLastSet.AddDays($expiryDays)
            $expiryDate -le $warningWindow -and $expiryDate -gt $now
          }
        }
      }
    }
    'Group' {
      switch ($GroupFilterType) {
        'All'                { return $Groups }
        'EmptyGroups'        { return $Groups | Where-Object { -not $_.Members -or $_.Members.Count -eq 0 }}
        'SecurityGroups'     { return $Groups | Where-Object { $_.GroupCategory -eq 'Security' -or $_.Type -eq 'Security' }}
        'DistributionGroups' { return $Groups | Where-Object { $_.GroupCategory -eq 'Distribution' -or $_.Type -eq 'Distribution' }}
      }
    }
  }
}

function Apply-CombinedFilters {
  <#
  .SYNOPSIS
  Apply all active filters from FilterOptions

  .DESCRIPTION
  Central function that applies name filter + quick filter + checkboxes
  Returns filtered user collection
  #>

  param(
    [array]$Users = $Script:Users
  )

  $filtered = $Users

  ## Apply quick filter if set
  if ($Script:FilterOptions.QuickFilter -and $Script:FilterOptions.QuickFilter -ne 'All') {
    $filtered = Get-LDAPFilteredObject -ObjectType User -FilterType $Script:FilterOptions.QuickFilter -Users $filtered
  }

  ## Apply enabled/disabled checkboxes (only when no quick filter overrides)
  if (-not $Script:FilterOptions.QuickFilter -or $Script:FilterOptions.QuickFilter -eq 'All') {
    $filtered = $filtered | Where-Object {($_.Disabled -and $Script:FilterOptions.ShowDisabledUsers) -or (-not $_.Disabled -and $Script:FilterOptions.ShowEnabledUsers)}
  }

  ## Apply name filter with operator
  $nameFilter = $Script:FilterOptions.NameFilter
  if ($nameFilter) {
    $nameFilter = $nameFilter.Trim()
    if ($nameFilter) {
      switch ($Script:FilterOptions.NameOperator) {
        'Contains'   { $filtered = $filtered | Where-Object { $_.Name -like "*${nameFilter}*" -or $_.EmailAddress -like "*${nameFilter}*" -or $_.Title -like "*${nameFilter}*"} }
        'StartsWith' { $filtered = $filtered | Where-Object { $_.Name -like "${nameFilter}*" -or $_.EmailAddress -like "${nameFilter}*" -or $_.Title -like "${nameFilter}*"} }
        'EndsWith'   { $filtered = $filtered | Where-Object { $_.Name -like "*${nameFilter}" -or $_.EmailAddress -like "*${nameFilter}" -or $_.Title -like "*${nameFilter}"} }
        'Equals'     { $filtered = $filtered | Where-Object { $_.Name -eq $nameFilter -or $_.EmailAddress -eq $nameFilter -or $_.SamAccountName -eq $nameFilter } }
      }
    }
  }
  return $filtered
}

function Apply-VisibilityFilters {
  <#
  .SYNOPSIS
  Rebuilds the tree view with current visibility filter settings
  .DESCRIPTION
  Called when Show* checkboxes are toggled (Groups, DCs, Computers, OUs)
  Does NOT re-query data, just rebuilds the display
  #>
  Debug-Log "Applying visibility filters..." -Type "Insight"

  ## CRITICAL: Re-Initialise FilterOptions if it's wrong
  if (-not $Script:FilterOptions -or $Script:FilterOptions -isnot [hashtable]) {
    Debug-Log "CRITICAL: FilterOptions broken in Apply-VisibilityFilters! Reinitializing..." -Type "Problem"
    $Script:FilterOptions = @{
      ShowDisabledUsers       = $true
      ShowEnabledUsers        = $true
      ShowPasswordExpiring72h = $true
      ShowPasswordExpired     = $true
      ShowLockedUsers         = $true
      ShowGroups              = $true
      ShowDCs                 = $true
      ShowComputers           = $true
      ShowOUs                 = $true
      ShowUsersNoGroups       = $true
      ShowDevicesNoLAPS       = $true
      ShowDevicesNoBitlocker  = $true
      NameFilter              = ""
      NameOperator            = "Contains"
      QuickFilter             = "All"
      SortBy                  = "Name"
      SortDescending          = $false
    }
  }

  [Terminal.Gui.Application]::MainLoop.Invoke({
    try {
      Build-Tree -domain $Script:CurrentDomain
      Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
      [Terminal.Gui.Application]::Refresh()
    } catch {
      Debug-Log "Visibility filter error: $_" -Type "Problem"
    }
  })
}

## Info panel to reduce clutter. Call once to create, then again to update, e.g. if domain changes
function Show-InfoPanel {
  param(
    [Terminal.Gui.View]$Parent,
    [int]$PanelWidth  = 40,
    [int]$PanelHeight = 12,
    [switch]$UpdateOnly
  )

  ## ----------{ Create (first run) }----------
  if (-not $Script:InfoPanel) {
    $infoPanel = [Terminal.Gui.FrameView]::new("Environment Info")
    $infoPanel.Width  = $PanelWidth
    $infoPanel.Height = $PanelHeight
    $infoPanel.X = [Terminal.Gui.Pos]::AnchorEnd($PanelWidth)

    ## DYNAMIC: Place below selection panel with 1 line gap
    if ($Parent -and $Script:SelectionPanel) {
      $infoPanel.Y = [Terminal.Gui.Pos]::Bottom($Script:SelectionPanel) + 0
    } else {
      $infoPanel.Y = 30  # Fallback
    }

    $yPos = 0

    $lblForest = [Terminal.Gui.Label]::new("")
    $lblForest.X = 1; $lblForest.Y = $yPos++
    $infoPanel.Add($lblForest)

    $lblDomain = [Terminal.Gui.Label]::new("")
    $lblDomain.X = 1; $lblDomain.Y = $yPos++
    $infoPanel.Add($lblDomain)

    $lblDC = [Terminal.Gui.Label]::new("")
    $lblDC.X = 1; $lblDC.Y = $yPos++
    $infoPanel.Add($lblDC)

    $lblInfoPanelUsers = [Terminal.Gui.Label]::new("")
    $lblInfoPanelUsers.X = 1; $lblInfoPanelUsers.Y = $yPos++
    $infoPanel.Add($lblInfoPanelUsers)

    $lblInfoPanelGroups = [Terminal.Gui.Label]::new("")
    $lblInfoPanelGroups.X = 1; $lblInfoPanelGroups.Y = $yPos++
    $infoPanel.Add($lblInfoPanelGroups)

    $lblInfoPanelComputers = [Terminal.Gui.Label]::new("")
    $lblInfoPanelComputers.X = 1; $lblInfoPanelComputers.Y = $yPos++
    $infoPanel.Add($lblInfoPanelComputers)

    $lblInfoPanelDCs = [Terminal.Gui.Label]::new("")
    $lblInfoPanelDCs.X = 1; $lblInfoPanelDCs.Y = $yPos++
    $infoPanel.Add($lblInfoPanelDCs)

    $lblTotalObjects = [Terminal.Gui.Label]::new("")
    $lblTotalObjects.X = 1; $lblTotalObjects.Y = $yPos++
    $infoPanel.Add($lblTotalObjects)
    $yPos++

    $lblTheme = [Terminal.Gui.Label]::new("")
    $lblTheme.X = 1; $lblTheme.Y = $yPos++
    $infoPanel.Add($lblTheme)

    ## Store label references
    $infoPanel | Add-Member -MemberType NoteProperty -Name Tag -Value @{
      ForestLabel       = $lblForest
      DomainLabel       = $lblDomain
      DCLabel           = $lblDC
      UsersLabel        = $lblInfoPanelUsers
      GroupsLabel       = $lblInfoPanelGroups
      ComputersLabel    = $lblInfoPanelComputers
      DCsLabel          = $lblInfoPanelDCs
      TotalObjectsLabel = $lblTotalObjects
      ThemeLabel        = $lblTheme
    } -Force

    $Script:InfoPanel = $infoPanel
  }

  ## ----------{ Update (every call) }----------
  $labels = $Script:InfoPanel.Tag

  ## Get current DC name from the single source of truth
  $dcName = if ($Script:CurrentDCName) { $Script:CurrentDCName } else { "(None)" }

  $labels.ForestLabel.Text       = [NStack.ustring]::Make("Forest: $($Script:ForestName)")
  $labels.DomainLabel.Text       = [NStack.ustring]::Make("Domain: $($Script:CurrentDomain)")
  $labels.DCLabel.Text           = [NStack.ustring]::Make("Current DC: $dcName")

  ## Agressive: Force Terminal.Gui to notice the change
  $labels.DCLabel.Width = [Terminal.Gui.Dim]::Fill()
  $labels.DCLabel.Height = 1

  ## Debug: Check what we just set (convert to string properly)
  Debug-Log "dcName variable = '$dcName'" -Type "Tracing"
  Debug-Log "DCLabel.Text after setting = '$($labels.DCLabel.Text.ToString())'" -Type "Tracing"
  Debug-Log "DCLabel.Visible = $($labels.DCLabel.Visible)" -Type "Tracing"
  Debug-Log "InfoPanel.Visible = $($Script:InfoPanel.Visible)" -Type "Tracing"

  $labels.UsersLabel.Text        = [NStack.ustring]::Make("Users: $($Script:Users.Count)")
  $labels.GroupsLabel.Text       = [NStack.ustring]::Make("Groups: $($Script:Groups.Count)")
  $labels.ComputersLabel.Text    = [NStack.ustring]::Make("Computers: $($Script:Computers.Count)")
  $labels.DCsLabel.Text          = [NStack.ustring]::Make("DCs: $($Script:DCs.Count)")
  $labels.TotalObjectsLabel.Text = [NStack.ustring]::Make("Objects: $($Script:ADObjects.Count)")
  $labels.ThemeLabel.Text        = [NStack.ustring]::Make("Theme: $($Script:ThemeMode)")

  ## Force visual update - MORE AGGRESSIVE
  $labels.DCLabel.SetNeedsDisplay()
  $labels.DCLabel.Redraw($labels.DCLabel.Bounds) ## Force immediate redraw
  $Script:InfoPanel.SetNeedsDisplay()
  $Script:InfoPanel.LayoutSubviews() ## Force layout recalculation

  if ($Script:InfoPanel.SuperView) {
    $Script:InfoPanel.SuperView.SetNeedsDisplay()
    $Script:InfoPanel.SuperView.LayoutSubviews()
  }

  if ($UpdateOnly) {
    ## Small delay for UI to process
    Start-Sleep -Milliseconds 100
    [Terminal.Gui.Application]::Refresh()
    Debug-Log "InfoPanel update and refresh completed" -Type "Tracing"
  }
  return $Script:InfoPanel
}

## ----------{ Filter Panel (Add to main window )----------
function Create-FilterPanel {
  param(
    [Terminal.Gui.View]$Parent
  )

  ## Safety: Check if FilterOptions is null first, then check type
  if (-not $Script:FilterOptions) {
    Debug-Log "FilterOptions is NULL, initializing..." -Type "Warning"
    $Script:FilterOptions = @{
      ShowDisabledUsers       = $true
      ShowEnabledUsers        = $true
      ShowPasswordExpiring72h = $true
      ShowPasswordExpired     = $true
      ShowLockedUsers         = $true
      ShowGroups              = $true
      ShowDCs                 = $true
      ShowComputers           = $true
      ShowOUs                 = $true
      ShowUsersNoGroups       = $true
      ShowDevicesNoLAPS       = $true
      ShowDevicesNoBitlocker  = $true
      NameFilter              = ""
      NameOperator            = "Contains"
      QuickFilter             = "All"
      SortBy                  = "Name"
      SortDescending          = $false
    }
  }
  elseif ($Script:FilterOptions -isnot [hashtable]) {
    Debug-Log "FilterOptions was converted to $($Script:FilterOptions.GetType().Name), fixing..." -Type "Warning"
    $temp = @{}
    foreach ($key in $Script:FilterOptions.PSObject.Properties.Name) {
      $temp[$key] = $Script:FilterOptions.$key
    }
    $Script:FilterOptions = $temp
  }

  ## Capture functions at the top for clousure
  $debugLogFunc        = ${function:Debug-Log}
  $buildTreeFunc       = ${function:Build-Tree}
  $manageFilterFunc    = ${function:Manage-FilterStatusLabel}
  $applyVisFiltersFunc = ${function:Apply-VisibilityFilters}

  ## Create frame
  $filterFrame = [Terminal.Gui.FrameView]::new("Filters")
  $filterFrame.Width = 40
  $filterFrame.Height = 26
  $filterFrame.X = [Terminal.Gui.Pos]::AnchorEnd(40)
  $filterFrame.Y = 0
  $y = 0

  ## ----------{ Name Filter with Operator }----------
  $lblNameFilter = [Terminal.Gui.Label]::new("Search Name/Email/Title:")
  $lblNameFilter.X=1; $lblNameFilter.Y=$y
  $filterFrame.Add($lblNameFilter)
  $y+=2

  ## Operator dropdown
  $lblOperator = [Terminal.Gui.Label]::new("Match:")
  $lblOperator.X=1; $lblOperator.Y=$y
  $filterFrame.Add($lblOperator)

  $cmbOperator = [Terminal.Gui.ComboBox]::new()
  $cmbOperator.X=8; $cmbOperator.Y=$y; $cmbOperator.Width=13
  $cmbOperator.SetSource([string[]]@("Contains", "StartsWith", "EndsWith", "Equals"))

  ## FIX: Ensure NameOperator has a default value
  if (-not $Script:FilterOptions.NameOperator) { $Script:FilterOptions.NameOperator = "Contains" }
  $cmbOperator.Text = [NStack.ustring]::Make($Script:FilterOptions.NameOperator.ToString())
  $cmbOperator.add_SelectedItemChanged({
    $Script:FilterOptions.NameOperator = $cmbOperator.Text.ToString()
  }.GetNewClosure())
  $filterFrame.Add($cmbOperator)
  $y+=2

  ## Text field
  $txtNameFilter = [Terminal.Gui.TextField]::new($Script:FilterOptions.NameFilter)
  $txtNameFilter.X=1; $txtNameFilter.Y=$y; $txtNameFilter.Width=35
  $txtNameFilter.add_TextChanged({
    $Script:FilterOptions.NameFilter = $txtNameFilter.Text.ToString()
  }.GetNewClosure())
  $filterFrame.Add($txtNameFilter)
  $y+=1

  ## Separator line (visual spacing)
  $lblSep1 = [Terminal.Gui.Label]::new("─────────────────────────────────────")
  $lblSep1.X=1; $lblSep1.Y=$y
  $filterFrame.Add($lblSep1)
  $y+=1

  ## ----------{ Quick Filters Dropdown }----------
  $lblQuickFilter = [Terminal.Gui.Label]::new("Quick Filter:")
  $lblQuickFilter.X=1; $lblQuickFilter.Y=$y
  $filterFrame.Add($lblQuickFilter)
  $y+=1

  $cmbQuickFilter = [Terminal.Gui.ComboBox]::new()
  $cmbQuickFilter.X=1; $cmbQuickFilter.Y=$y; $cmbQuickFilter.Width=35

  $quickFilters = @(
    "All", "LockedOnly", "DisabledOnly", "EnabledOnly", "NeverLoggedIn", "NoManager", "PasswordExpired", "PasswordExpiring72h",
    "PasswordNeverExpires", "AccountExpired", "AccountExpiring30d", "StaleAccounts90d", "EmptyEmail", "EmptyDepartment" )

  ## Friendly display names
  $quickFilterDisplay = @(
    "All Users", "Locked Accounts", "Disabled Accounts", "Enabled Accounts", "Never Logged In", "No Manager Assigned",
    "Password Expired", "Password Expiring (72h)", "Password Never Expires", "Account Expired", "Account Expiring (30d)",
    "Stale Accounts (90d+)", "No Email Address", "No Department" )

  $cmbQuickFilter.SetSource($quickFilterDisplay)
  $cmbQuickFilter.SelectedItem = 0
  $cmbQuickFilter.add_SelectedItemChanged({
    if ($cmbQuickFilter.SelectedItem -ge 0 -and $cmbQuickFilter.SelectedItem -lt $quickFilters.Count) {
      $Script:FilterOptions.QuickFilter = $quickFilters[$cmbQuickFilter.SelectedItem]
      & $debugLogFunc "Quick filter changed to: $($Script:FilterOptions.QuickFilter)" -Type "Insight"
    }
  }.GetNewClosure())
  $filterFrame.Add($cmbQuickFilter)
  $y+=1

  ## Separator line (visual spacing)
  $lblSep2 = [Terminal.Gui.Label]::new("─────────────────────────────────────")
  $lblSep2.X=1; $lblSep2.Y=$y
  $filterFrame.Add($lblSep2)
  $y+=1

  ## ----------{ Show/Hide Checkboxes }----------
  $chkEnabled = [Terminal.Gui.CheckBox]::new("Enabled Users")
  $chkEnabled.X=1; $chkEnabled.Y=$y; $chkEnabled.Checked=$Script:FilterOptions.ShowEnabledUsers
  $chkEnabled.add_Toggled({
    $Script:FilterOptions.ShowEnabledUsers = [bool]$chkEnabled.Checked
  })
  $filterFrame.Add($chkEnabled)
  $y+=1

  $chkLocked = [Terminal.Gui.CheckBox]::new("Locked Users")
  $chkLocked.X=1; $chkLocked.Y=$y; $chkLocked.Checked=$Script:FilterOptions.ShowLockedUsers
  $chkLocked.add_Toggled({
    $Script:FilterOptions.ShowLockedUsers = [bool]$chkLocked.Checked
  })
  $filterFrame.Add($chkLocked)
  $y+=1

  $chkDisabled = [Terminal.Gui.CheckBox]::new("Disabled Users")
  $chkDisabled.X=1; $chkDisabled.Y=$y; $chkDisabled.Checked=$Script:FilterOptions.ShowDisabledUsers
  $chkDisabled.add_Toggled({
    $Script:FilterOptions.ShowDisabledUsers = [bool]$chkDisabled.Checked
  })
  $filterFrame.Add($chkDisabled)
  $y+=1

  $chkPwdExpiring72h = [Terminal.Gui.CheckBox]::new("Passwords Expiring in next 72h")
  $chkPwdExpiring72h.X = 1
  $chkPwdExpiring72h.Y = $y
  $chkPwdExpiring72h.Checked = $Script:FilterOptions.ShowPasswordExpiring72h
  $chkPwdExpiring72h.add_Toggled({
    $Script:FilterOptions.ShowPasswordExpiring72h = [bool]$chkPwdExpiring72h.Checked
  })
  $filterFrame.Add($chkPwdExpiring72h)
  $y += 1

  $chkPwdExpired = [Terminal.Gui.CheckBox]::new("Users with expired passwords")
  $chkPwdExpired.X = 1
  $chkPwdExpired.Y = $y
  $chkPwdExpired.Checked = $Script:FilterOptions.ShowPasswordExpired
  $chkPwdExpired.add_Toggled({
    $Script:FilterOptions.ShowPasswordExpired = [bool]$chkPwdExpired.Checked
  })
  $filterFrame.Add($chkPwdExpired)
  $y += 1

  $chkNoGroups = [Terminal.Gui.CheckBox]::new("Users Not in Any Groups")
  $chkNoGroups.X = 1
  $chkNoGroups.Y = $y
  $chkNoGroups.Checked = $Script:FilterOptions.ShowUsersNoGroups
  $chkNoGroups.add_Toggled({
    $Script:FilterOptions.ShowUsersNoGroups = [bool]$chkNoGroups.Checked
  })
  $filterFrame.Add($chkNoGroups)
  $y += 1

  ## ----------{ Visibility filter checkboxes (Groups, OUs, DCs, Computers) }----------
  $chkGroups = [Terminal.Gui.CheckBox]::new("Groups")
  $chkGroups.X=1; $chkGroups.Y=$y; $chkGroups.Checked=$Script:FilterOptions.ShowGroups
  $chkGroups.add_Toggled({
    $Script:FilterOptions.ShowGroups = [bool]$chkGroups.Checked
    Apply-VisibilityFilters
  })
  $filterFrame.Add($chkGroups)
  $y+=1

  $chkOUs = [Terminal.Gui.CheckBox]::new("OUs")
  $chkOUs.X=1; $chkOUs.Y=$y; $chkOUs.Checked=$Script:FilterOptions.ShowOUs
  $chkOUs.add_Toggled({
    $Script:FilterOptions.ShowOUs = [bool]$chkOUs.Checked
    Apply-VisibilityFilters
  })
  $filterFrame.Add($chkOUs)
  $y+=1

  $chkDCs = [Terminal.Gui.CheckBox]::new("Domain Controllers")
  $chkDCs.X=1; $chkDCs.Y=$y; $chkDCs.Checked=$Script:FilterOptions.ShowDCs
  $chkDCs.add_Toggled({
    $Script:FilterOptions.ShowDCs = [bool]$chkDCs.Checked
    Apply-VisibilityFilters
  })
  $filterFrame.Add($chkDCs)
  $y+=1

  $chkComputers = [Terminal.Gui.CheckBox]::new("Computers")
  $chkComputers.X=1; $chkComputers.Y=$y; $chkComputers.Checked=$Script:FilterOptions.ShowComputers
  $chkComputers.add_Toggled({
    $Script:FilterOptions.ShowComputers = [bool]$chkComputers.Checked
    Apply-VisibilityFilters
  })
  $filterFrame.Add($chkComputers)
  $y+=1

  $chkNoLAPS = [Terminal.Gui.CheckBox]::new("Devices Without LAPS")
  $chkNoLAPS.X = 1
  $chkNoLAPS.Y = $y
  $chkNoLAPS.Checked = $Script:FilterOptions.ShowDevicesNoLAPS
  $chkNoLAPS.add_Toggled({
    $Script:FilterOptions.ShowDevicesNoLAPS = [bool]$chkNoLAPS.Checked
  })
  $filterFrame.Add($chkNoLAPS)
  $y += 1

  $chkNoBitlocker = [Terminal.Gui.CheckBox]::new("Devices Without BitLocker Keys")
  $chkNoBitlocker.X = 1
  $chkNoBitlocker.Y = $y
  $chkNoBitlocker.Checked = $Script:FilterOptions.ShowDevicesNoBitlocker
  $chkNoBitlocker.add_Toggled({
    $Script:FilterOptions.ShowDevicesNoBitlocker = [bool]$chkNoBitlocker.Checked
  })
  $filterFrame.Add($chkNoBitlocker)
  $y += 2

  ## Separator line (visual spacing)
  $lblSep3 = [Terminal.Gui.Label]::new("─────────────────────────────────────")
  $lblSep3.X=1; $lblSep3.Y=$y
  $filterFrame.Add($lblSep3)
  $y+=1

  ## ----------{ Apply/Reset Buttons }----------
  $btnApplyFilter = [Terminal.Gui.Button]::new("Apply Filter")
  $btnApplyFilter.X=1; $btnApplyFilter.Y=$y
  $btnApplyFilter.add_Clicked({
    & $debugLogFunc "Applying filters..." -Type "Insight"

    ## Rebuild tree with filters
    $rootNode = & $buildTreeFunc -domain $Script:CurrentDomain
    if ($rootNode) {
      [Terminal.Gui.Application]::Refresh()
    }

    ## Update status label
    & $manageFilterFunc -Action 'Update' -Label $Script:FilterStatusLabel
  }.GetNewClosure())
  $filterFrame.Add($btnApplyFilter)

  $btnResetFilter = [Terminal.Gui.Button]::new("Reset")
  $btnResetFilter.X=21; $btnResetFilter.Y=$y
  $btnResetFilter.add_Clicked({
    & $debugLogFunc "Resetting filters..." -Type "Insight"

    ## Reset all filters
    $Script:FilterOptions.ShowDisabledUsers       = $true
    $Script:FilterOptions.ShowEnabledUsers        = $true
    $Script:FilterOptions.ShowLockedUsers         = $true
    $Script:FilterOptions.ShowGroups              = $true
    $Script:FilterOptions.ShowDCs                 = $true
    $Script:FilterOptions.ShowComputers           = $true
    $Script:FilterOptions.ShowOUs                 = $true
    $Script:FilterOptions.ShowDevicesNoBitlocker  = $true
    $Script:FilterOptions.ShowDevicesNoLAPS       = $true
    $Script:FilterOptions.ShowUsersNoGroups       = $true
    $Script:FilterOptions.ShowPasswordExpiring72h = $true
    $Script:FilterOptions.ShowPasswordExpired     = $true
    $Script:FilterOptions.NameFilter              = ""
    $Script:FilterOptions.NameOperator            = "Contains"
    $Script:FilterOptions.QuickFilter             = "All"
    $Script:FilterOptions.SortBy                  = "Name"
    $Script:FilterOptions.SortDescending          = $false

    ## Reset UI controls
    $chkEnabled.Checked          = $true
    $chkDisabled.Checked         = $true
    $chkGroups.Checked           = $true
    $chkDCs.Checked              = $true
    $chkComputers.Checked        = $true
    $chkOUs.Checked              = $true
    $chkLocked.Checked           = $true
    $chkNoLAPS.Checked           = $true
    $chkNoBitlocker.Checked      = $true
    $chkNoGroups.Checked         = $true
    $chkPwdExpiring72h.Checked   = $true
    $chkPwdExpired.Checked       = $true
    $txtNameFilter.Text          = ""
    $cmbOperator.SelectedItem    = 0
    $cmbQuickFilter.SelectedItem = 0

    ## Rebuild tree
    $rootNode = & $buildTreeFunc -domain $Script:CurrentDomain
    if ($rootNode) { [Terminal.Gui.Application]::Refresh() }
    ## Update status label
    & $manageFilterFunc -Action 'Update' -Label $Script:FilterStatusLabel
  }.GetNewClosure())
  $filterFrame.Add($btnResetFilter)
  $y+=2

  ## ----------{ Filter Status Label }----------
  $Script:FilterStatusLabel = Manage-FilterStatusLabel -Action 'Create' -X 1 -Y $y -InPanel
  $filterFrame.Add($Script:FilterStatusLabel)

  & $debugLogFunc "Enhanced FilterPanel created with LDAP-style filters" -Type "Insight"

  ## Apply theme
  if ($Script:themeData -and $Script:themeData.MainWindow) { $filterFrame.ColorScheme = $Script:themeData.MainWindow }

  ## Store reference
  $Script:FilterPanel = $filterFrame

  return $filterFrame
}

## ----------{ Quick Filter Menu (for Menu Bar) }----------
function Show-QuickFilterDialog {
  <#
  .SYNOPSIS
  Simplified quick filter dialog that sets filter and rebuilds tree
  #>

  $dlg = [Terminal.Gui.Dialog]::new("Quick Filters", 50, 25)
  $y = 1

  $lbl = [Terminal.Gui.Label]::new("Select a quick filter:")
  $lbl.X=2; $lbl.Y=$y
  $dlg.Add($lbl)
  $y+=2

  ## Filter definitions
  $quickFilters = @(
    @{ Name = "👥 All Users"               ; Type = "All" }
    @{ Name = "🔒 Locked Accounts"         ; Type = "LockedOnly" }
    @{ Name = "⛔ Disabled Accounts"       ; Type = "DisabledOnly" }
    @{ Name = "✓  Enabled Accounts"        ; Type = "EnabledOnly" }
    @{ Name = "🆕 Never Logged In"         ; Type = "NeverLoggedIn" }
    @{ Name = "⊗  No Manager Assigned"     ; Type = "NoManager" }
    @{ Name = "🔑 Password Expired"        ; Type = "PasswordExpired" }
    @{ Name = "⏰ Password Expiring (72h)" ; Type = "PasswordExpiring72h" }
    @{ Name = "🔓 Password Never Expires"  ; Type = "PasswordNeverExpires" }
    @{ Name = "☠  Account Expired"         ; Type = "AccountExpired" }
    @{ Name = "⏰ Account Expiring (30d)"  ; Type = "AccountExpiring30d" }
    @{ Name = "💤 Stale Accounts (90d+)"   ; Type = "StaleAccounts90d" }
    @{ Name = "📧 No Email Address"        ; Type = "EmptyEmail" }
    @{ Name = "🏢 No Department"           ; Type = "EmptyDepartment" }
    @{ Name = "⚠  Empty Groups"            ; Type = "EmptyGroups" }
  )
  $filterNames = $quickFilters | ForEach-Object { $_.Name }

  $lstFilters = [Terminal.Gui.ListView]::new($filterNames)
  $lstFilters.X=2; $lstFilters.Y=$y; $lstFilters.Width=44; $lstFilters.Height=15
  $dlg.Add($lstFilters)

  $btnApply = [Terminal.Gui.Button]::new("Apply")
  $btnApply.add_Clicked({
    if ($lstFilters.SelectedItem -ge 0) {
      $selected = $quickFilters[$lstFilters.SelectedItem]
      Debug-Log "Applying quick filter: $($selected.Name)" -Type "Insight"

      ## Set the filter
      $Script:FilterOptions.QuickFilter = $selected.Type
      ## Special handling for group filter
      if ($selected.Type -eq 'EmptyGroups') {
        $Script:FilterOptions.ShowGroups = $true
        $Script:FilterOptions.ShowEnabledUsers = $false
        $Script:FilterOptions.ShowDisabledUsers = $false
      }

      ## Rebuild tree
      $rootNode = Build-Tree -domain $Script:CurrentDomain
      if ($rootNode) {
        $Script:tree.ClearObjects()
        $Script:tree.AddObject($rootNode)
      }

      ## Update status
      Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
      [Terminal.Gui.Application]::RequestStop()
    }
  }).GetNewClosure()
  $dlg.AddButton($btnApply)
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
  $dlg.AddButton($btnCancel)
  [Terminal.Gui.Application]::Run($dlg)
}

## ----------{ Main Get AD health function }----------
function Get-TrustStatusText {
  param([string]$Domain)

  $output = @()
  $output += "TRUST RELATIONSHIPS FOR $Domain"
  $output += "=" * 60
  $output += ""

  try {
    if ($Script:DemoMode) {
      $output += "Demo Mode - Simulated Trust Information"
      $output += ""
      $output += "Trust Partner: example.com"
      $output += "  Direction: Bidirectional"
      $output += "  Type: Forest"
      $output += "  Status: ✓ Active"
      $output += ""
      $output += "Trust Partner: fabrikam.com"
      $output += "  Direction: Outbound"
      $output += "  Type: External"
      $output += "  Status: ✓ Active"
      $output += ""
      $output += "To test trust in production, use 'Test Trust' button"
    } else {
      $trusts = Get-ADTrust -Filter * -Server $Domain -ErrorAction Stop

      if ($trusts) {
        foreach ($trust in $trusts) {
          $output += "Trust Partner: $($trust.Target)"
          $output += "  Direction: $($trust.Direction)"
          $output += "  Type: $($trust.TrustType)"
          $output += "  Status: $(if ($trust.TrustAttributes -band 0x20) { '✓ Active' } else { '⚠ Inactive' })"
          $output += ""
        }
      } else {
        $output += "No trust relationships found"
      }
    }
  } catch {
    $output += "⚠ Error retrieving trust information:"
    $output += $_.Exception.Message
  }
  return $output -join "`n"
}

## ----------{ Content generators }----------
function Get-SystemInfoText {
  $output = @()
  $output += ""

  ## Determine current OS
  if ($IsWindows)     { $currentOS = 'Windows'
  } elseif ($IsLinux) { $currentOS = 'Linux'
  } elseif ($IsMacOS) { $currentOS = 'macOS'
  } else { $currentOS = 'Unknown' }

  ## Add to output
  $output += "Operating System: $currentOS"

  ## PowerShell Version
  $output += "PowerShell: Version: $($PSVersionTable.PSVersion) Edition: $($PSVersionTable.PSEdition)"
  $output += ""

  ## Module Availability
  $output += "PowerShell Modules:"
  $output += "  ActiveDirectory:                      $(if ($Script:HasActiveDirectory) { '✓ Available' } else { '✗ Not Found' })"
  $output += "  DNSClient:                            $(if ($Script:hasDNSClient) { '✓ Available' } else { '✗ Not Found' })"
  $output += "  NerdFonts:                            $(if ($Script:hasNerdFonts) { '✓ Available' } else { '✗ Not Found' })"
  $output += "  PSWriteColor:                         $(if ($Script:hasPSWriteColor) { '✓ Available' } else { '✗ Not Found' })"
  $output += "  Microsoft.PowerShell.ConsoleGuiTools: $(if ($Script:hasConsoleTools) { '✓ Available' } else { '✗ Not Found' })"
  $output += "  Terminal-Icons:                       $(if ($hasTerminalIcons) { '✓ Available' } else { '✗ Not Found' })"
  $output += ""

  ## Installation Instructions
  if (-not $Script:HasActiveDirectory) {
    $output += " ⚠ ActiveDirectory Module Not Found"
    $output += ""
    if ($Script:ADHealthOS.IsWindows) {
      $output += " To install on Windows:"
      $output += "   1. Open PowerShell as Administrator"
      $output += "   2. Run: Add-WindowsFeature RSAT-AD-PowerShell Or: Install-WindowsFeature RSAT-AD-PowerShell"
    } elseif ($Script:ADHealthOS.IsLinux) {
      $output += " To install on Linux:"
      $output += "   1. Install realmd: sudo apt install realmd sssd sssd-tools adcli"
      $output += "   2. Join domain: sudo realm join domain.com"
    } elseif ($Script:ADHealthOS.IsMacOS) {
      $output += " To install on macOS:"
      $output += "   1. Install via Homebrew: brew install realmd"
      $output += "   2. Configure AD integration"
    }
    $output += ""
  }

  ## ----------{ Group Policy Tools Detection }----------
  if ($IsWindows) {
    ## Default feature state
    $gpmcFeatureInstalled = $false
    ## Get-WindowsFeature only exists on Windows Server
    if (Get-Command -Name Get-WindowsFeature -ErrorAction SilentlyContinue) {
      try { $gpmcFeatureInstalled = (Get-WindowsFeature -Name GPMC -ErrorAction Stop).InstallState -eq 'Installed' }
      catch { $gpmcFeatureInstalled = $false }
    }
    ## Combined result if you need a single flag
    $gpmcInstalled = $gpmcModuleInstalled -or $gpmcFeatureInstalled
    if ($gpmcInstalled) {
      Debug-Log "Group Policy Management tools detected." -Type "Success"
    } else {
      ## Detect OS edition: Windows 10/11 client vs Windows Server
      $osCaption = (Get-CimInstance Win32_OperatingSystem).Caption
      if ($osCaption -match 'Windows 10|Windows 11') {
        $output += " ⚠ Group Policy tools (RSAT) not installed on this Windows client. To install on Windows:"
        $output += "   1.) Install with: Add-WindowsCapability -Online -Name Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0 or"
        $output += "   2.) Using DISM: DISM /Online /Add-Capability /CapabilityName:Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0"
        $output += ""
      } elseif ($osCaption -match 'Windows Server') {
        $output += " ⚠ Group Policy Management Console (GPMC) not installed on this Windows Server. To install:"
        $output += "   1.) Install with: Install-WindowsFeature -Name GPMC"
        $output += ""
      } else {
        Debug-Log "Unknown Windows edition. Cannot suggest Group Policy installation." -Type "Warning"
      }
    }
  } else {
    Debug-Log "Group Policy tools are Windows-only. Skipping check for $($PSVersionTable.OS)." -Type "Insight"
  }
  ## Tool Availability
  $output += " External Tools:"
  foreach ($tool in $Script:ADHealthTools.Keys | Sort-Object) {
    $status = if ($Script:ADHealthTools[$tool]) { "✓ Available" } else { "✗ Not Found" }
    $output += "  {0,-20} {1}" -f "${tool}:", $status
  }
  $output += ""
  ## Demo Mode
  $output += "Application Mode:"
  $output += "Demo Mode: $($Script:DemoMode) using Domain: $($Script:ADHealthDomain)"
  $output += ""
  ## Data Summary (from Script variables)
  if ($Script:DemoMode) {
    $output += "Demo Data Summary:  Users: $(if ($Script:rawUsers) { $Script:rawUsers.Count } else { 0 })  Groups:  $(if ($Script:rawDemoGroups) { $Script:rawDemoGroups.Count } else { 0 }) DCs:  $(if ($Script:rawDCs) { $Script:rawDCs.Count } elseif ($Script:DCs) { $Script:DCs.Count } else { 0 })  Computers:  $(if ($Script:rawComputers) { $Script:rawComputers.Count } else { 0 })"
    $output += ""
  }
  return ($output -join "`n")
}

function Get-TrustTestingTabContent {
  param([string]$Domain)

  ## Get trust relationships
  $trusts = if ($Script:DemoMode) {
    @(
      @{ Name = "corp.example.com"; Direction = "Bidirectional"; Type = "Forest" }
      @{ Name = "partner.example.org"; Direction = "Outbound"; Type = "External" }
    )
  } else {
    try {
      if (Get-Command Get-ADTrust -ErrorAction SilentlyContinue) {
        Get-ADTrust -Filter * | ForEach-Object {
          @{
            Name = $_.Name
            Direction = $_.Direction
            Type = $_.TrustType
          }
        }
      } else {
        @()
      }
    } catch {
      @()
    }
  }

  if ($trusts.Count -eq 0) {
    return @"
Trust Testing
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

No trust relationships configured.

This domain does not have any external trust
relationships to test.
"@
  }

  $text = @"
Trust Testing
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Available Trust Relationships:

"@

  foreach ($trust in $trusts) {
    $text += "  • $($trust.Name) ($($trust.Direction), $($trust.Type))`n"
  }

  $text += @"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

INTERACTIVE TRUST TESTING

To test a trust relationship, use the dedicated
Trust Testing dialog from the main menu:

  Actions → Test Trust Relationships

Or use command-line tools:

  nltest /server:$env:COMPUTERNAME /sc_query:DOMAIN

  Test-ComputerSecureChannel -Credential (Get-Credential) -Server DOMAIN

  netdom trust $Domain /domain:DOMAIN /verify

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

TRUST HEALTH INDICATORS:

✓ Trust should respond to secure channel queries
✓ Authentication should succeed bidirectionally
✓ DNS should resolve partner domain
✓ Time sync should be within 5 minutes
✓ Firewall should allow RPC traffic (135, 49152-65535)

"@

  return $text
}

## ADD THIS NEW HELPER FUNCTION for Statistics tab:
function Get-DomainStatisticsText {
  param([string]$Domain)

  $totalUsers = $Script:Users.Count
  $enabledUsers = ($Script:Users | Where-Object { $_.Enabled -eq $true }).Count
  $disabledUsers = ($Script:Users | Where-Object { $_.Disabled -eq $true }).Count
  $lockedUsers = ($Script:Users | Where-Object { $_.LockedOut -eq $true }).Count

  $totalGroups = $Script:Groups.Count
  $securityGroups = ($Script:Groups | Where-Object { $_.GroupCategory -eq 'Security' }).Count

  $totalComputers = $Script:Computers.Count
  $enabledComputers = ($Script:Computers | Where-Object { $_.Enabled -eq $true }).Count

  $totalDCs = $Script:DCs.Count
  $totalObjects = $totalUsers + $totalGroups + $totalComputers + $totalDCs

  return @"
Domain Statistics
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Domain: $Domain
Forest: $($Script:ForestName)
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

USERS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Total Users:          $totalUsers
  Enabled:            $enabledUsers
  Disabled:           $disabledUsers
  Locked Out:         $lockedUsers

GROUPS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Total Groups:         $totalGroups
  Security Groups:    $securityGroups
  Distribution:       $($totalGroups - $securityGroups)

COMPUTERS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Total Computers:      $totalComputers
  Enabled:            $enabledComputers
  Disabled:           $($totalComputers - $enabledComputers)

DOMAIN CONTROLLERS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Total DCs:            $totalDCs

FOREST INFORMATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Forest Name:          $($Script:ForestName)
Total Domains:        $($Script:Domains.Count)
Total Sites:          $($Script:Sites.Count)

SUMMARY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TOTAL OBJECTS:        $totalObjects

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"@
}

function Get-DCStatusText {
  param([string]$Domain)

  $output = @()
  if ($Script:DemoMode) {
    ## Use Script:DCs or Script:rawDCs
    $dcs = if ($Script:DCs) { $Script:DCs } elseif ($Script:rawDCs) { $Script:rawDCs } else { @() }

    if ($dcs.Count -gt 0) {
      $fmt = "{0,-20} {1,-20} {2,-16} {3,-10} {4,-6} {5,-6} {6,-6} {7,-6}"
      $output += $fmt -f "DC Name", "Site", "IP Address", "Status", "LDAP", "Kerb", "SMB", "DNS"
      $output += $fmt -f ("─" * 20), ("─" * 20), ("─" * 16), ("─" * 10), ("─" * 6), ("─" * 6), ("─" * 6), ("─" * 6)
      $output += ""

      $onlineCount = 0
      foreach ($dc in $dcs) {
        $dcName = $dc.Name
        $site = if ($dc.Site) { $dc.Site } else { "Unknown" }
        $ip = if ($dc.IPv4Address) { $dc.IPv4Address } elseif ($dc.IP) { $dc.IP -join "," } else { "N/A" }
        $status = if ($dc.Enabled) { $onlineCount++; "✓ Online" } else { "✗ Offline" }
        $ldap = "✓ OK"
        $kerb = "✓ OK"
        $smb  = "✓ OK"
        $dns  = "✓ OK"
        $output += $fmt -f $dcName, $site, $ip, $status, $ldap, $kerb, $smb, $dns
      }

      $output += ""
      $output += "Summary:"
      $output += "  Total DCs:     $($dcs.Count)"
      $output += "  Online:        $onlineCount"
      $output += "  Offline:       $($dcs.Count - $onlineCount)"
      $output += "  Health:        $(if ($onlineCount -eq $dcs.Count) { '✓ HEALTHY' } elseif ($onlineCount -gt 0) { '⚠ DEGRADED' } else { '✗ CRITICAL' })"
      $output += ""
      $output += "Ports: LDAP (389), Kerberos (88), SMB (445), DNS (53)"
    } else {
      $output += "No domain controllers found in demo data."
    }
  } else {
    ## Production Mode - check ports
    $result = Test-DCStatus -Domain $Domain

    if ($result.Summary) {
      $fmt     = "{0,-20} {1,-20} {2,-16} {3,-10} {4,-6} {5,-6} {6,-6} {7,-6}"
      $output += $fmt -f "DC Name", "Site", "IP Address", "Status", "LDAP", "Kerb", "SMB", "DNS"
      $output += $fmt -f ("─" * 20), ("─" * 20), ("─" * 16), ("─" * 10), ("─" * 6), ("─" * 6), ("─" * 6), ("─" * 6)
      $output += ""

      foreach ($dc in $result.Summary) {
        $status = if ($dc.Reachable -eq "OK") { "✓ Online" } else { "✗ Offline" }
        ## Port checks
        $ldap = if ($dc.LDAP) { "✓ OK" } else { "✗ DOWN" }
        $kerb = if ($dc.Kerberos) { "✓ OK" } else { "✗ DOWN" }
        $smb = if ($dc.SMB) { "✓ OK" } else { "✗ DOWN" }
        $dns = if ($dc.DNS -eq "Running") { "✓ OK" } else { "✗ DOWN" }
        $output += $fmt -f $dc.Name, $dc.Site, $dc.IP, $status, $ldap, $kerb, $smb, $dns
      }

      $onlineCount = ($result.Summary | Where-Object { $_.Reachable -eq "OK" }).Count
      $totalCount  = $result.Summary.Count

      $output += ""
      $output += "Summary:"
      $output += "  Total DCs: ($totalCount)  Online:        ($onlineCount)  Offline:       ($($totalCount - $onlineCount))  Health:        $($result.Health)"
      $output += "  Ports: LDAP (389), Kerberos (88), SMB (445), DNS (53)"
    }

    if ($result.Details) {
      $output += ""
      $output += "Details:"
      $output += $result.Details
    }
  }
  return ($output -join "`n")
}

function Get-ReplicationStatusText {
  param([string]$Domain)

  $output = @()
  $output += ""
  $hasRepadmin = $Script:ADHealthTools['repadmin.exe']
  $output += "Tool Status: repadmin.exe $(if ($hasRepadmin) { '✓ Available' } else { '✗ Not Found' })"
  $output += ""

  if ($Script:DemoMode) {
    $output += "Demo Mode - Simulated replication status for domain: $Domain"
    $output += ""

    $dcs = if ($Script:DCs) { $Script:DCs } elseif ($Script:rawDCs) { $Script:rawDCs } else { @() }
    if ($dcs.Count -gt 1) {
      $fmt = "{0,-20} {1,-20} {2,-10} {3}"
      $output += $fmt -f "Source DC", "Target DC", "Status", "Last Sync"
      $output += $fmt -f ("─" * 20), ("─" * 20), ("─" * 10), ("─" * 20)
      foreach ($sourceDC in $dcs) {
        foreach ($targetDC in $dcs) {
          if ($sourceDC.Name -ne $targetDC.Name) { $output += $fmt -f $sourceDC.Name, $targetDC.Name, "✓ OK", (Get-Date -Format "yyyy-MM-dd HH:mm") }
        }
      }

      $output += ""
      $output += "All replication partners healthy."
      $output += "No errors detected."
    } else {
      $output += "Single DC - no replication partners."
    }
  } else {
    ## Production Mode - actually run repadmin
    $result = Test-ADReplication -Domain $Domain
    if ($hasRepadmin) {
      $output += "Running repadmin /replsummary..."
      $output += ""
    }
    if ($result.Summary) {
      foreach ($line in $result.Summary) {
        $output += $line
      }
    }
    $output += ""
    $output += "Health: $($result.Health)"
    if ($result.Details) {
      $output += ""
      $output += "Details:"
      $output += $result.Details
    }
  }
  return ($output -join "`n")
}

function Get-DNSStatusText {
  param([string]$Domain)

  $output = @()
  $output += ""

  $hasNslookup = $Script:ADHealthTools['nslookup.exe']
  $output += "Tool Status: nslookup.exe $(if ($hasNslookup) { '✓ Available' } else { '✗ Not Found' })"
  $output += ""

  if ($Script:DemoMode) {
    $output += "Demo Mode - Simulated DNS status for domain: $Domain"
    $output += ""

    $dcs = if ($Script:DCs) { $Script:DCs } elseif ($Script:rawDCs) { $Script:rawDCs } else { @() }

    $fmt = "{0,-45} {1}"
    $output += $fmt -f "SRV Record", "Count"
    $output += $fmt -f ("─" * 45), ("─" * 10)
    $output += $fmt -f "_ldap._tcp.dc._msdcs.$Domain", "$($dcs.Count)"
    $output += $fmt -f "_kerberos._tcp.$Domain", "$($dcs.Count)"
    $output += $fmt -f "_ldap._tcp.$Domain", "$($dcs.Count)"
    $output += ""

    if ($dcs.Count -gt 0) {
      $output += "DC DNS Service Status:"
      foreach ($dc in $dcs) { $output += "  $($dc.Name): ✓ Running" }
    }

    $output += ""
    $output += "All DNS records appear healthy."
  } else {
    ## Production Mode - actually query DNS
    $result = Test-ADDnsRecords -Domain $Domain
    $output += "Querying DNS records for domain: $Domain"
    $output += ""

    if ($result.Summary) { foreach ($line in $result.Summary) { $output += $line } }

    $output += ""
    $output += "Health: $($result.Health)"

    if ($result.Details) {
      $output += ""
      $output += "Details:"
      $output += $result.Details
    }
  }
  return ($output -join "`n")
}

function Get-SysvolStatusText {
  param([string]$Domain)

  $output = @()
  $output += ""
  $hasDfsrdiag = $Script:ADHealthTools['dfsrdiag.exe']
  $output += "Tool Status: dfsrdiag.exe $(if ($hasDfsrdiag) { '✓ Available' } else { '✗ Not Found' })"
  $output += ""

  if ($Script:DemoMode) {
    $output += "Demo Mode - Simulated SYSVOL status for domain: $Domain"
    $output += ""
    $dcs = if ($Script:DCs) { $Script:DCs } elseif ($Script:rawDCs) { $Script:rawDCs } else { @() }
    if ($dcs.Count -gt 0) {
      $fmt = "{0,-25} {1,-20} {2}"
      $output += $fmt -f "DC Name", "Share", "Status"
      $output += $fmt -f ("─" * 25), ("─" * 20), ("─" * 20)
      foreach ($dc in $dcs) {
        $output += $fmt -f $dc.Name, "SYSVOL", "✓ Available"
        $output += $fmt -f $dc.Name, "NETLOGON", "✓ Available"
      }
      $output += ""
      $output += "All shares accessible."
    } else {
      $output += "No domain controllers found."
    }
  } else {
    ## Production Mode - actually check shares
    $result = Test-SysvolHealth -Domain $Domain
    $output += "Checking SYSVOL/NETLOGON shares for domain: $Domain"
    $output += ""
    if ($result.Summary) { foreach ($line in $result.Summary) { $output += $line } }
    $output += ""
    $output += "Health: $($result.Health)"
    if ($result.Details) {
      $output += ""
      $output += "Details:"
      $output += $result.Details
    }
  }
  return ($output -join "`n")
}

## ----------{ DFS Tab for AD Health Check }----------
function Get-DFSStatusText{

  param([string]$Domain)

  $dfsTab = @{
    Name = "DFS"
    ContentGenerator = {
    param($Domain)

    $output = @()
    $output += ""
    $output += "DFS (Distributed File System) Status"
    $output += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    $output += ""

    ## Check for imported DFS data
    $dfsLinks = @()
    $dfsrObjects = @()

    if ($Script:rawDFSLinks) {
      $dfsLinks = $Script:rawDFSLinks | Where-Object { $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')") }
      if ($dfsLinks.Count -eq 0) { $dfsLinks = $Script:rawDFSLinks }
    }
    if ($Script:rawDFSR) {
      $dfsrObjects = $Script:rawDFSR | Where-Object { $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')")
      }
      if ($dfsrObjects.Count -eq 0) { $dfsrObjects = $Script:rawDFSR }
    }
    ## DFS Namespaces
    $output += "DFS Namespaces:"
    $output += ""
    if ($dfsLinks.Count -gt 0) {
      $output += "  Found $($dfsLinks.Count) DFS link(s):"
      $output += ""
      foreach ($link in $dfsLinks) {
        $linkName = $link.Name ?? "(Unnamed)"
        $linkType = $link.Type ?? "Unknown"
        $output += "  • $linkName [$linkType]"
        if ($link.LinkPath) { $output += "    Path: $($link.LinkPath)" }
        if ($link.TargetList) {
          $targets = $link.TargetList -split ';' | Select-Object -First 3
          $output += "    Targets:"
          foreach ($target in $targets) { if ($target) { $output += "      - $target" } }
          if (($link.TargetList -split ';').Count -gt 3) {
            $remaining = ($link.TargetList -split ';').Count - 3
            $output += "      ... and $remaining more"
          }
        }
        $output += ""
      }
      $output += "  ✓ DFS Namespaces configured"
    } else {
      $output += "  (No DFS namespaces found)"
      $output += ""
      $output += "  DFS allows users to access shared folders using"
      $output += "  a unified namespace path regardless of physical location."
    }
    $output += ""

    ## DFSR (DFS Replication)
    $output += "DFS Replication (DFSR):"
    $output += ""
    if ($dfsrObjects.Count -gt 0) {
      $output += "  Found $($dfsrObjects.Count) DFSR object(s):"
      $output += ""
      # Group by class
      $groupedDFSR = $dfsrObjects | Group-Object Class
      foreach ($group in $groupedDFSR) { $output += "  • $($group.Name): $($group.Count) object(s)" }
      $output += ""
      $output += "  ✓ DFSR configured for replication"
      $output += ""
      $output += "  DFSR objects replicate folders between servers"
      $output += "  for redundancy and load balancing."
    } else {
      $output += "  (No DFSR objects found)"
      $output += ""
      $output += "  DFSR provides multi-master replication of"
      $output += "  shared folders between file servers."
    }
    $output += ""

    ## Health check commands
    $output += "DFS Health Check Commands:"
    $output += ""
    $output += "  Windows PowerShell:"
    $output += "    Get-DfsnRoot                    # List DFS roots"
    $output += "    Get-DfsnFolder                  # List DFS folders"
    $output += "    Get-DfsnFolderTarget            # List folder targets"
    $output += "    Get-DfsrBacklog                 # Check replication backlog"
    $output += "    Get-DfsReplicationGroup         # List replication groups"
    $output += ""
    $output += "  Command Line:"
    $output += "    dfsutil /root:\\domain\namespace /display"
    $output += "    dfsrdiag replicationstate       # Check DFSR state"
    $output += "    dfsrdiag backlog                # Check backlog"
    $output += ""
    $output += "  GUI Tools:"
    $output += "    dfsmgmt.msc                     # DFS Management Console"
    $output += ""

    ## Summary
    $output += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    $output += ""

    if ($dfsLinks.Count -gt 0 -or $dfsrObjects.Count -gt 0) {
      $output += "Status: ✓ DFS configured and operational"
    } else {
      $output += "Status: ⓘ DFS not configured (optional service)"
    }
    return ($output -join "`n")
  }
}

function Get-FSMOStatusText {
  param([string]$Domain)

  $output = @()
  $output += ""

  if ($Script:DemoMode) {
    $output += "Demo Mode - Simulated FSMO roles for domain: $Domain"
    $output += ""

    $dcs = if ($Script:DCs) { $Script:DCs } elseif ($Script:rawDCs) { $Script:rawDCs } else { @() }

    if ($dcs.Count -gt 0) {
      $fmt = "{0,-30} {1,-30} {2}"
      $output += $fmt -f "Role", "Holder", "Status"
      $output += $fmt -f ("─" * 30), ("─" * 30), ("─" * 15)

      $primaryDC = $dcs[0].Name
      $output += $fmt -f "Schema Master", $primaryDC, "✓ Online"
      $output += $fmt -f "Domain Naming Master", $primaryDC, "✓ Online"
      $output += $fmt -f "PDC Emulator", $primaryDC, "✓ Online"
      $output += $fmt -f "RID Master", $primaryDC, "✓ Online"
      $output += $fmt -f "Infrastructure Master", $primaryDC, "✓ Online"
      $output += ""
      $output += "All FSMO role holders are reachable."
    } else {
      $output += "No domain controllers found."
    }
  } else {
    ## Production Mode - actually query FSMO
    $result = Test-FSMORoles -Domain $Domain
    $output += "Checking FSMO roles for domain: $Domain"
    $output += ""

    if ($result.Summary) {
      $fmt = "{0,-30} {1,-30} {2}"
      $output += $fmt -f "Role", "Holder", "Status"
      $output += $fmt -f ("─" * 30), ("─" * 30), ("─" * 15)

      foreach ($line in $result.Summary) {
        if ($line -match '^(.+?)\s*->\s*(.+)$') {
          $role = $matches[1]
          $holder = $matches[2]
          $output += $fmt -f $role, $holder, "✓ Online"
        }
      }
    }
    $output += ""
    $output += "Health: $($result.Health)"
    if ($result.Details) {
      $output += ""
      $output += "Details:"
      $output += $result.Details
    }
  }
  return ($output -join "`n")
  }
}

## This one is pre-merge. I believe it now needs rewritten to be smaller/cleaner
function Get-GPOStatusText {
  param([string]$Domain)

  $output = @()
  $output += ""

  if ($Script:DemoMode) {
    $output += "Demo Mode - Simulated GPO status for domain: $Domain"
    $output += ""
    $demoGPOs = @(
      @{ Name = "Default Domain Policy"; Status = "✓ Healthy" }
      @{ Name = "Default Domain Controllers Policy"; Status = "✓ Healthy" }
      @{ Name = "Password Policy"; Status = "✓ Healthy" }
      @{ Name = "Desktop Settings"; Status = "✓ Healthy" }
      @{ Name = "Security Baseline"; Status = "✓ Healthy" }
    )

    $output += "Total GPOs: $($demoGPOs.Count)"
    $output += ""

    $fmt = "{0,-45} {1}"
    $output += $fmt -f "GPO Name", "Status"
    $output += $fmt -f ("─" * 45), ("─" * 15)
    foreach ($gpo in $demoGPOs) { $output += $fmt -f $gpo.Name, $gpo.Status }
    $output += ""
    $output += "All GPOs in sync."
  } else {
    ## Production Mode - actually query GPOs
    $result = Test-GPOHealth -Domain $Domain
    $output += "Checking GPOs for domain: $Domain"
    $output += ""

    if ($result.Summary) { foreach ($line in $result.Summary) { $output += $line }}
    $output += ""
    $output += "Health: $($result.Health)"

    if ($result.Details) {
      $output += ""
      $output += "Details:"
      $output += $result.Details
    }
  }
  return ($output -join "`n")
}

## ----------{ Production check functions }----------
function Test-ToolsAvailability {
  <#
  .SYNOPSIS
  Tests availability of common diagnostic tools (OS-aware)
  #>

  ## Determine OS using pwsh built-ins
  $os =
    if ($IsWindows) { 'Windows' }
    elseif ($IsLinux) { 'Linux' }
    elseif ($IsMacOS) { 'MacOS' }
    else { 'Windows' }

  <#
  Linux based tools for querying Windows / Active directory
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  NOTE:
  These are NOT Linux-native replacements for the Windows tools. They are the closest Linux-side tools used to QUERY or DIAGNOSE
  Microsoft Active Directory and Windows infrastructure remotely.

  ----------{ repadmin.exe  (AD replication topology & status) }----------
  Closest Linux equivalents:
  - samba-tool drs showrepl <dc>
  - ldapsearch against Configuration/NTDS objects

  Purpose:
  - Inspect replication partners
  - View replication failures and metadata
  - Examine USNs and invocation IDs

  ----------{ dfsrdiag.exe  (DFS-R / SYSVOL replication health) }----------
  No direct Linux equivalent (DFS-R is Windows-only).

  Linux-side validation methods:
  - smbclient //DC/SYSVOL (presence/consistency checks)
  - WinRM from Linux to query DFS Replication event logs

  Purpose:
  - Verify SYSVOL availability
  - Detect missing or inconsistent GPO data

  ----------{ dcdiag.exe  (Domain Controller health checks) }----------
  No single equivalent; functionality is composed from:
  - dig (DNS & SRV record validation)
  - ldapwhoami / ldapsearch (LDAP bind & directory access)
  - kinit / klist (Kerberos authentication)
  - smbclient (SYSVOL / NETLOGON access)

Purpose:
- Validate DNS, LDAP, Kerberos, SMB, and DC reachability

----------{ nltest.exe  (Secure channel, trust, and DC discovery) }----------
Closest Linux equivalents:
- kinit user@REALM (secure channel / trust validation)
- wbinfo -t (trust verification via Samba)
- dig _ldap._tcp / _kerberos._tcp SRV queries

Purpose:
- Domain discovery
- Trust and authentication validation

----------{ csvde.exe  (Directory export) ]----------
Direct Linux equivalent:
- ldapsearch (export AD objects, optionally converted to CSV)

Purpose:
- Extract users, groups, computers, or other AD objects

----------{ portqry.exe  (Port and service reachability) ]----------

Linux equivalents:
- nc (netcat) for TCP/UDP connectivity
- nmap for multi-port and service scanning
- rpcclient for RPC endpoint interrogation

Purpose:
- Validate AD-related port availability (LDAP, Kerberos, SMB, RPC)

----------{ Summary }----------
Windows diagnostic tools tend to be monolithic.
Linux diagnostics are compositional, combining DNS, LDAP,
Kerberos, SMB, and RPC checks to reach the same conclusions.

This block documents the conceptual and operational mapping.

MacOS / Homebrew Availability notes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Most Linux-side AD diagnostic tools are available on macOS via Homebrew.
However, a few are LIMITED or are NOT functionally equivalent on macOS.

----------{ Availablity & functionally equivalent on macOS (via Homebrew) }----------
- ldapsearch / ldapwhoami
  Source: openldap
  Notes: Fully functional; identical usage to Linux.

- dig
  Source: bind
  Notes: Ships with macOS by default; Homebrew version available.

- nc (netcat)
  Source: netcat / openbsd-netcat
  Notes: Built-in on macOS; feature-complete.

- nmap
  Source: nmap
  Notes: Fully supported; often superior to portqry.exe.

- kinit / klist
  Source: krb5
  Notes: macOS includes MIT Kerberos; works natively with AD.

----------{ Partially available / Conditional }----------
- samba-tool drs
  Source: samba
  Notes:
    - Installable via Homebrew
    - DRS commands may be missing or limited depending on build
    - Some Homebrew Samba builds disable AD DC functionality

- wbinfo
  Source: samba
  Notes:
    - Available, but requires Samba + winbind configuration
    - Not commonly used on macOS; more fragile than on Linux

----------{ Limited or not practical on MacOS }----------
- rpcclient
  Source: samba
  Notes:
    - May install via Homebrew
    - Often unreliable due to SMB framework conflicts
    - Not recommended on modern macOS releases

- DFS-R diagnostics (dfsrdiag equivalents)
  Notes:
    - No native DFS-R client exists outside Windows
    - macOS can only validate SYSVOL access, not DFS-R health

----------{ Summary }----------
✔ Most LDAP, DNS, Kerberos, and port-query tooling works identically
  on Linux and macOS.

⚠ Samba-based trust, DRS, and RPC tooling is the weakest area on macOS
  due to build flags, sandboxing, and Apple SMB stack differences.

❌ DFS-R diagnostics remain Windows-only regardless of platform.

This distinction matters when documenting "Linux/macOS from Windows" diagnostic parity.
#>

  ## Tool matrix per OS
  $toolMatrix = @{
    Windows = @('repadmin.exe', 'dfsrdiag.exe', 'dcdiag.exe', 'nltest.exe', 'csvde.exe', 'portqry.exe', 'ping.exe', 'netstat.exe', 'nslookup.exe','PsExec64.exe',
                'PsExec64.exe','psfile64.exe','PsGetsid64.exe','PsInfo64.exe','pskill64.exe','pslist64.exe','PsLoggedon64.exe','psloglist64.exe',
                'pspasswd64.exe','psping64.exe','PsService64.exe','psshutdown64.exe','pssuspend64.exe')
    Linux   = @('dig', 'ldapsearch', 'ldapwhoami', 'kinit', 'klist', 'nmap', 'nc', 'netstat', 'nslookup', 'ping', 'rpcclient', 'smbclient', 'samba-tool', 'wbinfo')
    MacOS   = @('dig', 'kinit', 'klist', 'ldapsearch', 'ldapwhoami', 'nc', 'netstat', 'nmap', 'nslookup', 'ping', 'smbclient')
  }

  <#
    NB: Samba-based tools (samba-tool, wbinfo, rpcclient) are excluded because Homebrew Samba builds on macOS often lack full AD/DC
    features, conflict with Apple’s native SMB stack, or behave inconsistently across macOS releases. LDAP, DNS,  Kerberos, SMB
    client access, and port-scanning tools are stable and fully supported.
  #>

  $tools = @{}
  foreach ($tool in $toolMatrix[$os]) { $tools[$tool] = [bool](Get-Command $tool -ErrorAction SilentlyContinue) }
  return $tools
}

function Invoke-ExternalCommand {
  param(
    [Parameter(Mandatory)]
    [string]$Exe,
    [Parameter(Mandatory)]
    [string]$ArgumentList
  )

  try {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = $Exe
    $psi.Arguments              = $ArgumentList.Trim()
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.Start() | Out-Null

    $out = $proc.StandardOutput.ReadToEnd()
    $err = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    return @{
      ExitCode = $proc.ExitCode
      StdOut   = $out
      StdErr   = $err
    }
  }
  catch {
    return @{
      ExitCode = -1
      StdOut   = ""
      StdErr   = $_.Exception.Message
    }
  }
}

function Refresh-Tools {
  $os = @('Windows','Linux','MacOS')[$osRadio.SelectedItem]

  ## Temporarily override OS for detection
  $script:__oldIsWindows = $IsWindows
  $script:__oldIsLinux   = $IsLinux
  $script:__oldIsMacOS   = $IsMacOS

  switch ($os) {
    'Windows' { $script:IsWindows=$true;  $script:IsLinux=$false; $script:IsMacOS=$false }
    'Linux'   { $script:IsWindows=$false; $script:IsLinux=$true;  $script:IsMacOS=$false }
    'MacOS'   { $script:IsWindows=$false; $script:IsLinux=$false; $script:IsMacOS=$true  }
    }

  $results = Test-ToolsAvailability

  ## Restore globals
  $script:IsWindows = $script:__oldIsWindows
  $script:IsLinux   = $script:__oldIsLinux
  $script:IsMacOS   = $script:__oldIsMacOS
  $names            = @()
  $status           = @()

  foreach ($kv in $results.GetEnumerator()) {
    $names  += $kv.Key
    $status += $(if ($kv.Value) { '✔' } else { '✖' })
  }

  $toolNames.SetSource($names)
  $toolStatus.SetSource($status)

  for ($i = 0; $i -lt $names.Count; $i++) {
    $toolStatus.SetRowColor(
      $i,
      $(if ($results[$names[$i]]) { $greenScheme } else { $redScheme })
      )
    }

  ## Keep columns aligned
  $toolNames.add_SelectedItemChanged({
    $toolStatus.SelectedItem = $toolNames.SelectedItem
    $toolStatus.TopItem      = $toolNames.TopItem
  })
  $toolStatus.add_SelectedItemChanged({
    $toolNames.SelectedItem = $toolStatus.SelectedItem
    $toolNames.TopItem      = $toolStatus.TopItem
  })
  $osRadio.add_SelectedItemChanged({ Refresh-Tools })
  Refresh-Tools
  return $frame
}

function Test-DCStatus {
  param([string]$Domain)

  $summary = @()
  $details = ""
  $health = "OK"

  try {
    $dcs = Get-ADDomainController -Filter * -Server $Domain -ErrorAction Stop

    foreach ($dc in $dcs) {
      $dc   = Get-ADDomainController -Identity $dc.HostName -Server $Domain
      $name = $dc.HostName
      $ip   = if ($dc.IPv4Address) { $dc.IPv4Address } else { "N/A" }
      $site = $dc.Site
      $ping = Test-Connection -ComputerName $name -Count 1 -Quiet -ErrorAction SilentlyContinue
      $reachable = if ($ping) { "OK" } else { "FAIL" }

      $uptimeDays = "?"
      try {
        $wmi = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $name -ErrorAction Stop
        $lastBoot = [Management.ManagementDateTimeConverter]::ToDateTime($wmi.LastBootUpTime)
        $uptimeDays = ((Get-Date) - $lastBoot).Days
      } catch { }

      ## Port checks using Test-NetConnection
      $ldapOK = $false
      $kerbOK = $false
      $smbOK  = $false
      $dnsOK  = $false

      try {
        $ldapOK = (Test-NetConnection -ComputerName $name -Port 389 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue)
        $kerbOK = (Test-NetConnection -ComputerName $name -Port 88 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue)
        $smbOK  = (Test-NetConnection -ComputerName $name -Port 445 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue)
        $dnsOK  = (Test-NetConnection -ComputerName $name -Port 53 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue)
      } catch { }

      $svcNTDS = "?"
      $svcDNS = "?"
      try {
        $svcNTDS = (Get-Service -ComputerName $name -Name "NTDS" -ErrorAction SilentlyContinue).Status
        $svcDNS  = (Get-Service -ComputerName $name -Name "DNS" -ErrorAction SilentlyContinue).Status
      } catch { }

      $obj = [PSCustomObject]@{
        Name       = $name
        IP         = $ip
        Site       = $site
        Reachable  = $reachable
        UptimeDays = $uptimeDays
        NTDS       = $svcNTDS
        DNS        = $svcDNS
        LDAP       = $ldapOK
        Kerberos   = $kerbOK
        SMB        = $smbOK
      }
      $summary += $obj
      if ($reachable -ne "OK") { $health = "FAIL" }
    }
    $details = "Checked $($dcs.Count) domain controller(s) in $Domain"
  } catch {
    $summary = @()
    $details = "Error: $($_.Exception.Message)"
    $health  = "FAIL"
  }
  return @{
    Summary  = $summary
    Details  = $details
    Health   = $health
  }
}

function Test-ADReplication {
  param([string]$Domain)

  $summary = @()
  $details = ""
  $health  = "OK"

  if ($Script:ADHealthTools['repadmin.exe']) {
    ## Actually run repadmin
    $replresult = Invoke-ExternalCommand -Exe "repadmin.exe" -ArgumentList "/replsummary $Domain"
    if ($replresult.ExitCode -eq 0) {
      $lines    = ($replresult.StdOut -split "`r?`n") | Where-Object { $_ -match '\S' } | Select-Object -First 20
      $summary  = $lines
      $details  = $replresult.StdOut
      if ($replresult.StdOut -match "error|fail") { $health = "FAIL" }
    } else {
      $summary += "repadmin returned error code $($replresult.ExitCode)"
      $details  = $replresult.StdErr
      $health   = "WARN"
    }
  } else {
    ## Fallback to AD cmdlet
    try {
      $failures = Get-ADReplicationFailure -Scope Domain -Target $Domain -ErrorAction Stop
      if ($failures) {
        foreach ($failure in $failures) { $summary += "FAIL: $($failure.Server) - $($failure.FirstFailureMessage)" }
        $health = "FAIL"
      } else {
        $summary += "No replication failures detected."
      }
    } catch {
      $summary += "Error checking replication: $($_.Exception.Message)"
      $health = "WARN"
    }
  }

  return @{
    Summary = $summary
    Details = $details
    Health  = $health
  }
}

function Test-ADDnsRecords {
  param([string]$Domain)

  $summary = @()
  $details = ""
  $health  = "OK"

  $srvRecords = @(
    "_ldap._tcp.dc._msdcs.$Domain"
    "_kerberos._tcp.$Domain"
    "_ldap._tcp.$Domain"
  )

  ## Actually query DNS
  foreach ($srv in $srvRecords) {
    try {
      $q        = Resolve-DnsName -Name $srv -Type SRV -ErrorAction Stop
      $count    = ($q | Measure-Object).Count
      $summary += "$srv : $count record(s)"
      $details += "SRV $srv : $count records`n"
    } catch {
      $summary += "$srv : NOT FOUND"
      $details += "SRV $srv : Error - $($_.Exception.Message)`n"
      $health = "FAIL"
    }
  }

  ## Check DNS service on DCs
  try {
    $dcs = Get-ADDomainController -Filter * -Server $Domain -ErrorAction Stop
    foreach ($dc in $dcs) {
      try {
        $svc = (Get-Service -ComputerName $dc.HostName -Name "DNS" -ErrorAction SilentlyContinue).Status
        $summary += "DNS on $($dc.HostName): $svc"
      } catch {
        $summary += "DNS on $($dc.HostName): Error"
        $health = "WARN"
      }
    }
  } catch {
    $summary += "Could not enumerate DCs for DNS check"
    $health = "WARN"
  }

  return @{
    Summary = $summary
    Details = $details
    Health  = $health
  }
}

function Test-SysvolHealth {
  param([string]$Domain)

  $summary = @()
  $details = ""
  $health = "OK"

  try {
    $dcs = Get-ADDomainController -Filter * -Server $Domain -ErrorAction Stop

    foreach ($dc in $dcs) {
      ## Actually check shares
      try {
        $sysvolPath     = "\\$($dc.HostName)\SYSVOL"
        $netlogonPath   = "\\$($dc.HostName)\NETLOGON"
        $sysvolOK       = Test-Path $sysvolPath -ErrorAction SilentlyContinue
        $netlogonOK     = Test-Path $netlogonPath -ErrorAction SilentlyContinue
        $sysvolStatus   = if ($sysvolOK) { "✓ Available" } else { "✗ Unavailable"; $health = "FAIL" }
        $netlogonStatus = if ($netlogonOK) { "✓ Available" } else { "✗ Unavailable"; $health = "FAIL" }

        $summary += "SYSVOL on $($dc.HostName): $sysvolStatus"
        $summary += "NETLOGON on $($dc.HostName): $netlogonStatus"
        $details += "$($dc.HostName) - SYSVOL: $sysvolStatus, NETLOGON: $netlogonStatus`n"
      } catch {
        $summary += "Error checking $($dc.HostName): $($_.Exception.Message)"
        $details += "Error on $($dc.HostName): $($_.Exception.Message)`n"
        $health = "WARN"
      }
    }
  } catch {
    $summary += "Could not enumerate DCs for SYSVOL check"
    $details = "Error: $($_.Exception.Message)"
    $health = "FAIL"
  }

  return @{
    Summary = $summary
    Details = $details
    Health  = $health
  }
}

function Test-FSMORoles {
  param([string]$Domain)

  $summary = @()
  $details = ""
  $health  = "OK"

  try {
    $forest = Get-ADForest -Identity $Domain -ErrorAction Stop
    $domainObj = Get-ADDomain -Identity $Domain -ErrorAction Stop

    $summary += "Schema Master -> $($forest.SchemaMaster)"
    $summary += "Domain Naming Master -> $($forest.DomainNamingMaster)"
    $summary += "PDC Emulator -> $($domainObj.PDCEmulator)"
    $summary += "RID Master -> $($domainObj.RIDMaster)"
    $summary += "Infrastructure Master -> $($domainObj.InfrastructureMaster)"

    $details = "Forest FSMO:`n"
    $details += "  SchemaMaster: $($forest.SchemaMaster)`n"
    $details += "  DomainNamingMaster: $($forest.DomainNamingMaster)`n"
    $details += "`nDomain FSMO:`n"
    $details += "  PDCEmulator: $($domainObj.PDCEmulator)`n"
    $details += "  RIDMaster: $($domainObj.RIDMaster)`n"
    $details += "  InfrastructureMaster: $($domainObj.InfrastructureMaster)`n"

    ## Test reachability
    $holders = @($forest.SchemaMaster, $forest.DomainNamingMaster, $domainObj.PDCEmulator, $domainObj.RIDMaster, $domainObj.InfrastructureMaster) | Where-Object { $_ } | Select-Object -Unique

    foreach ($holder in $holders) {
      $ping = Test-Connection -ComputerName $holder -Count 1 -Quiet -ErrorAction SilentlyContinue
      if (-not $ping) {
        $health = "WARN"
        $details += "`nWARNING: Cannot reach $holder`n"
      }
    }
  } catch {
    $summary += "Error checking FSMO roles: $($_.Exception.Message)"
    $details = "Error: $($_.Exception.Message)"
    $health = "FAIL"
  }
  return @{
    Summary = $summary
    Details = $details
    Health  = $health
  }
}

function Test-GPOHealth {
  param([string]$Domain)

  $summary = @()
  $details = ""
  $health = "OK"

  try {
    ## Try to get GPOs from imported data first, then fall back to production
    $gpos = @()

    if ($Script:rawGPOs -and $Script:rawGPOs.Count -gt 0) {
      ## Use imported GPO data
      Debug-Log "Using imported GPO data ($($Script:rawGPOs.Count) GPOs)" -Type "Insight"

      ## Filter by domain if specified
      if ($Domain) {
        $gpos = $Script:rawGPOs | Where-Object { $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')") }
      } else {
        $gpos = $Script:rawGPOs
      }
    } else {
      ## Fall back to production query
      Debug-Log "Querying GPOs from Active Directory for domain: $Domain" -Type "Insight"
      $gpos = Get-GPO -All -Domain $Domain -ErrorAction Stop
    }
    if ($gpos.Count -eq 0) {
      $summary += "No GPOs found"
      $health = "WARN"
      return @{
        Summary = $summary
        Details = "No Group Policy Objects found for domain: $Domain"
        Health  = $health
      }
    }

    ## Build summary
    $summary += "Total GPOs: $($gpos.Count)"
    $summary += ""

    ## Format GPO list
    $fmt = "{0,-50} {1}"
    $summary += $fmt -f "GPO Name", "Status"
    $summary += $fmt -f ("─" * 50), ("─" * 15)
    $details  = "GPO Details:`n`n"

    foreach ($gpo in $gpos | Select-Object -First 50) {
      ## Handle both imported GPO format and production AD GPO objects
      $gpoName    = if ($gpo.DisplayName) { $gpo.DisplayName } else { $gpo.Name }
      $gpoPath    = if ($gpo.GPCFileSysPath) { $gpo.GPCFileSysPath } else { "N/A" }
      $gpoVersion = if ($gpo.VersionNumber) { $gpo.VersionNumber } else { "N/A" }

      $status = "✓ Present"

      $summary += $fmt -f $gpoName, $status
      $details += "  Name: $gpoName`n"
      $details += "    Path: $gpoPath`n"
      $details += "    Version: $gpoVersion`n"

      if ($gpo.Description) { $details += "    Description: $($gpo.Description)`n" }
      $details += "`n"
    }

    if ($gpos.Count -gt 50) {
      $summary += ""
      $summary += "... and $($gpos.Count - 50) more GPOs"
      $details += "`n... and $($gpos.Count - 50) more GPOs not shown"
    }

    $summary += ""
    $summary += "All GPOs accounted for."

  } catch {
    $summary += "Error checking GPOs: $($_.Exception.Message)"
    $details = "Error: $($_.Exception.Message)"
    $health = "FAIL"
  }

  return @{
    Summary = $summary
    Details = $details
    Health  = $health
    GPOs    = $gpos  ## Return the GPO objects for further use
  }
}

function Get-GPOStatusText {
  param([string]$Domain)

  if (-not $Domain) { $Domain = $Script:CurrentDomain }
  $result = Test-GPOHealth -Domain $Domain
  $output = @()
  $output += ""
  $output += "Group Policy Objects - Domain: $Domain"
  $output += ""
  if ($result.Summary) { foreach ($line in $result.Summary) { $output += $line  }}
  $output += ""
  $output += "Health: $($result.Health)"
  if ($result.Details) {
    $output += ""
    $output += $result.Details
  }
  return ($output -join "`n")
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

  ## Generate-RandomPassword UI
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

  ## Entropy display
  $lblStrength = [Terminal.Gui.Label]::new(30,7,"Strength: Not generated")
  $lblStrength.Width = 30
  $dlg.Add($lblStrength)
  $progEntropy = [Terminal.Gui.ProgressBar]::new()
  $progEntropy.X = 30; $progEntropy.Y = 8; $progEntropy.Width = 25
  $progEntropy.Fraction = 0.0
  $progEntropy = [Terminal.Gui.ProgressBar]::new()
  $progEntropy.X = 30; $progEntropy.Y = 8; $progEntropy.Width = 25
  $progEntropy.Fraction = 0.0

  ## Apply theme colours to progress bar
  if ($Script:ThemeMode) {
    $themeData = Get-Theme -mode $Script:ThemeMode
    if ($themeData -and $themeData.MainWindow) {
      ## Create a custom ColorScheme for the progress bar
      $progColorScheme = [Terminal.Gui.ColorScheme]::new()
      ## Use MainWindow colours for the background (matches dialog background)
      $progColorScheme.Normal = [Terminal.Gui.Attribute]::Make(
        $themeData.MainWindow.Normal.Foreground,
        $themeData.MainWindow.Normal.Background
      )
      $progColorScheme.Focus = [Terminal.Gui.Attribute]::Make(
        $themeData.MainWindow.Normal.Foreground,
        $themeData.MainWindow.Normal.Background
      )
      ## The filled bar - use MainWindow Focus colours to stand out
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

  ## Generate Logic
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
      if ($entropy -lt 40)     { $strengthCategory="Weak"; $fraction=0.2 }
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
    }).GetNewClosure()

    ## Show Password toggle
    $chkShowPwd.add_Toggled({
      if ($chkShowPwd.Checked) { $txtPwd.Text = $Script:actualPassword
      } else { $txtPwd.Text = ('*' * $Script:actualPassword.Length)
      }
    })

    ## Copy to Clipboard
    $btnCopy.add_Clicked({
      if (-not $Script:actualPassword) { return }
      if ($IsWindows) { Set-Clipboard -Value $Script:actualPassword }
      elseif ($IsMacOS) { $Script:actualPassword | pbcopy }
      else { $Script:actualPassword | xsel --clipboard --input }
      Show-Modal "Copied" "Password copied to clipboard."
    }).GetNewClosure()
    ## Close
    $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
  [Terminal.Gui.Application]::Run($dlg)
  return $Script:actualPassword
}

function Dump-LAPSDevices {
  <#
    .SYNOPSIS
    Dumps devices with LAPS data and outputs only fully-populated attributes

    .DESCRIPTION
    Filters computer records that have either Legacy LAPS (ms-Mcs-AdmPwd)
    or Windows LAPS (msLAPS-*) attributes populated, then removes any
    property that is empty or missing on any matching device.

    Designed for CSV-imported AD dumps (all values are strings).
  #>

  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [object[]]$Computers
  )

  ##  Step 1: identify LAPS-enabled devices
  $LAPSDevices = $Computers | Where-Object {

    ## Legacy LAPS
    ( $_.PSObject.Properties['ms-Mcs-AdmPwd'] -and -not [string]::IsNullOrWhiteSpace($_.'ms-Mcs-AdmPwd') ) -or
    ## Windows LAPS (any populated msLAPS-* attribute)
    ( $_.PSObject.Properties | Where-Object { $_.Name -like 'msLAPS-*' -and -not [string]::IsNullOrWhiteSpace($_.Value) } )
  }

  if (-not $LAPSDevices) {
    Debug-Log "No LAPS-enabled devices found" -Type "Warning"
    return @()
  }

  ##  Step 2: find properties populated on ALL LAPS devices
  $AllProps    = $LAPSDevices[0].PSObject.Properties.Name
  $CommonProps = $AllProps | Where-Object { $prop = $_; -not ($LAPSDevices | Where-Object { -not $_.PSObject.Properties[$prop] -or [string]::IsNullOrWhiteSpace($_.$prop) }) }
  ##  Step 3: output only those properties
  $LAPSDevices | Select-Object -Property $CommonProps
}

## Show LAPS Passwords
function Show-LAPSSearchModal {
  <#
  .SYNOPSIS
  Modal LAPS lookup UI with inline password reveal/hide

  .DESCRIPTION
  - Browse all LAPS-enabled computers
  - Search by computer name
  - Inline masked password with Reveal/Hide toggle
  - Copy buttons next to each field
  - Expiry warning highlighting
  - DemoMode support
  #>

  Debug-Log "Show-LAPSSearchModal called" -Type "Insight"

  ## Create dialog
  $dialog = [Terminal.Gui.Dialog]::new("LAPS Password Lookup", 50, 25)

  $lblSearch = [Terminal.Gui.Label]::new("Computer name (blank = all):")
  $lblSearch.X = 1; $lblSearch.Y = 1
  $dialog.Add($lblSearch)

  $txtSearch = [Terminal.Gui.TextField]::new("")
  $txtSearch.X = 1; $txtSearch.Y = 2; $txtSearch.Width = 30
  $dialog.Add($txtSearch)

  $lstComputers = [Terminal.Gui.ListView]::new()
  $lstComputers.X = 1; $lstComputers.Y = 4
  $lstComputers.Width = 30; $lstComputers.Height = 16
  $dialog.Add($lstComputers)

  $Script:LAPSComputers = @()

  ## ----------{ Data Loader }----------
  $loadComputers = {
    param($filter)

    Debug-Log "Loading LAPS computers with filter: '$filter'" -Type "Insight"
    try {
      if ($Script:DemoMode) {
        ## Demo mode - check computers for LAPS properties
        Debug-Log "Demo mode - using sample data" -Type "Insight"
        Debug-Log "Total computers: $($Script:Computers.Count)" -Type "Tracing"
        $Script:LAPSComputers = @()
        foreach ($comp in $Script:Computers) {
          if ($filter -and $comp.Name -notlike "*$filter*") { continue } ## Skip if filter doesn't match

          ## Check for LAPS properties
          $hasLaps = $false
          $lapsUser = ""
          $lapsPass = ""
          $lapsExpiry = ""

          ## Check Windows LAPS
          if ($comp.'msLAPS-Password') {
            $hasLaps    = $true
            $lapsUser   = if ($comp.'msLAPS-AccountName') { $comp.'msLAPS-AccountName' } else { 'Administrator' }
            $lapsPass   = $comp.'msLAPS-Password'
            $lapsExpiry = $comp.'msLAPS-PasswordExpirationTime'
            Debug-Log "$($comp.Name) has Windows LAPS (pass length: $($lapsPass.Length))" -Type "Tracing"
          }
          ## Check Legacy LAPS
          elseif ($comp.'ms-Mcs-AdmPwd') {
            $hasLaps    = $true
            $lapsUser   = 'Administrator'
            $lapsPass   = $comp.'ms-Mcs-AdmPwd'
            $lapsExpiry = $comp.'ms-Mcs-AdmPwdExpirationTime'
            Debug-Log "$($comp.Name) has Legacy LAPS (pass length: $($lapsPass.Length))" -Type "Tracing"
          }
          if ($hasLaps) {
            $Script:LAPSComputers += [PSCustomObject]@{
              Name        = $comp.Name
              DNSHostName = $comp.DNSHostName ?? "$($comp.Name).$($Script:CurrentDomain)"
              LapsUser    = $lapsUser
              LapsPass    = $lapsPass
              LapsExpiry  = $lapsExpiry
            }
          }
        }
        Debug-Log "Found $($Script:LAPSComputers.Count) computers with LAPS" -Type "Insight"

      } else {
        ## Production Mode - detect LAPS schema once
        if (-not $Script:LAPSSchemaDetected) {
          try {
            $windowsLapsSchema = Get-ADObject -SearchBase (Get-ADRootDSE).SchemaNamingContext  -LDAPFilter "(lDAPDisplayName=msLAPS-Password)" -ErrorAction SilentlyContinue
            if ($windowsLapsSchema) {
              $Script:LAPSType = "Windows"
              $Script:LAPSProps = @('Name', 'DNSHostName', 'msLAPS-Password', 'msLAPS-PasswordExpirationTime', 'msLAPS-AccountName')
              Debug-Log "Detected Windows LAPS schema" -Type "Insight"
            } else {
              $Script:LAPSType = "Legacy"
              $Script:LAPSProps = @('Name', 'DNSHostName', 'ms-Mcs-AdmPwd', 'ms-Mcs-AdmPwdExpirationTime')
              Debug-Log "Detected Legacy LAPS schema" -Type "Insight"
            }
            $Script:LAPSSchemaDetected = $true
          } catch {
            Debug-Log "Schema detection failed, defaulting to Legacy LAPS" -Type "Warning"
            $Script:LAPSType = "Legacy"
            $Script:LAPSProps = @('Name', 'DNSHostName', 'ms-Mcs-AdmPwd', 'ms-Mcs-AdmPwdExpirationTime')
            $Script:LAPSSchemaDetected = $true
          }
        }

        ## Query AD with correct properties
        $filterString = if ([string]::IsNullOrWhiteSpace($filter)) { "*" } else { "*$filter*" }
        Debug-Log "Querying AD with filter: Name -like '$filterString'" -Type "Insight"
        try {
          $rawComputers = Get-ADComputer -Filter "Name -like '$filterString'" -Properties $Script:LAPSProps -ErrorAction Stop
        } catch {
          Debug-Log "AD query failed: $($_.Exception.Message)" -Type "Problem"
          throw
        }

        ## Filter to only computers with LAPS passwords
        $Script:LAPSComputers = @()
        foreach ($comp in $rawComputers) {
          $hasLaps = $false
          $lapsUser = ""
          $lapsPass = ""
          $lapsExpiry = ""
          if ($Script:LAPSType -eq "Windows") {
            if ($comp.'msLAPS-Password') {
              $hasLaps    = $true
              $lapsUser   = $comp.'msLAPS-AccountName'
              $lapsPass   = $comp.'msLAPS-Password'
              $lapsExpiry = $comp.'msLAPS-PasswordExpirationTime'
            }
          } else {
            if ($comp.'ms-Mcs-AdmPwd') {
              $hasLaps    = $true
              $lapsUser   = 'Administrator'
              $lapsPass   = $comp.'ms-Mcs-AdmPwd'
              $lapsExpiry = $comp.'ms-Mcs-AdmPwdExpirationTime'
            }
          }
          if ($hasLaps) {
            $Script:LAPSComputers += [PSCustomObject]@{
              Name        = $comp.Name
              DNSHostName = $comp.DNSHostName
              LapsUser    = $lapsUser
              LapsPass    = $lapsPass
              LapsExpiry  = $lapsExpiry
            }
          }
        }
        Debug-Log "Found $($Script:LAPSComputers.Count) computers with LAPS" -Type "Insight"
      }
      ## Update ListView
      if ($Script:LAPSComputers.Count -eq 0) {
        $lstComputers.SetSource(@("(No computers with LAPS found)"))
      } else {
        $computerNames = $Script:LAPSComputers | ForEach-Object { $_.Name }
        $lstComputers.SetSource($computerNames)
      }
    } catch {
      Debug-Log "Error loading LAPS computers: $($_.Exception.Message)" -Type "Problem"
      Show-Modal "Error" "Failed to load LAPS computers:`n`n$($_.Exception.Message)"
      $lstComputers.SetSource(@("(Error loading computers)"))
    }
  }
  ## Initial load
  & $loadComputers ""

  ## ----------{ Buttons }----------
  $btnSearch = [Terminal.Gui.Button]::new("Search")
  $btnSearch.X = 32; $btnSearch.Y = 2
  $btnSearch.add_Clicked({
    $searchText = $txtSearch.Text.ToString()
    & $loadComputers $searchText
  }).GetNewClosure()
  $dialog.Add($btnSearch)

  $btnView = [Terminal.Gui.Button]::new("View")
  $btnView.X = 10; $btnView.Y = 21
  $dialog.Add($btnView)

  $btnClose = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = 30; $btnClose.Y = 21
  $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
  $dialog.Add($btnClose)

  ## ----------{ View Handler }----------
  $btnView.add_Clicked({
    if ($lstComputers.SelectedItem -lt 0 -or $Script:LAPSComputers.Count -eq 0) {
      Show-Modal "No Selection" "Please select a computer from the list."
      return
    }
    try {
      $computer = $Script:LAPSComputers[$lstComputers.SelectedItem]
      Debug-Log "Selected computer: $($computer.Name)" -Type "Insight"
      Debug-Log "LapsUser: '$($computer.LapsUser)'" -Type "Tracing"
      Debug-Log "LapsPass length: $($computer.LapsPass.Length)" -Type "Tracing"
      Debug-Log "LapsExpiry: '$($computer.LapsExpiry)'" -Type "Tracing"
      ## Validate we have password data
      if ([string]::IsNullOrWhiteSpace($computer.LapsPass)) {
        Show-Modal "No LAPS Data" "No LAPS password found for $($computer.Name)"
        return
      }
      ## Calculate expiry
      $expires = Get-Date
      $daysLeft = 999
      if ($computer.LapsExpiry) {
        try {
          if ($computer.LapsExpiry -is [DateTime]) {
            $expires = $computer.LapsExpiry
          } else {
            $expires = [DateTime]::FromFileTimeUtc([long]$computer.LapsExpiry)
          }
          $daysLeft = [Math]::Round(($expires - (Get-Date)).TotalDays)
        } catch {
          Debug-Log "Failed to parse expiry: $($_.Exception.Message)" -Type "Warning"
        }
      }

      ## Expiry warning
      $warn = if ($daysLeft -le 3) { "WARNING: Password expires in $daysLeft day(s)"
      } elseif ($daysLeft -le 7) { "Password expires in $daysLeft days"
      } else { "Expires in $daysLeft days"
      }

      ## Create detail dialog
      $detailDlg = [Terminal.Gui.Dialog]::new("LAPS Details - $($computer.Name)", 50, 14)

      ## Computer name
      $lblComputer = [Terminal.Gui.Label]::new("Computer: $($computer.Name)")
      $lblComputer.X = 2; $lblComputer.Y = 1
      $detailDlg.Add($lblComputer)

      ## LAPS User
      $lblUserLabel = [Terminal.Gui.Label]::new("LAPS User:")
      $lblUserLabel.X = 2; $lblUserLabel.Y = 3
      $detailDlg.Add($lblUserLabel)

      $lblUserValue = [Terminal.Gui.Label]::new($computer.LapsUser)
      $lblUserValue.X = 18; $lblUserValue.Y = 3; $lblUserValue.Width = 20
      $detailDlg.Add($lblUserValue)

      $btnCopyUser = [Terminal.Gui.Button]::new("📋")
      $btnCopyUser.X = 38; $btnCopyUser.Y = 3; $btnCopyUser.Width = 6
      $btnCopyUser.add_Clicked({
        Set-Clipboard -Value $computer.LapsUser
        Show-Modal "Copied" "Username copied to clipboard."
      })
      $detailDlg.Add($btnCopyUser)

      ## LAPS Password (with reveal/hide)
      $lblPasswordLabel    = [Terminal.Gui.Label]::new("LAPS Password:")
      $lblPasswordLabel.X = 2; $lblPasswordLabel.Y = 5
      $detailDlg.Add($lblPasswordLabel)

      ## Ensure password is a string and create masked version
      $passStr = $computer.LapsPass.ToString()
      $masked  = '*' * $passStr.Length
      $Script:LAPSPasswordRevealed = $false

      $lblPasswordValue   = [Terminal.Gui.Label]::new($masked)
      $lblPasswordValue.X = 18; $lblPasswordValue.Y = 5; $lblPasswordValue.Width = 16
      $detailDlg.Add($lblPasswordValue)

      ## Show/Hide button (text-based)
      $btnReveal   = [Terminal.Gui.Button]::new("👁")
      $btnReveal.X = 32; $btnReveal.Y = 5; $btnReveal.Width = 6
      $btnReveal.add_Clicked({
        if ($Script:LAPSPasswordRevealed) {
          ## Hide password
          $lblPasswordValue.Text = [NStack.ustring]::Make($masked)
          $btnReveal = [Terminal.Gui.Button]::new("👁")
          $Script:LAPSPasswordRevealed = $false
          Debug-Log "Password hidden" -Type "Tracing"
        } else {
          ## Show password
          $lblPasswordValue.Text = [NStack.ustring]::Make($passStr)
          $btnReveal.Text = [NStack.ustring]::Make("🕶️")
          $Script:LAPSPasswordRevealed = $true
          Debug-Log "Password revealed" -Type "Tracing"
        }
      })
      $detailDlg.Add($btnReveal)

      $btnCopyPassword   = [Terminal.Gui.Button]::new("📋")
      $btnCopyPassword.X = 38; $btnCopyPassword.Y = 5; $btnCopyPassword.Width = 6
      $btnCopyPassword.add_Clicked({
        Set-Clipboard -Value $passStr
        Show-Modal "Copied" "Password copied to clipboard."
      })
      $detailDlg.Add($btnCopyPassword)

      ## Expiry warning
      $lblExpiry   = [Terminal.Gui.Label]::new($warn)
      $lblExpiry.X = 2; $lblExpiry.Y = 7; $lblExpiry.Width = 45
      $detailDlg.Add($lblExpiry)

      ## Force Rotation button (only in production)
      if (-not $Script:DemoMode) {
        $btnRotate = [Terminal.Gui.Button]::new("🔁 Force Rotation")
        $btnRotate.X = 2; $btnRotate.Y = 9; $btnRotate.Width = 16
        $btnRotate.add_Clicked({
          try {
            $adComputer = Get-ADComputer -Identity $computer.Name -Properties 'ms-Mcs-AdmPwdExpirationTime'
            Set-ADComputer -Identity $adComputer -Replace @{ 'ms-Mcs-AdmPwdExpirationTime' = '0' }
            Debug-Log "LAPS password expiration set to 0 for $($computer.Name)" -Type "Success"
            try {
              Invoke-GPUpdate -Computer $computer.Name -Force -ErrorAction Stop
              Show-Modal "Rotation Complete" "LAPS password rotation initiated!`n`nExpiration set to 0 and GP updated."
            } catch {
              Show-Modal "Rotation Requested" "LAPS will rotate on next GP refresh.`n`nExpiration set to 0."
            }
          } catch {
            Show-Modal "Error" "Failed to force rotation:`n`n$($_.Exception.Message)"
            Debug-Log "Failed to force LAPS rotation: $($_.Exception.Message)" -Type "Problem"
          }
        })
        $detailDlg.Add($btnRotate)
      }

      ## Close button
      $btnOk = [Terminal.Gui.Button]::new("Close")
      $btnOk.X = 22; $btnOk.Y = 9; $btnOk.Width = 7
      $btnOk.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
      $detailDlg.Add($btnOk)
      [Terminal.Gui.Application]::Run($detailDlg)
    } catch {
      Show-Modal "Error" "Failed to show LAPS details:`n`n$($_.Exception.Message)"
      Debug-Log "Error showing LAPS details: $($_.Exception.Message)" -Type "Problem"
    }
  }).GetNewClosure()
  [Terminal.Gui.Application]::Run($dialog)
}

## ----------{ Main Dialog Function }----------

# DSA-TUI Object Management Module v1.0
# Create, Delete, and Move AD Objects

## ----------{ Create New Object Wizard }----------
function Show-NewObjectWizard {
  $dlg = [Terminal.Gui.Dialog]::new("New Object Wizard", 74, 30)

  ## Step 1: Select object type
  $lblType = [Terminal.Gui.Label]::new("Select object type to create:"); $lblType.X=2; $lblType.Y=1; $dlg.Add($lblType)
  $rdoType = [Terminal.Gui.RadioGroup]::new(@("User", "Group", "Organizational Unit", "Computer", "Contact"))
  $rdoType.X=2; $rdoType.Y=3; $rdoType.Height=5
  $dlg.Add($rdoType)

  ## Common fields
  $y = 9
  $lblName = [Terminal.Gui.Label]::new("Name:"); $lblName.X=2; $lblName.Y=$y; $dlg.Add($lblName)
  $txtName = [Terminal.Gui.TextField]::new(""); $txtName.X=20; $txtName.Y=$y; $txtName.Width=45; $dlg.Add($txtName)
  $y+=2

  $lblDisplayName = [Terminal.Gui.Label]::new("Display Name:"); $lblDisplayName.X=2; $lblDisplayName.Y=$y; $dlg.Add($lblDisplayName)
  $txtDisplayName = [Terminal.Gui.TextField]::new(""); $txtDisplayName.X=20; $txtDisplayName.Y=$y; $txtDisplayName.Width=45; $dlg.Add($txtDisplayName)
  $y+=2

  $lblOU = [Terminal.Gui.Label]::new("Organizational Unit:"); $lblOU.X=2; $lblOU.Y=$y; $dlg.Add($lblOU)

  ## Get list of OUs
  $ouList = if ($Script:DemoMode) {
    $Script:Users | Select-Object -ExpandProperty OU -Unique | Sort-Object
  } else {
    try {
      Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName | Select-Object -ExpandProperty DistinguishedName | Sort-Object
    } catch { @("CN=Users,DC=example,DC=com") }
  }

  $cmbOU = [Terminal.Gui.ComboBox]::new()
  $cmbOU.X=20; $cmbOU.Y=$y; $cmbOU.Width=45
  $cmbOU.SetSource($ouList)
  $dlg.Add($cmbOU)
  $y+=2

  ## User-specific fields (shown/hidden based on type)
  $lblSam = [Terminal.Gui.Label]::new("Username (SAM):"); $lblSam.X=2; $lblSam.Y=$y; $dlg.Add($lblSam)
  $txtSam = [Terminal.Gui.TextField]::new(""); $txtSam.X=20; $txtSam.Y=$y; $txtSam.Width=45; $dlg.Add($txtSam)
  $y+=2

  $lblEmail = [Terminal.Gui.Label]::new("Email:"); $lblEmail.X=2; $lblEmail.Y=$y; $dlg.Add($lblEmail)
  $txtEmail = [Terminal.Gui.TextField]::new(""); $txtEmail.X=20; $txtEmail.Y=$y; $txtEmail.Width=45; $dlg.Add($txtEmail)
  $y+=2

  $lblPassword = [Terminal.Gui.Label]::new("Password:"); $lblPassword.X=2; $lblPassword.Y=$y; $dlg.Add($lblPassword)
  $txtPassword = [Terminal.Gui.TextField]::new(""); $txtPassword.X=20; $txtPassword.Y=$y; $txtPassword.Width=45; $txtPassword.Secret=$true; $dlg.Add($txtPassword)

  ## Show/hide fields based on type
  $rdoType.add_SelectedItemChanged({
    $isUser              = $rdoType.SelectedItem -eq 0
    $lblSam.Visible      = $isUser
    $txtSam.Visible      = $isUser
    $lblEmail.Visible    = $isUser
    $txtEmail.Visible    = $isUser
    $lblPassword.Visible = $isUser
    $txtPassword.Visible = $isUser
  })

  ## Create button
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
      ## Demo mode - add to in-memory structures
      switch ($objType) {
        "User" {
          $sam = $txtSam.Text.ToString().Trim()
          $email = $txtEmail.Text.ToString().Trim()
          if (-not $sam) { $sam = $name.ToLower().Replace(' ', '.') }
          if (-not $email) { $email = "$sam@example.com" }

          ## Add to RAW user data
          $newUser             = @{
            Name               = $name
            OU                 = @($ou)  # OU as array
            Groups             = @()
            Title              = ""
            Email              = $email
            Country            = ""
            Disabled           = $false
            Locked             = $false
            MustChangePassword = $true
            Department         = ""
            Office             = ""
            Phone              = ""
            MobilePhone        = ""
            Street             = ""
            City               = ""
            PostalCode         = ""
            Company            = ""
            Manager            = ""
            Description        = $displayName
          }

          ## Add to raw users array
          $Script:rawUsers += $newUser
          ## Reconvert to update $Script:Users with AD-like objects
          $converted = Convert-DataToADObjects -Users $Script:rawUsers -DCs $Script:rawDCs -Groups $Script:rawDemoGroups -Domain $Script:CurrentDomain
          $Script:Users = $converted.Users
          Debug-Log (" Created user $name in demo mode") -Type "Insight"
        }

        "Group" {
          ## Add to RAW group data
          $newGroup = @{
            Name        = $name
            Description = $displayName
            Type        = 'Security'
            Scope       = 'Global'
            ManagedBy   = ''
            Email       = ''
          }
          ## Add to raw groups array
          $Script:rawDemoGroups += $newGroup
          ## Reconvert to update $Script:Groups with AD-like objects
          $converted = Convert-DataToADObjects -Users $Script:rawUsers -DCs $Script:rawDCs -Groups $Script:rawDemoGroups -Domain $Script:CurrentDomain
          $Script:Groups = $converted.Groups
          Debug-Log (" Created group $name in demo mode") -Type "Insight"
        }
        "OrganizationalUnit" {
          ## For OUs, we need to track them in a structure. OUs are built from the OU arrays in users, so we could either:
          ## 1. Add a $Script:rawOUs array (cleaner)
          ## 2. Just ensure the OU path exists when we rebuild the tree

          ## For now, let's just ensure it's tracked
          if (-not $Script:rawOUs) { $Script:rawOUs = @() }
          $newOU = @{
            Name        = $name
            Path        = $ou
            Description = $displayName
          }
          $Script:rawOUs += $newOU
          Debug-Log (" Created OU $name in demo mode") -Type "Insight"
        }
        "Computer" {
          ## Similar to users, add to a computers array
          if (-not $Script:rawComputers) {
            $Script:rawComputers = @()
          }
          $newComputer = @{
            Name        = $name
            OU          = $ou
            Description = $displayName
          }
          $Script:rawComputers += $newComputer
          Debug-Log (" Created computer $name in demo mode") -Type "Insight"
        }
      }
      Show-Modal "Success" "$objType '$name' created successfully (demo mode)"
      Build-Tree -domain $Script:CurrentDomain
      Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
      [Terminal.Gui.Application]::RequestStop()
    } else {
      ## Production Mode - create in AD
      switch ($objType) {
        "User" {
          $sam      = $txtSam.Text.ToString().Trim()
          $email    = $txtEmail.Text.ToString().Trim()
          $password = $txtPassword.Text.ToString()
          if (-not $sam) {
            Show-Modal "Error" "Username (SAM) is required for users!"
            return
          }
          if (-not $password) {
            Show-Modal "Error" "Password is required for users!"
            return
          }
          $secPwd = ConvertTo-SecureString -String $password -AsPlainText -Force
          $params = @{
            Name                  = $name
            SamAccountName        = $sam
            UserPrincipalName     = "$sam@$($Script:CurrentDomain)"
            AccountPassword       = $secPwd
            Enabled               = $true
            Path                  = $ou
            ChangePasswordAtLogon = $true
          }
          if ($displayName) { $params['DisplayName'] = $displayName }
          if ($email) { $params['EmailAddress'] = $email }
          New-ADUser @params -ErrorAction Stop
          Debug-Log (" Created user $name in AD") -Type "Insight"
        }
        "Group" {
          $params = @{
          Name          = $name
          GroupScope    = "Global"
          GroupCategory = "Security"
          Path          = $ou
        }

        if ($displayName) { $params['Description'] = $displayName }
          New-ADGroup @params -ErrorAction Stop
          Debug-Log (" Created group $name in AD") -Type "Insight"
        }

        "OrganizationalUnit" {
          $params = @{
            Name = $name
            Path = $ou
          }

          if ($displayName) { $params['Description'] = $displayName }
            New-ADOrganizationalUnit @params -ErrorAction Stop
            Debug-Log (" Created OU $name in AD") -Type "Insight"
          }
        "Computer" {
          $params = @{
            Name  = $name
            Path  = $ou
          }

          New-ADComputer @params -ErrorAction Stop
          Debug-Log (" Created computer $name in AD") -Type "Insight"
        }
        "Contact" {
          $params = @{
            Name = $name
            Type = "Contact"
            Path = $ou
          }

          if ($displayName) { $params['DisplayName'] = $displayName }
          New-ADObject @params -ErrorAction Stop
          Debug-Log (" Created contact $name in AD") -Type "Insight"
        }
      }
      Show-Modal "Success" "$objType '$name' created successfully"

      ## Refresh data
      Refresh-Data -domain $Script:CurrentDomain
      Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
      [Terminal.Gui.Application]::RequestStop()
    }

  } catch {
    $errMsg = $_.Exception.Message
    Show-Modal "Error" "Failed to create $objType`:`n$errMsg"
  }
  }).GetNewClosure()

  $dlg.AddButton($btnCreate)
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
  $dlg.AddButton($btnCancel)
  [Terminal.Gui.Application]::Run($dlg)
}

# ----------[ Change Domain Dialog }----------
function Show-ChangeDomainDialog {
  <#
  .SYNOPSIS
  Dialog to change the current domain with fallback support

  .DESCRIPTION
  Allows user to change to a different domain. If the domain change fails,
  automatically falls back to the previous domain. Performs full reset and
  reload of all Script variables and tree view.
  #>

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
    Debug-Log (" OK pressed, Domain = $domainString") -Type "Insight"

    ## Close the dialog first
    [Terminal.Gui.Application]::RequestStop()

    ## Schedule the domain change after dialog closes
    [Terminal.Gui.Application]::MainLoop.Invoke({
      ## Save previous domain for fallback
      $Script:PreviousDomain = $Script:CurrentDomain
      Debug-Log (" Saved previous domain: $Script:PreviousDomain") -Type "Insight"

      try {
        Set-StatusBar "Changing domain to $domainString..." -Icon 'Working'

        ## Update the current domain
        $Script:CurrentDomain = $domainString
        $Script:Domain = $Script:CurrentDomain  ## Compatibility
        Debug-Log (" CurrentDomain set to: $Script:CurrentDomain") -Type "Insight"

        ## Reset all Script variables
        Debug-Log (" Resetting Script variables...") -Type "Insight"
        $Script:CurrentDC       = $null
        $Script:Users           = @()
        $Script:Groups          = @()
        $Script:DCs             = @()
        $Script:ADObjects       = @()
        $Script:SelectedObjects = @()
        $Script:SelectionMode   = $false

        ## Load domain data
        Set-StatusBar "Loading domain data for $($Script:CurrentDomain)..." -Icon 'Working'
        Debug-Log "Loading domain data for $($Script:CurrentDomain)..." -Type "Insight"

        Initialise-DataSource -Domain $Script:CurrentDomain

        Debug-Log "POST-LOAD: Users=$(${Script:Users}.Count), DCs=$(${Script:DCs}.Count), Computers=$(${Script:Computers}.Count), Groups=$(${Script:Groups}.Count), Objects=$(${Script:ADObjects}.Count)" -Type"Info"
        Debug-Log "Forest/Domain initialization complete: CurrentDomain=$($Script:CurrentDomain)" -Type "Insight"

        ## Build tree
        Set-StatusBar "Building tree..." -Icon 'Working'
        $rootNode = Build-Tree -domain $Script:CurrentDomain

        if ($null -ne $rootNode) {
          $Script:tree.ClearObjects()
          $Script:tree.AddObject($rootNode)
          Debug-Log "Root node added to TreeView" -Type "Success"
          Debug-Log "TreeView created and added to window successfully" -Type "Success"
          ## Update filter status if it exists
          if ($Script:FilterStatusLabel) {
            Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
          }
          Set-StatusBar "Ready" -Icon 'Success'
          Debug-Log (" Domain change successful to $domainString") -Type "Success"
        } else {
          throw "Build-Tree returned null root node"
        }
      } catch {
        Debug-Log (" Domain change error: $($_.Exception.Message)") -Type "Problem"
        Debug-Log (" Falling back to previous domain: $Script:PreviousDomain") -Type "Warning"

        ## Fallback to previous domain
        $Script:CurrentDomain = $Script:PreviousDomain
        $Script:Domain = $Script:CurrentDomain

        ## Reset variables
        $Script:CurrentDC       = $null
        $Script:Users           = @()
        $Script:Groups          = @()
        $Script:DCs             = @()
        $Script:ADObjects       = @()
        $Script:SelectedObjects = @()
        $Script:SelectionMode   = $false

        ## Reload previous domain
        try {
          Set-StatusBar "Restoring previous domain $Script:PreviousDomain..." -Icon 'Working'
          Initialise-DataSource -Domain $Script:CurrentDomain
          Show-InfoPanel -UpdateOnly
          $rootNode = Build-Tree -domain $Script:CurrentDomain

          if ($null -ne $rootNode) {
            $Script:tree.ClearObjects()
            $Script:tree.AddObject($rootNode)
            Debug-Log "Restored previous domain successfully" -Type "Success"
          }
        } catch {
          Debug-Log (" Failed to restore previous domain: $($_.Exception.Message)") -Type "Problem"
        }
        Set-StatusBar "Domain change failed - restored to $Script:PreviousDomain" -Icon 'Success'
        Show-Modal "Error" "Failed to load domain '$domainString'`n`nError: $($_.Exception.Message)`n`nRestored to previous domain: $Script:PreviousDomain"
      }
    })
  }).GetNewClosure()
  $dlg.Add($okBtn)

  $cancelBtn   = [Terminal.Gui.Button]::new("Cancel")
  $cancelBtn.X = 25
  $cancelBtn.Y = 4
  $cancelBtn.add_Clicked({
    Debug-Log (" Cancel pressed") -Type "Insight"
    [Terminal.Gui.Application]::RequestStop()
  }).GetNewClosure()
  $dlg.Add($cancelBtn)
  [Terminal.Gui.Application]::Run($dlg)
}

## ----------{ Change DC Dialog }----------
function Show-ChangeDCDialog {
  Debug-Log "Opening Change DC dialog" -Type "Insight"

  ## Capture functions for closures
  $debugLogFunc      = ${function:Debug-Log}
  $showInfoPanelFunc = ${function:Show-InfoPanel}
  $refreshDataFunc   = ${function:Refresh-Data}
  $showModalFunc     = { param($title, $msg) Show-Modal -Title $title -Message $msg }
  $demoMode          = $Script:DemoMode

  ## Get available DCs
  $availableDCs = if ($Script:DemoMode) {
    if ($Script:DCs) { $Script:DCs | Sort-Object Name } else { @() }
    } else {
      try { Get-ADDomainController -Filter * | Select-Object Name, Site, IPv4Address, IsGlobalCatalog, OperatingSystem }
      catch { & $debugLogFunc "Failed to get DCs: $($_.Exception.Message)" -Type "Problem"; @() }
    }
    if ($availableDCs.Count -eq 0) {
      & $showModalFunc "No DCs Found" "No Domain Controllers found"
      return
    }
    & $debugLogFunc "Found $($availableDCs.Count) DCs total" -Type "Insight"
    $dlg = [Terminal.Gui.Dialog]::new("Change Domain Controller", 90, 28)
    $dlg.Data = @{ SelectedDC = $null }
    $y = 1
    ## Header
    $lblHeader = [Terminal.Gui.Label]::new("Select Domain Controller:")
    $lblHeader.X = 2; $lblHeader.Y = $y
    $dlg.Add($lblHeader)
    $y += 2
    ## Current DC label - use the global variable
    $currentDCName = if ($Script:CurrentDCName) { $Script:CurrentDCName } else { "None" }
    $lblCurrent = [Terminal.Gui.Label]::new("Current: $currentDCName")
    $lblCurrent.X = 2; $lblCurrent.Y = $y
    $dlg.Add($lblCurrent)
    $y += 2
    ## ListView
    $lstDCs = [Terminal.Gui.ListView]::new()
    $lstDCs.X = 2; $lstDCs.Y = $y
    $lstDCs.Width = [Terminal.Gui.Dim]::Fill(2)
    $lstDCs.Height = 12
    $displayItems = @()
    foreach ($dc in $availableDCs) {
      $currentMarker = if ($dc.Name -eq $currentDCName) { "► " } else { "  " }
      $gcIcon = if ($dc.IsGlobalCatalog) { "🌐 " } else { "   " }
      $displayItems += "$currentMarker$gcIcon$($dc.Name.PadRight(18)) | Site: $($dc.Site ?? 'N/A') | IP: $($dc.IPv4Address ?? $dc.IPAddress ?? 'N/A')"
    }
    $lstDCs.SetSource($displayItems)

    ## Select current DC
    $currentIndex = ($availableDCs | ForEach-Object {$_.Name}).IndexOf($currentDCName)
    $lstDCs.SelectedItem = if ($currentIndex -ge 0) { $currentIndex } else { 0 }
    $dlg.Add($lstDCs)
    $y += 14

    ## Details label
    $lblDetails = [Terminal.Gui.Label]::new("Details:")
    $lblDetails.X = 2; $lblDetails.Y = $y
    $dlg.Add($lblDetails)
    $y += 1
    $lblDetailText = [Terminal.Gui.Label]::new("")
    $lblDetailText.X = 2; $lblDetailText.Y = $y
    $lblDetailText.Width = 84; $lblDetailText.Height = 3
    $dlg.Add($lblDetailText)

    ## Update details when selection changes
    $lstDCs.add_SelectedItemChanged({
      if ($lstDCs.SelectedItem -ge 0 -and $lstDCs.SelectedItem -lt $availableDCs.Count) {
        $selectedDC = $availableDCs[$lstDCs.SelectedItem]
        $detailText = "Name: $($selectedDC.Name)`nSite: $($selectedDC.Site ?? 'N/A') | IP: $($selectedDC.IPv4Address ?? $selectedDC.IPAddress ?? 'N/A') | GC: $(if ($selectedDC.IsGlobalCatalog) {'Yes'} else {'No'})`n"
        if ($selectedDC.OperatingSystem) { $detailText += "OS: $($selectedDC.OperatingSystem)" }
        $lblDetailText.Text = [NStack.ustring]::Make($detailText)
        $lblDetailText.SetNeedsDisplay()
      }
    })

    ## Connect button
    $btnConnect = [Terminal.Gui.Button]::new("Connect")
    $btnConnect.add_Clicked({
    if ($lstDCs.SelectedItem -ge 0 -and $lstDCs.SelectedItem -lt $availableDCs.Count) {
      $selectedDC = $availableDCs[$lstDCs.SelectedItem]
      $oldDCName = $Script:CurrentDCName  ## Capture the OLD name before changing
      $Script:CurrentDC = $selectedDC
      $Script:CurrentDCName = $selectedDC.Name
      & $debugLogFunc "Changed DC from '$oldDCName' to '$($Script:CurrentDCName)'" -Type "Success"
      & $showInfoPanelFunc -UpdateOnly
      if (-not $demoMode) { & $refreshDataFunc -Domain $Script:CurrentDomain -RebuildTree }
      [Terminal.Gui.Application]::RequestStop()
    } else {
    & $showModalFunc "No Selection" "Please select a Domain Controller"
    }
  })
  $dlg.AddButton($btnConnect)

  ## Cancel button
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })  # <-- NO .GetNewClosure() HERE!
  $dlg.AddButton($btnCancel)

  ## Run dialog
  [Terminal.Gui.Application]::Run($dlg)

  ## Refresh info panel after dialog closes
  & $showInfoPanelFunc -UpdateOnly  ## No -CurrentDC parameter!
  [Terminal.Gui.Application]::Top.SetNeedsDisplay()
  [Terminal.Gui.Application]::Top.Redraw([Terminal.Gui.Application]::Top.Bounds)
  [Terminal.Gui.Application]::Refresh()
  Debug-Log "InfoPanel update completed" -Type "Tracing"

  ## Refresh the tree since we changed DC
  Refresh-Data -Domain $Script:CurrentDomain -RebuildTree
}

## ----------{ AD Search Dialog }----------
## DSA-TUI Advanced Search Module v1.0
## Features: LDAP filters, saved searches, export results

## Global for saved searches
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

## Do the heavy lifting when searching AD
function Invoke-ADSearch {
  param(
    [Parameter(Mandatory=$true)] [Terminal.Gui.TextField] $UserField,
    [Parameter(Mandatory=$true)] [Terminal.Gui.TextField] $DomainField,
    [Parameter(Mandatory=$true)] [string] $ObjType,
    [Parameter(Mandatory=$true)] [Terminal.Gui.TabView] $TabView,
    [Parameter(Mandatory=$true)] [Terminal.Gui.TextView] $TxtOutput,
    [Parameter(Mandatory=$true)] [Terminal.Gui.CheckBox] $ChkDisabledOnly,
    [Parameter(Mandatory=$true)] [Terminal.Gui.TextView] $LdapFilter,
    [Parameter(Mandatory=$true)] [Terminal.Gui.TabView+Tab] $AdvTab,
    ## New parameters for clipboard buttons
    [Parameter(Mandatory=$false)] [Terminal.Gui.Button] $BtnCopyQuery,
    [Parameter(Mandatory=$false)] [Terminal.Gui.Button] $BtnPasteQuery,
    [Parameter(Mandatory=$false)] [Terminal.Gui.Button] $BtnCopyResults,
    ## NEW: Stale account parameters
    [Parameter(Mandatory=$false)] [Terminal.Gui.CheckBox] $ChkStaleOnly,
    [Parameter(Mandatory=$false)] [Terminal.Gui.TextField] $TxtDaysInactive,
    [Parameter(Mandatory=$false)] [Terminal.Gui.CheckBox] $ChkIncludeNeverLoggedOn
  )

  $searchName = $UserField.Text.ToString().Trim()
  $domain     = $DomainField.Text.ToString().Trim()
  $objType    = $ObjType
  $currentTab = $TabView.SelectedTab

  ## Parse stale account settings
  $staleOnly = if ($ChkStaleOnly) { $ChkStaleOnly.Checked } else { $false }
  $includeNeverLoggedOn = if ($ChkIncludeNeverLoggedOn) { $ChkIncludeNeverLoggedOn.Checked } else { $false }
  $daysInactive = 90
  if ($TxtDaysInactive -and $TxtDaysInactive.Text.ToString().Trim() -ne "") {
    if (-not [int]::TryParse($TxtDaysInactive.Text.ToString().Trim(), [ref]$daysInactive)) {
      $TxtOutput.Text = "Invalid days inactive value. Using default: 90"
      $daysInactive   = 90
    }
  }

  try {
    $objs = @()
    if ($currentTab -eq $AdvTab) {
      ## LDAP filter search
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
      ## Basic search
      if (-not $searchName -and -not $staleOnly) { $TxtOutput.Text="Please enter a name or enable stale account search."; return }
      if ($Script:DemoMode) {
        switch ($objType) {
          "User"  {
            $objs = $Script:Users | Where-Object {
              $nameMatch = if ($searchName) { $_.Name -like "*$searchName*" } else { $true }
              ## Apply stale filtering if enabled
              if ($staleOnly -and $nameMatch) {
                $cutoffDate = (Get-Date).AddDays(-$daysInactive)
                $lastLogon = $_.LastLogonDate
                if (-not $lastLogon) {
                  $includeNeverLoggedOn
                } else {
                  $lastLogon -lt $cutoffDate
                }
              } else {
                $nameMatch
              }
            } | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"user"}}, @{Name='LastLogon';Expression={if($_.LastLogonDate){$_.LastLogonDate.ToString('yyyy-MM-dd')}else{'Never'}}}
          }

          "Computer" {
            $objs = $Script:Computers | Where-Object {
              $nameMatch = if ($searchName) { $_.Name -like "*$searchName*" } else { $true }
              ## Apply stale filtering if enabled
              if ($staleOnly -and $nameMatch) {
                $cutoffDate = (Get-Date).AddDays(-$daysInactive)
                $lastLogon = $_.LastLogonDate
                if (-not $lastLogon) {
                  $includeNeverLoggedOn
                } else {
                  $lastLogon -lt $cutoffDate
                }
              } else {
                $nameMatch
              }
            } | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"computer"}}, @{Name='LastLogon';Expression={if($_.LastLogonDate){$_.LastLogonDate.ToString('yyyy-MM-dd')}else{'Never'}}}
          }
          "DomainController" {
            $objs = $Script:DCs | Where-Object {
              $nameMatch = if ($searchName) { $_.Name -like "*$searchName*" } else { $true }
              ## Apply stale filtering if enabled
              if ($staleOnly -and $nameMatch) {
                $cutoffDate = (Get-Date).AddDays(-$daysInactive)
                $lastLogon = $_.LastLogonDate
                if (-not $lastLogon) {
                  $includeNeverLoggedOn
                } else {
                  $lastLogon -lt $cutoffDate
                }
              } else {
                $nameMatch
              }
            } | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"domainController"}}, @{Name='LastLogon';Expression={if($_.LastLogonDate){$_.LastLogonDate.ToString('yyyy-MM-dd')}else{'Never'}}}
          }
          "Group" {
            if ($staleOnly) {
              $cutoffDate = (Get-Date).AddDays(-$daysInactive)
              $objs = $Script:Groups | Where-Object {
                $nameMatch = if ($searchName) { $_.Name -like "*$searchName*" } else { $true }
                if ($nameMatch) {
                  $whenChanged = $_.WhenChanged ?? $_.Modified
                  if (-not $whenChanged) {
                    $includeNeverLoggedOn
                  } else {
                    $whenChanged -lt $cutoffDate
                  }
                } else {
                  $false
                }
              } | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"group"}}, @{Name='LastModified';Expression={if($_.WhenChanged){$_.WhenChanged.ToString('yyyy-MM-dd')}else{'Never'}}}
            } else {
              $matchedGroups = @()
              foreach ($u in $Script:Users) {
                foreach ($g in $u.Groups) {
                  if ($g -like "*$searchName*") { $matchedGroups += $g }
                }
              }
              $objs = ($matchedGroups | Sort-Object -Unique) | ForEach-Object { [PSCustomObject]@{ Name=$_; Type="group" } }
            }
          }
          "OU"    {
            $ouNames = ($Script:Users | Select-Object -ExpandProperty OU -Unique)
            $objs = ($ouNames | Where-Object { $_ -like "*$searchName*" }) | ForEach-Object { [PSCustomObject]@{ Name=$_; Type="organizationalUnit" } }
          }
          "Contact"  { $objs = @() }
        }
        $Script:lastSearchType = if ($staleOnly) { "Stale $objType (${daysInactive}+ days)" } else { "Basic ($objType)" }
      } else {
        ## Real AD mode
        $loading = Show-LoadingDialog -Message "Searching AD for $objType '$searchName'..."
        try {
          $filterStr = if ($searchName) { "Name -like '*$searchName*'" } else { "*" }
          if ($ChkDisabledOnly.Checked -and $objType -eq "User") { $filterStr = if ($searchName) { "Name -like '*$searchName*' -and Enabled -eq $false" } else { "Enabled -eq $false" } }
          switch ($objType) {
            "User"     {
              $props = @('Enabled', 'LastLogonDate')
              $rawObjs = Get-ADUser -Filter $filterStr -Properties $props -ErrorAction Stop
              ## Apply stale filtering
              if ($staleOnly) {
                $cutoffDate = (Get-Date).AddDays(-$daysInactive)
                $rawObjs = $rawObjs | Where-Object {
                  $lastLogon = $_.LastLogonDate
                  if (-not $lastLogon) {
                    $includeNeverLoggedOn
                  } else {
                    $lastLogon -lt $cutoffDate
                  }
                }
              }
              $objs = $rawObjs | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"user"}}, @{Name='Enabled';Expression={$_.Enabled}}, @{Name='LastLogon';Expression={if($_.LastLogonDate){$_.LastLogonDate.ToString('yyyy-MM-dd')}else{'Never'}}}
            }
            "Computer" {
              $props = @('LastLogonDate')
              $rawObjs = Get-ADComputer -Filter $filterStr -Properties $props -ErrorAction Stop
              ## Apply stale filtering
              if ($staleOnly) {
                $cutoffDate = (Get-Date).AddDays(-$daysInactive)
                $rawObjs = $rawObjs | Where-Object {
                  $lastLogon = $_.LastLogonDate
                  if (-not $lastLogon) {
                    $includeNeverLoggedOn
                  } else {
                    $lastLogon -lt $cutoffDate
                  }
                }
              }
              $objs = $rawObjs | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"computer"}}, @{Name='LastLogon';Expression={if($_.LastLogonDate){$_.LastLogonDate.ToString('yyyy-MM-dd')}else{'Never'}}}
            }
            "Group"    {
              $props = @('Modified')
              $rawObjs = Get-ADGroup -Filter $filterStr -Properties $props -ErrorAction Stop
              ## Apply stale filtering
              if ($staleOnly) {
                $cutoffDate = (Get-Date).AddDays(-$daysInactive)
                $rawObjs = $rawObjs | Where-Object {
                  $modified = $_.Modified
                  if (-not $modified) {
                    $includeNeverLoggedOn
                  } else {
                    $modified -lt $cutoffDate
                  }
                }
              }
              $objs = $rawObjs | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"group"}}, @{Name='LastModified';Expression={if($_.Modified){$_.Modified.ToString('yyyy-MM-dd')}else{'Never'}}}
            }
            "OU"       { $objs = Get-ADOrganizationalUnit -Filter $filterStr -ErrorAction Stop | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"organizationalUnit"}} }
            "Contact"  { $objs = Get-ADObject -Filter "ObjectClass -eq 'contact' -and $filterStr" -Properties Name -ErrorAction Stop | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"contact"}} }
            default    { $objs=@() }
          }
          $Script:lastSearchType = if ($staleOnly) { "Stale $objType (${daysInactive}+ days)" } else { "Basic ($objType)" }
        } finally { Close-LoadingDialog $loading }
      }
    }

    ## Store results for export
    $Script:lastSearchResults = $objs
    if (-not $objs -or $objs.Count -eq 0) { $TxtOutput.Text = "No results found"; return }
    ## Display results
    $resultText = "Found $($objs.Count) object(s):`n`n"
    foreach ($obj in $objs) {
      $line = "$($obj.Name) [$($obj.Type)]"
      if ($obj.PSObject.Properties['LastLogon']) { $line += " - Last Logon: $($obj.LastLogon)" }
      if ($obj.PSObject.Properties['LastModified']) { $line += " - Last Modified: $($obj.LastModified)" }
      $resultText += "$line`n"
    }
    $TxtOutput.Text = $resultText
  } catch {
    $TxtOutput.Text = "Error: $($_.Exception.Message)"
  }
}

## ----------{ Clipboard Helper Functions }----------
function Copy-LDAPQueryToClipboard {
  param([Terminal.Gui.TextView] $LdapFilter)

  try {
    $query = $LdapFilter.Text.ToString().Trim()
    if (-not $query) {
      Show-Modal "Info" "No LDAP query to copy"
      return
    }
    Set-Clipboard -Value $query
    Debug-Log "Copied LDAP query to clipboard" -Type "Success"
    Show-Modal "Success" "LDAP query copied to clipboard"
  } catch {
    Debug-Log "Failed to copy to clipboard: $($_.Exception.Message)" -Type "Problem"
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
    Debug-Log "Pasted LDAP query from clipboard" -Type "Success"
    [Terminal.Gui.Application]::Refresh()
  } catch {
    Debug-Log "Failed to paste from clipboard: $($_.Exception.Message)" -Type "Problem"
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
    ## Format results for clipboard
    $clipboardText = "# AD Search Results - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
    $clipboardText += "# Search Type: $($Script:lastSearchType)`n"
    $clipboardText += "# Total Results: $($Script:lastSearchResults.Count)`n"
    $clipboardText += "`n"
    ## Add results in multiple formats for flexibility
    ## Format 1: Simple list
    $clipboardText += "========== Simple List ==========`n"
    foreach ($obj in $Script:lastSearchResults) { $clipboardText += "$($obj.Name) [$($obj.Type)]`n" }
    $clipboardText += "`n========== CSV Format ==========`n"
    $clipboardText += "Name,Type"
    if ($Script:lastSearchResults[0].PSObject.Properties['DN']) { $clipboardText += ",DistinguishedName" }
    if ($Script:lastSearchResults[0].PSObject.Properties['Enabled']) { $clipboardText += ",Enabled" }
    $clipboardText += "`n"
    ## Format 2: CSV
    foreach ($obj in $Script:lastSearchResults) {
      $clipboardText += "`"$($obj.Name)`",`"$($obj.Type)`""
      if ($obj.PSObject.Properties['DN']) { $clipboardText += ",`"$($obj.DN)`"" }
      if ($obj.PSObject.Properties['Enabled']) { $clipboardText += ",`"$($obj.Enabled)`"" }
      $clipboardText += "`n"
    }
    ## Format 3: PowerShell array (if useful)
    $clipboardText += "`n========== Powershell Names ==========`n"
    $clipboardText += "@(`n"
    $names = $Script:lastSearchResults | ForEach-Object { '   "' + $_.Name + '"' }
    $clipboardText += ($names -join ",`n")
    $clipboardText += "`n)`n"
    Set-Clipboard -Value $clipboardText
    Debug-Log "Copied $($Script:lastSearchResults.Count) search results to clipboard" -Type "Success"
    Show-Modal "Success" "Copied $($Script:lastSearchResults.Count) results to clipboard`n`nFormats included:`n• Simple list`n• CSV`n• PowerShell array"
  } catch {
    Debug-Log "Failed to copy results to clipboard: $($_.Exception.Message)" -Type "Problem"
    Show-Modal "Error" "Failed to copy results to clipboard:`n$($_.Exception.Message)"
  }
}

function Show-ADSearchDialog {
  <#
    .SYNOPSIS
    Displays an AD Search dialog with Basic and Advanced LDAP search capabilities

    .DESCRIPTION
    Creates a modal dialog for searching Active Directory with:
    - Basic search (by name and object type)
    - Advanced LDAP filter search
    - Stale account search (inactive for X days)
    - Clipboard support for queries and results
    - Export functionality
  #>

  Debug-Log "Opening AD Search Dialog" -Type "Insight"

  ## Capture functions for closures
  $debugLogFunc          = ${function:Debug-Log}
  $showModalFunc         = ${function:Show-Modal}
  $copySearchResultsFunc = ${function:Copy-SearchResultsToClipboard}
  $copyLDAPQueryFunc     = ${function:Copy-LDAPQueryToClipboard}
  $pasteLDAPQueryFunc    = ${function:Paste-LDAPQueryFromClipboard}
  $invokeADSearchFunc    = ${function:Invoke-ADSearch}

  try {
    ## ----------{ Buttons }----------
    $btnClose = [Terminal.Gui.Button]::new("Close")
    $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }.GetNewClosure())

    ## ----------{ Dialog }----------
    $dlg = [Terminal.Gui.Dialog]::new("Active Directory Search", 120, 40, $btnClose)

    ## ----------{ TabView for Basic/Advanced }----------
    $searchTabView = [Terminal.Gui.TabView]::new()
    $searchTabView.X = 0
    $searchTabView.Y = 0
    $searchTabView.Width = [Terminal.Gui.Dim]::Fill()
    $searchTabView.Height = [Terminal.Gui.Dim]::Fill(1)

    ## ----------{ Basic Search Tab }---------
    $basicTab = [Terminal.Gui.TabView+Tab]::new()
    $basicTab.Text = "Basic Search"
    $basicView = [Terminal.Gui.View]::new()
    $basicView.X = 0; $basicView.Y = 0
    $basicView.Width = [Terminal.Gui.Dim]::Fill()
    $basicView.Height = [Terminal.Gui.Dim]::Fill()
    $y = 1

    ## Domain
    $lbl = [Terminal.Gui.Label]::new("Domain:"); $lbl.X=2; $lbl.Y=$y; $basicView.Add($lbl)
    $txtDomain = [Terminal.Gui.TextField]::new($Script:CurrentDomain ?? ""); $txtDomain.X=20; $txtDomain.Y=$y; $txtDomain.Width=40
    $basicView.Add($txtDomain); $y+=2

    ## Search Name
    $lbl = [Terminal.Gui.Label]::new("Name:"); $lbl.X=2; $lbl.Y=$y; $basicView.Add($lbl)
    $txtSearchName = [Terminal.Gui.TextField]::new(""); $txtSearchName.X=20; $txtSearchName.Y=$y; $txtSearchName.Width=40
    $basicView.Add($txtSearchName); $y+=2

    ## Object Type
    $lbl = [Terminal.Gui.Label]::new("Object Type:"); $lbl.X=2; $lbl.Y=$y; $basicView.Add($lbl)
    $cmbObjectType = [Terminal.Gui.ComboBox]::new(); $cmbObjectType.X=20; $cmbObjectType.Y=$y; $cmbObjectType.Width=20
    $cmbObjectType.SetSource(@("User", "Group", "Computer", "OU", "Contact", "DomainController"))
    $cmbObjectType.SelectedItem = 0
    $basicView.Add($cmbObjectType); $y+=2

    ## Disabled only checkbox (for users)
    $chkDisabledOnly = [Terminal.Gui.CheckBox]::new("Disabled accounts only"); $chkDisabledOnly.X=20; $chkDisabledOnly.Y=$y
    $basicView.Add($chkDisabledOnly); $y+=1

    ## ----------{ Stale Account filtering }---------
    ## Stale accounts checkbox
    $chkStaleOnly = [Terminal.Gui.CheckBox]::new("Stale accounts only (inactive)"); $chkStaleOnly.X=20; $chkStaleOnly.Y=$y
    $basicView.Add($chkStaleOnly); $y+=2

    ## Days inactive field
    $lbl = [Terminal.Gui.Label]::new("Days inactive:"); $lbl.X=20; $lbl.Y=$y; $basicView.Add($lbl)
    $txtDaysInactive = [Terminal.Gui.TextField]::new("90"); $txtDaysInactive.X=36; $txtDaysInactive.Y=$y; $txtDaysInactive.Width=10
    $basicView.Add($txtDaysInactive)

    $lblDaysHelp = [Terminal.Gui.Label]::new("(Default: 90)"); $lblDaysHelp.X=47; $lblDaysHelp.Y=$y
    $basicView.Add($lblDaysHelp); $y+=1

    ## Include never logged on checkbox
    $chkIncludeNeverLoggedOn = [Terminal.Gui.CheckBox]::new("Include never logged on"); $chkIncludeNeverLoggedOn.X=20; $chkIncludeNeverLoggedOn.Y=$y
    $basicView.Add($chkIncludeNeverLoggedOn); $y+=2

    ## Search button (with Copy Results on same line)
    $btnSearch = [Terminal.Gui.Button]::new("Search"); $btnSearch.X=20; $btnSearch.Y=$y
    $basicView.Add($btnSearch)

    ## Copy Results button on same line
    $btnCopyResults = [Terminal.Gui.Button]::new("Copy Results to Clipboard")
    $btnCopyResults.X = 35
    $btnCopyResults.Y = $y
    $btnCopyResults.add_Clicked({
      & $copySearchResultsFunc -TxtOutput $txtResults
    }.GetNewClosure())
    $basicView.Add($btnCopyResults)
    $y+=2  ## One line space below buttons

    ## Results
    $lbl = [Terminal.Gui.Label]::new("Results:"); $lbl.X=2; $lbl.Y=$y; $basicView.Add($lbl); $y+=1
    $txtResults = [Terminal.Gui.TextView]::new()
    $txtResults.X = 2
    $txtResults.Y = $y
    $txtResults.Width = [Terminal.Gui.Dim]::Fill(2)
    $txtResults.Height = [Terminal.Gui.Dim]::Fill(1)  ## Go to bottom -1 line
    $txtResults.ReadOnly = $true
    $txtResults.Text = "Enter search criteria and click Search"
    $basicView.Add($txtResults)
    $basicTab.View = $basicView
    $searchTabView.AddTab($basicTab, $false)

    ## ----------{ Advanced LDAP Tab }---------
    $advTab = [Terminal.Gui.TabView+Tab]::new()
    $advTab.Text = "Advanced (LDAP)"
    $advView = [Terminal.Gui.View]::new()
    $advView.X = 0; $advView.Y = 0
    $advView.Width = [Terminal.Gui.Dim]::Fill()
    $advView.Height = [Terminal.Gui.Dim]::Fill()
    $y = 1

    ## Domain
    $lbl = [Terminal.Gui.Label]::new("Domain:"); $lbl.X=2; $lbl.Y=$y; $advView.Add($lbl)
    $txtDomainAdv = [Terminal.Gui.TextField]::new($Script:CurrentDomain ?? ""); $txtDomainAdv.X=20; $txtDomainAdv.Y=$y; $txtDomainAdv.Width=40
    $advView.Add($txtDomainAdv); $y+=2

    ## LDAP Filter
    $lbl = [Terminal.Gui.Label]::new("LDAP Filter:"); $lbl.X=2; $lbl.Y=$y; $advView.Add($lbl); $y+=1

    ## Example filters label
    $lblExamples = [Terminal.Gui.Label]::new("Examples: (&(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))");
    $lblExamples.X=2; $lblExamples.Y=$y; $lblExamples.ColorScheme = [Terminal.Gui.Colors]::TopLevel
    $advView.Add($lblExamples); $y+=1

    $txtLdapFilter = [Terminal.Gui.TextView]::new()
    $txtLdapFilter.X = 2
    $txtLdapFilter.Y = $y
    $txtLdapFilter.Width = [Terminal.Gui.Dim]::Fill(2)
    $txtLdapFilter.Height = 5
    $advView.Add($txtLdapFilter); $y+=6

    ## Clipboard buttons for LDAP query
    $btnCopyQuery = [Terminal.Gui.Button]::new("Copy Query"); $btnCopyQuery.X=2; $btnCopyQuery.Y=$y
    $btnCopyQuery.add_Clicked({
      & $copyLDAPQueryFunc -LdapFilter $txtLdapFilter
    }.GetNewClosure())
    $advView.Add($btnCopyQuery)

    $btnPasteQuery = [Terminal.Gui.Button]::new("Paste Query"); $btnPasteQuery.X=18; $btnPasteQuery.Y=$y
    $btnPasteQuery.add_Clicked({
      & $pasteLDAPQueryFunc -LdapFilter $txtLdapFilter
    }.GetNewClosure())
    $advView.Add($btnPasteQuery); $y+=2

    ## Execute LDAP Query button (with Copy Results on same line)
    $btnSearchAdv = [Terminal.Gui.Button]::new("Execute LDAP Query"); $btnSearchAdv.X=2; $btnSearchAdv.Y=$y
    $advView.Add($btnSearchAdv)

    ## Copy Results button on same line
    $btnCopyResultsAdv = [Terminal.Gui.Button]::new("Copy Results to Clipboard")
    $btnCopyResultsAdv.X = 25
    $btnCopyResultsAdv.Y = $y
    $btnCopyResultsAdv.add_Clicked({
      & $copySearchResultsFunc -TxtOutput $txtResultsAdv
    }.GetNewClosure())
    $advView.Add($btnCopyResultsAdv)
    $y+=2  ## One line space below buttons

    ## Results
    $lbl = [Terminal.Gui.Label]::new("Results:"); $lbl.X=2; $lbl.Y=$y; $advView.Add($lbl); $y+=1
    $txtResultsAdv = [Terminal.Gui.TextView]::new()
    $txtResultsAdv.X = 2
    $txtResultsAdv.Y = $y
    $txtResultsAdv.Width = [Terminal.Gui.Dim]::Fill(2)
    $txtResultsAdv.Height = [Terminal.Gui.Dim]::Fill(1)  ## Go to bottom -1 line
    $txtResultsAdv.ReadOnly = $true
    $txtResultsAdv.Text = "Enter LDAP filter and click Execute"
    $advView.Add($txtResultsAdv)
    $advTab.View = $advView
    $searchTabView.AddTab($advTab, $false)

    ## ----------{ Wire up Search Buttons }---------
    ## Basic Search
    $btnSearch.add_Clicked({
      ## Get the selected object type as a string
      $selectedType = $cmbObjectType.Text.ToString()

      & $invokeADSearchFunc -UserField $txtSearchName -DomainField $txtDomain -ObjType $selectedType -TabView $searchTabView -TxtOutput $txtResults -ChkDisabledOnly $chkDisabledOnly -LdapFilter $txtLdapFilter -AdvTab $advTab -ChkStaleOnly $chkStaleOnly -TxtDaysInactive $txtDaysInactive -ChkIncludeNeverLoggedOn $chkIncludeNeverLoggedOn
    }.GetNewClosure())

    ## Advanced Search
    $btnSearchAdv.add_Clicked({
      ## Get the selected object type as a string
      $selectedType = $cmbObjectType.Text.ToString()
      & $invokeADSearchFunc -UserField $txtSearchName -DomainField $txtDomainAdv -ObjType $selectedType -TabView $searchTabView -TxtOutput $txtResultsAdv -ChkDisabledOnly $chkDisabledOnly -LdapFilter $txtLdapFilter -AdvTab $advTab -ChkStaleOnly $chkStaleOnly -TxtDaysInactive $txtDaysInactive -ChkIncludeNeverLoggedOn $chkIncludeNeverLoggedOn
    }.GetNewClosure())

    ## Add TabView to dialog
    $dlg.Add($searchTabView)
    & $debugLogFunc "AD Search Dialog created, running" -Type "Success"

    ## Run the dialog
    [Terminal.Gui.Application]::Run($dlg)
    & $debugLogFunc "AD Search Dialog closed" -Type "Insight"

  } catch {
    & $debugLogFunc "Exception in Show-ADSearchDialog: $($_.Exception.Message)" -Type "Problem"
    & $showModalFunc "Error" "Failed to open search dialog:`n$($_.Exception.Message)"
  }
}

function Show-DCPropertiesDialog {
  param($dc)

  ## Accept either DC object or DC name
  if ($dc -is [string]) {
    $dcName = $dc
    Debug-Log "Looking for DC: $dcName" -Type "Insight"

    if (-not $Script:DCs) {
      Debug-Log "Script:DCs is null or not Initialised" -Type "Problem"
      Show-Modal "Error" "Domain Controllers list is not loaded"
      return
    }

    $dc = $Script:DCs | Where-Object { $_.Name -eq $dcName } | Select-Object -First 1

    if (-not $dc) {
      Debug-Log "DC '$dcName' not found in Script:DCs" -Type "Problem"
      Show-Modal "Not Found" "DC '$dcName' not found in the domain controllers list"
      return
    }
    Debug-Log "Found DC: $($dc.Name)" -Type "Success"
  }

  if (-not $dc) {
    Debug-Log "DC object is null after lookup" -Type "Problem"
    Show-Modal "Error" "DC object is null"
    return
  }

  Debug-Log "Showing DC properties for: $($dc.Name)" -Type "Insight"

  ## ----------{ General Tab }---------
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Domain Controller Information"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Name:" -FieldName 'lblName' -State $state -Value ($dc.Name ?? "Unknown") -FieldX 25 -IsTextField $false

      $hostname = if ($dc.HostName) { $dc.HostName } elseif ($dc.DNSHostName) { $dc.DNSHostName } else { $dc.Name }
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Hostname:" -FieldName 'lblHostname' -State $state -Value $hostname -FieldX 25 -IsTextField $false

      if ($dc.Site)     { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Site:" -FieldName 'lblSite' -State $state -Value $dc.Site -FieldX 25 -IsTextField $false }
      if ($dc.Domain)   { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Domain:" -FieldName 'lblDomain' -State $state -Value $dc.Domain -FieldX 25 -IsTextField $false }
      if ($dc.Forest)   { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Forest:" -FieldName 'lblForest' -State $state -Value $dc.Forest -FieldX 25 -IsTextField $false }
      if ($dc.Location) { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Location:" -FieldName 'lblLocation' -State $state -Value $dc.Location -FieldX 25 -IsTextField $false }
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Network"
      if ($dc.IPv4Address -or $dc.IPAddress) {
        $ip = if ($dc.IPv4Address) { $dc.IPv4Address } else { $dc.IPAddress }
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "IPv4 Address:" -FieldName 'lblIPv4' -State $state -Value $ip -FieldX 25 -IsTextField $false
      }
      if ($dc.IPv6Address) { Add-LabelAndField -View $view -Y ([ref]$y) -Label "IPv6 Address:" -FieldName 'lblIPv6' -State $state -Value $dc.IPv6Address -FieldX 25 -IsTextField $false }
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Operating System"
      if ($dc.OperatingSystem -or $dc.OS) {
        $os = if ($dc.OperatingSystem) { $dc.OperatingSystem } else { $dc.OS }
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "OS:" -FieldName 'lblOS' -State $state -Value $os -FieldX 25 -IsTextField $false
      }
      if ($dc.OperatingSystemVersion) { Add-LabelAndField -View $view -Y ([ref]$y) -Label "OS Version:" -FieldName 'lblOSVer' -State $state -Value $dc.OperatingSystemVersion -FieldX 25 -IsTextField $false }
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Capabilities"
      $isGC = if ($null -ne $dc.IsGlobalCatalog) { $dc.IsGlobalCatalog.ToString() } else { "Unknown" }
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Global Catalog:" -FieldName 'lblGC' -State $state -Value $isGC -FieldX 25 -IsTextField $false
      if ($null -ne $dc.IsReadOnly) { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Read-Only:" -FieldName 'lblRO' -State $state -Value $dc.IsReadOnly.ToString() -FieldX 25 -IsTextField $false }
      if ($null -ne $dc.Enabled) { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Enabled:" -FieldName 'lblEnabled' -State $state -Value $dc.Enabled.ToString() -FieldX 25 -IsTextField $false }
    }
  }

  ## ----------{ Roles Tab }---------
  $rolesTab = @{
    Name = "Roles"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "FSMO Roles"
      $fsmoRoles = @()
      if ($dc.PSObject.Properties['FSMORoles'] -and $dc.FSMORoles -and $dc.FSMORoles.Count -gt 0) { $fsmoRoles = $dc.FSMORoles
      } elseif ($dc.PSObject.Properties['OperationMasterRoles'] -and $dc.OperationMasterRoles -and $dc.OperationMasterRoles.Count -gt 0) { $fsmoRoles = $dc.OperationMasterRoles }

      if ($fsmoRoles.Count -gt 0) {
        foreach ($role in $fsmoRoles) {
          $lbl = [Terminal.Gui.Label]::new("• $role")
          $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
        }
      } else {
        $lbl = [Terminal.Gui.Label]::new("(None)")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
    }
  }

  ## ----------{ Replication Tab }---------
  $replicationTab = @{
    Name = "Replication"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Replication Status"

      if ($dc.ReplicationHealth) { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Health:" -FieldName 'lblHealth' -State $state -Value $dc.ReplicationHealth -FieldX 25 -IsTextField $false }
      if ($dc.LastReplication) {
        $lastRep = $dc.LastReplication.ToString('yyyy-MM-dd HH:mm:ss')
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Last Replication:" -FieldName 'lblLastRep' -State $state -Value $lastRep -FieldX 25 -IsTextField $false
        $y += 1
      }
      if ($dc.ReplicationPartners -and $dc.ReplicationPartners.Count -gt 0) {
        $lbl = [Terminal.Gui.Label]::new("Replication Partners:")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
        foreach ($partner in $dc.ReplicationPartners) {
          $lbl = [Terminal.Gui.Label]::new("• $partner")
          $lbl.X=6; $lbl.Y=$y; $view.Add($lbl); $y+=1
        }
      }
    }
  }

  ## ----------{ Services Tab }---------
  $servicesTab = @{
    Name = "Services"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      if ($dc.Services) {
        Add-SectionHeader -View $view -Y ([ref]$y) -Text "Service Status"
        foreach ($service in $dc.Services.Keys | Sort-Object) {
          $status = $dc.Services[$service]
          $lbl = [Terminal.Gui.Label]::new("${service}:")
          $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
          $lbl = [Terminal.Gui.Label]::new($status)
          $lbl.X=20; $lbl.Y=$y; $view.Add($lbl)
          $y+=1
        }
        $y+=1
      }
      if ($dc.LastBoot -or $dc.LastBootUpTime) {
        Add-SectionHeader -View $view -Y ([ref]$y) -Text "System Information" -SpaceBefore 0
        $lastBoot = if ($dc.LastBootUpTime) { $dc.LastBootUpTime } else { $dc.LastBoot }
        if ($lastBoot) {
          $bootTime = $lastBoot.ToString('yyyy-MM-dd HH:mm:ss')
          $uptime = (Get-Date) - $lastBoot
          $uptimeStr = "$($uptime.Days) days, $($uptime.Hours) hours"
          $lbl = [Terminal.Gui.Label]::new("Last Boot:")
          $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
          $lbl = [Terminal.Gui.Label]::new($bootTime)
          $lbl.X=20; $lbl.Y=$y; $view.Add($lbl); $y+=1
          $lbl = [Terminal.Gui.Label]::new("Uptime:")
          $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
          $lbl = [Terminal.Gui.Label]::new($uptimeStr)
          $lbl.X=20; $lbl.Y=$y; $view.Add($lbl); $y+=1
        }
      }
    }
  }

  ## ----------{ Disk Space Tab }---------
  $diskTab = @{
    Name = "Disk Space"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      if ($dc.DiskSpace) {
        Add-SectionHeader -View $view -Y ([ref]$y) -Text "Disk Usage"
        foreach ($drive in $dc.DiskSpace.Keys | Sort-Object) {
          $diskInfo = $dc.DiskSpace[$drive]
          $lbl = [Terminal.Gui.Label]::new("Drive ${drive}")
          $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
          if ($diskInfo.Total) {
            $lbl = [Terminal.Gui.Label]::new("  Total:")
            $lbl.X=6; $lbl.Y=$y; $view.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($diskInfo.Total)
            $lbl.X=20; $lbl.Y=$y; $view.Add($lbl); $y+=1
          }
          if ($diskInfo.Used) {
            $lbl = [Terminal.Gui.Label]::new("  Used:")
            $lbl.X=6; $lbl.Y=$y; $view.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($diskInfo.Used)
            $lbl.X=20; $lbl.Y=$y; $view.Add($lbl); $y+=1
          }
          if ($diskInfo.Free) {
            $lbl = [Terminal.Gui.Label]::new("  Free:")
            $lbl.X=6; $lbl.Y=$y; $view.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new($diskInfo.Free)
            $lbl.X=20; $lbl.Y=$y; $view.Add($lbl); $y+=1
          }
          if ($null -ne $diskInfo.PercentFree) {
            $lbl = [Terminal.Gui.Label]::new("  % Free:")
            $lbl.X=6; $lbl.Y=$y; $view.Add($lbl)
            $lbl = [Terminal.Gui.Label]::new("$($diskInfo.PercentFree)%")
            $lbl.X=20; $lbl.Y=$y; $view.Add($lbl); $y+=1
          }
          $y+=1
        }
      } else {
        $lbl = [Terminal.Gui.Label]::new("(No disk space information available)")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
    }
  }

  ## ----------{ RID Pool Tab }---------
  $ridPoolTab = @{
    Name = "RID Pool"
    Builder = {
      param($view, $dc, $state)
      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Relative Identifier (RID) Pool Status"
      $ridSet = $null
      if ($Script:rawRIDSets) {
        $ridSet = $Script:rawRIDSets | Where-Object {
          $_.DN -match [regex]::Escape($dc.Name)
        } | Select-Object -First 1
      }
      if ($ridSet) {
        $lbl = [Terminal.Gui.Label]::new("RID Set Information:")
        $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 2

        if ($ridSet.RIDAvailablePool)  { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Available Pool:" -FieldName 'lblRIDAvail' -State $state -Value $ridSet.RIDAvailablePool.ToString() -LabelX 6 -FieldX 30 -IsTextField $false }
        if ($ridSet.RIDAllocationPool) { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Allocation Pool:" -FieldName 'lblRIDAlloc' -State $state -Value $ridSet.RIDAllocationPool.ToString() -LabelX 6 -FieldX 30 -IsTextField $false }
        if ($ridSet.RIDPreviousAllocationPool) { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Previous Pool:" -FieldName 'lblRIDPrev' -State $state -Value $ridSet.RIDPreviousAllocationPool.ToString() -LabelX 6 -FieldX 30 -IsTextField $false }
        if ($ridSet.RIDUsedPool) { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Used Pool:" -FieldName 'lblRIDUsed' -State $state -Value $ridSet.RIDUsedPool.ToString() -LabelX 6 -FieldX 30 -IsTextField $false }
        if ($ridSet.RIDNextRID) { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Next RID:" -FieldName 'lblRIDNext' -State $state -Value $ridSet.RIDNextRID.ToString() -LabelX 6 -FieldX 30 -IsTextField $false }

        $y += 1
        Add-SectionHeader -View $view -Y ([ref]$y) -Text "About RID Pools" -SpaceBefore 0

        $info = @(
          "• RIDs are allocated in pools of 500 by default",
          "• Each DC requests a new pool when current pool is exhausted",
          "• RID Master FSMO role holder allocates pools to DCs",
          "• Monitor RID pool usage to prevent RID exhaustion"
        )

        foreach ($line in $info) {
          $lbl = [Terminal.Gui.Label]::new($line)
          $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 1
        }

      } else {
        $lbl = [Terminal.Gui.Label]::new("No RID Set information found for this DC")
        $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 2
        $lbl = [Terminal.Gui.Label]::new("This may indicate:")
        $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 1

        $reasons = @(
          "  • DC is an RODC (Read-Only DC)",
          "  • RID data not included in CSV export",
          "  • Use 'dcdiag /test:ridmanager' for diagnostics"
        )

        foreach ($reason in $reasons) {
          $lbl = [Terminal.Gui.Label]::new($reason)
          $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 1
        }
      }
    }
  }

  ## ----------{ DFSR Tab }---------
  $dfsrTab = @{
    Name = "DFSR"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "DFS Replication Status"

      $dfsrObjects = @()
      if ($Script:rawDFSR) {
        $dfsrObjects = $Script:rawDFSR | Where-Object {
          $_.DN -match [regex]::Escape($dc.Name)
        }
      }

      if ($dfsrObjects.Count -gt 0) {
        $lbl = [Terminal.Gui.Label]::new("Found $($dfsrObjects.Count) DFSR object(s) for this DC:")
        $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 2

        $state.lstDFSR = [Terminal.Gui.ListView]::new()
        $state.lstDFSR.X = 4; $state.lstDFSR.Y = $y
        $state.lstDFSR.Width = [Terminal.Gui.Dim]::Fill(2)
        $state.lstDFSR.Height = 15

        $dfsrList = $dfsrObjects | ForEach-Object { "$($_.Class): $($_.Name)" }
        $state.lstDFSR.SetSource($dfsrList)
        $view.Add($state.lstDFSR)
        $y += 16

        $btnViewDFSR = [Terminal.Gui.Button]::new("View Details...")
        $btnViewDFSR.X = 4; $btnViewDFSR.Y = $y
        $btnViewDFSR.add_Clicked({
          $selectedIndex = $state.lstDFSR.SelectedItem
          if ($selectedIndex -ge 0 -and $selectedIndex -lt $dfsrObjects.Count) {
            $dfsrObj = $dfsrObjects[$selectedIndex]
            $details = @"
DFSR Object Details
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Class: $($dfsrObj.Class)
Name:  $($dfsrObj.Name)
DN:    $($dfsrObj.DN)

This DFSR object is part of the Active Directory
replication topology for this domain controller.
"@
            Show-Modal "DFSR Details" $details
          } else {
            Show-Modal "Info" "Please select a DFSR object to view"
          }
        }.GetNewClosure()).GetNewClosure()
        $view.Add($btnViewDFSR)

      } else {
        $lbl = [Terminal.Gui.Label]::new("No DFSR objects found for this DC")
        $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 2
        $lbl = [Terminal.Gui.Label]::new("This may indicate:")
        $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 1

        $reasons = @(
          "  • DFSR is not configured on this DC",
          "  • DFSR objects not included in CSV export",
          "  • Use 'dfsrdiag' or 'Get-DfsrBacklog' for diagnostics"
        )

        foreach ($reason in $reasons) {
          $lbl = [Terminal.Gui.Label]::new($reason)
          $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 1
        }
      }
    }
  }

  ## ----------{ Create Dialog }---------
  ## Note: Search tab auto-added with DC-specific checkboxes
  $tabs = @($generalTab, $rolesTab, $replicationTab, $servicesTab, $diskTab, $ridPoolTab, $dfsrTab)
  Debug-Log "All DC tabs added, running dialog" -Type "Success"
  New-PropertiesDialog -Title "Domain Controller Properties - $($dc.Name)" -Width 110 -Height 35 -Tabs $tabs -Data $dc -IncludeSearchTab $true -SearchTabConfig @{ ObjectType='DomainController'; SearchTypes=@("Domain Controller","Computer","OU") }
  Debug-Log "DC dialog closed normally" -Type "Insight"
}

## Move/Copy/Delete objects function
function Invoke-ObjectOperation {
  <#
  .SYNOPSIS
  Unified function for moving or deleting AD objects

  .PARAMETER Objects
  Object or array of objects to operate on

  .PARAMETER Operation
  'Move' or 'Delete'

  .PARAMETER ObjectType
  Optional explicit type: 'User', 'Group', 'Computer', 'OU', 'Domain Controller'

  .PARAMETER IsBulk
  If true, shows bulk UI (for multiple objects)
  #>

  param(
    [Parameter(Mandatory=$true)]
    [object[]]$Objects,
    [Parameter(Mandatory=$true)]
    [ValidateSet('Move', 'Delete')]
    [string]$Operation,
    [Parameter(Mandatory=$false)]
    [ValidateSet('User', 'Group', 'Computer', 'OU', 'Domain Controller')]
    [string]$ObjectType,
    [switch]$IsBulk
  )

  ## Convert single object to array if needed
  if ($Objects -isnot [array]) { $Objects = @($Objects) }
  ## Auto-detect object type if not provided
  if (-not $ObjectType) {
    $obj = $Objects[0]
    ## Check if this is a tree node wrapper (has Object, Type, Name keys)
    if ($obj -is [hashtable] -and $obj.ContainsKey('Type') -and $obj.ContainsKey('Object')) {
      Debug-Log "Detected tree node wrapper! Type: $($obj['Type']), Name: $($obj['Name'])" -Type "Insight"
      $ObjectType = $obj['Type']
      ## Unwrap all objects in the array
      for ($i = 0; $i -lt $Objects.Count; $i++) {
        if ($Objects[$i] -is [hashtable] -and $Objects[$i].ContainsKey('Object')) {
          $Objects[$i] = $Objects[$i]['Object']
        }
      }
      Debug-Log "Unwrapped $($Objects.Count) object(s)" -Type "Tracing"
    }
  }

  ## If still no type, try to detect from the actual object
  if (-not $ObjectType) {
    $obj = $Objects[0]
    ## Handle hashtable objects
    if ($obj -is [hashtable]) {
      Debug-Log "Object is a hashtable, keys: $($obj.Keys -join ', ')" -Type "Tracing"
      ## Try ObjectClass first
      if ($obj.ContainsKey('ObjectClass') -and $obj['ObjectClass']) {
        $ObjectType = switch ($obj['ObjectClass'].ToLower()) {
          'user'     { 'User'  }
          'group'    { 'Group' }
          'computer' { 'Computer' }
          'organizationalunit' { 'OU' }
          default { $null }
        }
        if ($ObjectType) { Debug-Log "Detected type from ObjectClass: $ObjectType" -Type "Tracing" }
      }

      ## Fallback to property-based detection for hashtables
      if (-not $ObjectType) {
        if ($obj.ContainsKey('Site') -or $obj.ContainsKey('IsGlobalCatalog')) {
          $ObjectType = 'Domain Controller'
        }
        elseif ($obj.ContainsKey('ComputerType') -or $obj.ContainsKey('OperatingSystem')) {
          $ObjectType = 'Computer'
        }
        elseif ($obj.ContainsKey('Members') -or $obj.ContainsKey('GroupType')) {
          $ObjectType = 'Group'
        }
        elseif ($obj.ContainsKey('distinguishedName') -and $obj['distinguishedName'] -match '^OU=') {
          $ObjectType = 'OU'
        }
        elseif ($obj.ContainsKey('SamAccountName') -or $obj.ContainsKey('UserPrincipalName')) {
          $ObjectType = 'User'
        }
      }
    }
    else {
      ## Handle PSCustomObject
      Debug-Log "Object is PSCustomObject, properties: $($obj.PSObject.Properties.Name -join ', ')" -Type "Tracing"
      ## Try ObjectClass first (most reliable)
      if ($obj.PSObject.Properties['ObjectClass'] -and $obj.ObjectClass) {
        $ObjectType = switch ($obj.ObjectClass.ToLower()) {
          'user'      { 'User' }
          'group'    { 'Group' }
          'computer' { 'Computer' }
          'organizationalunit' { 'OU' }
          default { $null }
        }
        if ($ObjectType) { Debug-Log "Detected type from ObjectClass: $ObjectType" -Type "Tracing" }
      }

      ## Fallback to property-based detection
      if (-not $ObjectType) {
        if ($obj.PSObject.Properties['Site'] -or $obj.PSObject.Properties['IsGlobalCatalog']) {
          $ObjectType = 'Domain Controller'
        }
        elseif ($obj.PSObject.Properties['ComputerType'] -or $obj.PSObject.Properties['OperatingSystem']) {
          $ObjectType = 'Computer'
        }
        elseif ($obj.PSObject.Properties['Members'] -or $obj.PSObject.Properties['GroupType']) {
          $ObjectType = 'Group'
        }
        elseif ($obj.PSObject.Properties['distinguishedName'] -and $obj.distinguishedName -match '^OU=') {
          $ObjectType = 'OU'
        }
        elseif ($obj.PSObject.Properties['SamAccountName'] -or $obj.PSObject.Properties['UserPrincipalName']) {
          $ObjectType = 'User'
        }
      }
    }
  }

  ## If still not detected, fail gracefully
  if (-not $ObjectType) {
    $objName = if ($obj -is [hashtable]) { $obj['Name'] } else { $obj.Name }
    Debug-Log "Unable to determine object type for: $objName" -Type "Problem"
    if ($obj -is [hashtable]) {
      Debug-Log "Hashtable keys: $($obj.Keys -join ', ')" -Type "Tracing"
    } else {
      Debug-Log "Object properties: $($obj.PSObject.Properties.Name -join ', ')" -Type "Tracing"
    }
    Show-Modal -title "Error" -msg "Unable to determine object type.`n`nObject: $objName"
    return $false
  }
  Debug-Log "Invoke-ObjectOperation - Operation: $Operation, ObjectType: $ObjectType, Objects: $($Objects.Count)" -Type "Insight"

  ## ----------{ Delete Operation }---------
  if ($Operation -eq 'Delete') {
    foreach ($obj in $Objects) {
      $name = if ($obj -is [hashtable]) { $obj['Name'] } else { $obj.Name }
      Debug-Log "Delete requested for $ObjectType '$name'" -Type "Warning"
      ## Double confirmation for safety
      $result = Show-Modal -title "DELETE CONFIRMATION" -msg "⚠️  WARNING: You are about to DELETE:`n`n Type: $ObjectType`n Name: $name`n`n This action CANNOT be undone!`n`n Are you absolutely sure?" -YesNo
      if ($result -ne 0) {
        Debug-Log "Delete cancelled by user" -Type "Insight"
        return $false
      }
      try {
        if ($Script:DemoMode) {
          ## Demo mode - remove from arrays
          switch ($ObjectType) {
            'User' {
              $Script:Users = $Script:Users | Where-Object {
                $objName = if ($_ -is [hashtable]) { $_['Name'] } else { $_.Name }
                $objName -ne $name
              }
              if ($Script:rawUsers) {
                $Script:rawUsers = $Script:rawUsers | Where-Object {
                  $objName = if ($_ -is [hashtable]) { $_['Name'] } else { $_.Name }
                  $objName -ne $name
                }
              }
            }
            'Group' {
              if ($Script:Groups) {
                $Script:Groups = $Script:Groups | Where-Object {
                  $objName = if ($_ -is [hashtable]) { $_['Name'] } else { $_.Name }
                  $objName -ne $name
                }
              }
              if ($Script:rawDemoGroups) {
                $Script:rawDemoGroups = $Script:rawDemoGroups | Where-Object {
                  $objName = if ($_ -is [hashtable]) { $_['Name'] } else { $_.Name }
                  $objName -ne $name
                }
              }
            }
            'Computer' {
              if ($Script:Computers) {
                $Script:Computers = $Script:Computers | Where-Object {
                  $objName = if ($_ -is [hashtable]) { $_['Name'] } else { $_.Name }
                  $objName -ne $name
                }
              }
            }
            'Domain Controller' {
              if ($Script:DCs) {
                $Script:DCs = $Script:DCs | Where-Object {
                  $objName = if ($_ -is [hashtable]) { $_['Name'] } else { $_.Name }
                  $objName -ne $name
                }
              }
            }
            'OU' {
              if ($Script:rawOUs) {
                $Script:rawOUs = $Script:rawOUs | Where-Object {
                  $objName = if ($_ -is [hashtable]) { $_['Name'] } else { $_.Name }
                  $objName -ne $name
                }
              }
            }
          }
          Debug-Log "Deleted $ObjectType '$name' (demo mode)" -Type "Success"
          Show-Modal -title "Success" -msg "$ObjectType '$name' deleted successfully (demo mode)"
        } else {
          ## Production Mode - use AD cmdlets
          $samAccountName = if ($obj -is [hashtable]) { $obj['SamAccountName'] } else { $obj.SamAccountName }
          $dn = if ($obj -is [hashtable]) { $obj['distinguishedName'] } else { $obj.DistinguishedName }
          switch ($ObjectType) {
            'User'     { Remove-ADUser -Identity $samAccountName -Confirm:$false -ErrorAction Stop }
            'Group'    { Remove-ADGroup -Identity $name -Confirm:$false -ErrorAction Stop }
            'Computer' { Remove-ADComputer -Identity $samAccountName -Confirm:$false -ErrorAction Stop }
            'OU' {
              ## Check if OU is protected from deletion
              $adOU = Get-ADOrganizationalUnit -Identity $dn -Properties ProtectedFromAccidentalDeletion -ErrorAction Stop
              if ($adOU.ProtectedFromAccidentalDeletion) { Set-ADOrganizationalUnit -Identity $dn -ProtectedFromAccidentalDeletion $false -ErrorAction Stop }
              Remove-ADOrganizationalUnit -Identity $dn -Confirm:$false -ErrorAction Stop
            }
            'Domain Controller' {
              Show-Modal -title "Not Supported" -msg "Domain Controllers cannot be deleted from this interface"
              return $false
            }
            default { Remove-ADObject -Identity $dn -Confirm:$false -ErrorAction Stop }
          }
          Debug-Log "Deleted $ObjectType '$name' from AD" -Type "Success"
          Show-Modal -title "Success" -msg "$ObjectType '$name' deleted successfully"
        }
        ## Refresh UI
        Refresh-Data -domain $Script:CurrentDomain -RebuildTree
        if ($Script:FilterStatusLabel) { Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel }
        return $true
      } catch {
        Debug-Log "Failed to delete $ObjectType '$name': $($_.Exception.Message)" -Type "Problem"
        Show-Modal -title "Delete Failed" -msg "Failed to delete $ObjectType '$name':`n`n$($_.Exception.Message)"
        return $false
      }
    }
  }

  ## ----------{ Move operation }---------
  if ($Operation -eq 'Move') {
    ## Validate objects are moveable (users, groups, or computers)
    if ($ObjectType -notin @('User', 'Group', 'Computer')) {
      Show-Modal -title "Not Supported" -msg "Object type '$ObjectType' cannot be moved"
      return $false
    }

    ## Get current OU (from first object)
    $firstObj = $Objects[0]
    $currentOU = if ($firstObj -is [hashtable]) {
      if ($firstObj.ContainsKey('distinguishedName')) {
        ($firstObj['distinguishedName'] -replace '^CN=[^,]+,', '')
      } else {
        $ouVal = $firstObj['OU']
        if ($ouVal -is [array]) { $ouVal -join ' > ' } else { $ouVal ?? "N/A" }
      }
    } else {
      if ($firstObj.DistinguishedName) {
        ($firstObj.DistinguishedName -replace '^CN=[^,]+,', '')
      } else {
        if ($firstObj.OU -is [array]) { $firstObj.OU -join ' > ' } else { $firstObj.OU ?? "N/A" }
      }
    }

    ## Get list of available OUs
    $ouList = if ($Script:DemoMode) {
      $Script:Users | Where-Object {
        $ou = if ($_ -is [hashtable]) { $_['OU'] } else { $_.OU }
        $ou
      } | ForEach-Object {
        $ou = if ($_ -is [hashtable]) { $_['OU'] } else { $_.OU }
        if ($ou -is [array]) { $ou -join ' > ' } else { $ou }
      } | Select-Object -Unique | Sort-Object
    } else {
      try {
        Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty DistinguishedName | Sort-Object
      } catch {
        @("CN=Users,DC=example,DC=com")
      }
    }

    ## Create dialog
    if ($IsBulk) {
      $title = "Bulk Move - $($Objects.Count) Objects"
      $dlg = [Terminal.Gui.Dialog]::new($title, 70, 18)
      $lblInfo = [Terminal.Gui.Label]::new("Moving $($Objects.Count) object(s) to:")
      $lblInfo.X = 2
      $lblInfo.Y = 1
      $dlg.Add($lblInfo)
      $yStart = 3
    } else {
      $objName = if ($firstObj -is [hashtable]) { $firstObj['Name'] } else { $firstObj.Name }
      $title = "Move Object - $objName"
      $dlg = [Terminal.Gui.Dialog]::new($title, 70, 18)

      ## Show current location
      $lblCurrent = [Terminal.Gui.Label]::new("Current location:")
      $lblCurrent.X = 2
      $lblCurrent.Y = 1
      $dlg.Add($lblCurrent)
      $lblCurrentOU = [Terminal.Gui.Label]::new($currentOU)
      $lblCurrentOU.X = 20
      $lblCurrentOU.Y = 1
      $dlg.Add($lblCurrentOU)
      $lblTarget = [Terminal.Gui.Label]::new("Move to OU:")
      $lblTarget.X = 2
      $lblTarget.Y = 3
      $dlg.Add($lblTarget)
      $yStart = 4
    }

    ## OU ListView
    $lstOU = [Terminal.Gui.ListView]::new($ouList)
    $lstOU.X = 2
    $lstOU.Y = $yStart
    $lstOU.Width = [Terminal.Gui.Dim]::Fill(2)
    $lstOU.Height = if ($IsBulk) { 10 } else { 8 }
    $dlg.Add($lstOU)

    ## Capture functions for closure
    $showModalFunc = ${function:Show-Modal}
    $debugLogFunc = ${function:Debug-Log}
    $refreshDataFunc = ${function:Refresh-Data}
    $manageFilterFunc = ${function:Manage-FilterStatusLabel}

    ## Move button
    $btnMoveText = if ($IsBulk) { "Move All" } else { "Move" }
    $btnMove = [Terminal.Gui.Button]::new($btnMoveText)
    $btnMove.add_Clicked({
      if ($lstOU.SelectedItem -lt 0) {
        & $showModalFunc -title "Error" -msg "Please select a target OU"
        return
      }
      $targetOU = $ouList[$lstOU.SelectedItem]
      if ($targetOU -eq $currentOU) {
        $errMsg = if ($IsBulk) { "Objects are already in that OU" } else { "Object is already in that OU" }
        & $showModalFunc -title "Error" -msg $errMsg
        return
      }
      ## Confirm move
      $firstObjName = if ($Objects[0] -is [hashtable]) { $Objects[0]['Name'] } else { $Objects[0].Name }
      $confirmMsg = if ($IsBulk) {
        "Move $($Objects.Count) object(s) to:`n$targetOU?"
      } else {
        "Move '$firstObjName' to:`n$targetOU?"
      }
      $confirmTitle = if ($IsBulk) { "Confirm Bulk Move" } else { "Confirm Move" }
      $confirm = & $showModalFunc -title $confirmTitle -msg $confirmMsg -YesNo
      if ($confirm -ne 0) { return }
      ## Perform the move
      $successCount = 0
      $failCount    = 0
      $errors       = @()
      foreach ($obj in $Objects) {
        $name = if ($obj -is [hashtable]) { $obj['Name'] } else { $obj.Name }
        try {
          if ($Script:DemoMode) {
            ## Demo mode - update OU property
            if ($obj -is [hashtable]) {
              if ($obj.ContainsKey('OU')) { $obj['OU'] = $targetOU }
            } else {
              if ($obj.PSObject.Properties['OU']) { $obj.OU = $targetOU }
            }
            $successCount++
            & $debugLogFunc "Moved $name to $targetOU (demo mode)" -Type "Insight"
          } else {
            ## Production Mode
            $dn = if ($obj -is [hashtable]) { $obj['distinguishedName'] } else { $obj.DistinguishedName }
            Move-ADObject -Identity $dn -TargetPath $targetOU -ErrorAction Stop
            $successCount++
            & $debugLogFunc "Moved $name to $targetOU in AD" -Type "Insight"
          }
        } catch {
          $failCount++
          $errors += "${name}: $($_.Exception.Message)"
          & $debugLogFunc "Failed to move ${name}: $($_.Exception.Message)" -Type "Problem"
        }
      }
      ## Show results
      if ($IsBulk -or $failCount -gt 0) {
        $msg = "Successfully moved $successCount object(s)"
        if ($failCount -gt 0) {
          $msg += "`n`nFailed: $failCount"
          if ($errors.Count -gt 0 -and $errors.Count -le 5) { $msg += "`n`nErrors:`n" + ($errors -join "`n") }
        }
        $resultTitle = if ($failCount -eq 0) { "Success" } else { "Move Complete" }
        & $showModalFunc -title $resultTitle -msg $msg
      } else {
        $successMsg = if ($Script:DemoMode) { "Object moved successfully (demo mode)" } else { "Object moved successfully" }
        & $showModalFunc -title "Success" -msg $successMsg
      }
      ## Refresh UI
      & $refreshDataFunc -domain $Script:CurrentDomain -RebuildTree
      if ($Script:FilterStatusLabel) { & $manageFilterFunc -Action 'Update' -Label $Script:FilterStatusLabel }
      [Terminal.Gui.Application]::RequestStop()
    }.GetNewClosure())
    $dlg.AddButton($btnMove)

    ## Cancel button
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }.GetNewClosure())
    $dlg.AddButton($btnCancel)
    [Terminal.Gui.Application]::Run($dlg)
  }
}

## ----------{ Single unified move/delete function }---------
function Invoke-BulkMove {
  if ($Script:SelectedObjects.Count -eq 0) {
    Show-Modal "No Selection" "No objects selected"
    return
  }
  $objects = @()
  foreach ($objName in $Script:SelectedObjects) {
    $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
    $obj = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
    if (-not $obj) { $obj = $Script:Groups | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1 }
    if (-not $obj) { $obj = $Script:Computers | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1 }
    if ($obj) { $objects += $obj }
  }
  if ($objects.Count -eq 0) {
    Show-Modal "Error" "No valid objects found in selection"
    return
  }

  ## Determine object types in selection
  $hasUsers     = $false
  $hasGroups    = $false
  $hasComputers = $false
  $hasOther     = $false

  foreach ($obj in $objects) {
    if ($obj.PSObject.Properties['SamAccountName'] -and -not $obj.PSObject.Properties['Members'] -and -not $obj.PSObject.Properties['ComputerType']) {
      $hasUsers = $true
    } elseif ($obj.PSObject.Properties['Members']) {
      $hasGroups = $true
    } elseif ($obj.PSObject.Properties['ComputerType']) {
      $hasComputers = $true
    } else {
      $hasOther = $true
    }
  }

  ## Check for mixed types
  $typeCount = @($hasUsers, $hasGroups, $hasComputers, $hasOther) | Where-Object { $_ } | Measure-Object | Select-Object -ExpandProperty Count

  if ($typeCount -gt 1) {
    Show-Modal "Mixed Types" "Bulk move currently only supports objects of the same type.`n`nYour selection contains mixed object types (Users, Groups, Computers).`n`nPlease select objects of only one type."
    return
  }

  ## Determine the single object type
  $objectType = if ($hasUsers) { 'User' }
                elseif ($hasGroups) { 'Group' }
                elseif ($hasComputers) { 'Computer' }
                else { $null }

  if (-not $objectType) {
    Show-Modal "Error" "Unable to determine object type for bulk move"
    return
  }

  ## Perform bulk move
  Invoke-ObjectOperation -Objects $objects -Operation 'Move' -ObjectType $objectType -IsBulk

  ## Clear selection
  $Script:SelectedObjects = @()
  $Script:SelectionMode   = $false
  Show-SelectionPanel -Parent $win
}

## ----------{ Bulk account status management }---------
function Manage-AccountStatus {
  <#
  .SYNOPSIS
  Enable or disable multiple AD accounts in bulk

  .PARAMETER Objects
  Array of user or computer objects to modify

  .PARAMETER Action
  'Enable' or 'Disable'

  .PARAMETER Reason
  Optional reason for status change (logged only)

  .EXAMPLE
  Manage-AccountStatus -Objects $selectedUsers -Action 'Disable'

  .EXAMPLE
  Manage-AccountStatus -Objects $users -Action 'Enable' -Reason "Rehire"
  #>

  param(
    [Parameter(Mandatory=$true)]
    [object[]]$Objects,
    [Parameter(Mandatory=$true)]
    [ValidateSet('Enable', 'Disable')]
    [string]$Action,
    [string]$Reason
  )

  if (-not $Objects -or $Objects.Count -eq 0) {
    Show-Modal "No Objects" "No objects provided"
    return
  }

  ## Validate all objects are users or computers
  $validObjects = @()
  foreach ($obj in $Objects) {
    $isUser = $obj.PSObject.Properties.Match('SamAccountName').Count -gt 0 -and -not $obj.PSObject.Properties.Match('ComputerType')
    $isComputer = $obj.PSObject.Properties.Match('ComputerType').Count -gt 0
    if ($isUser -or $isComputer) { $validObjects += $obj }
  }

  if ($validObjects.Count -eq 0) {
    Show-Modal "Invalid Objects" "No valid user or computer objects found"
    return
  }

  ## Confirmation
  $actionVerb = if ($Action -eq 'Enable') { "ENABLE" } else { "DISABLE" }
  $reasonText = if ($Reason) { "`n`nReason: $Reason" } else { "" }
  $confirmMsg = "Are you sure you want to $actionVerb $($validObjects.Count) account(s)?$reasonText"
  $confirm = Show-Modal "Confirm $Action" $confirmMsg -YesNo
  if ($confirm -ne 0) {
    Debug-Log "Bulk $Action cancelled by user" -Type "Insight"
    return
  }

  Debug-Log "Starting bulk $Action on $($validObjects.Count) objects" -Type "Insight"
  if ($Reason) { Debug-Log "Reason: $Reason" -Type "Insight" }

  ## Process objects
  $successCount = 0
  $failCount    = 0
  $errors       = @()

  foreach ($obj in $validObjects) {
    $name = $obj.Name
    $objectType = if ($obj.PSObject.Properties.Match('ComputerType')) { 'Computer' } else { 'User' }
    try {
      if ($Script:DemoMode) {
        ## Demo mode - update object properties
        if ($Action -eq 'Enable') {
          $obj.Enabled = $true
          $obj.Disabled = $false
        } else {
          $obj.Enabled = $false
          $obj.Disabled = $true
        }
          $successCount++
          Debug-Log "${Action}d $objectType '$name' (demo mode)" -Type "Insight"
      } else {
        ## Production Mode - use AD cmdlets
        if ($Action -eq 'Enable') {
          if ($objectType -eq 'User') { Enable-ADAccount -Identity $obj.SamAccountName -ErrorAction Stop
          } else { Enable-ADAccount -Identity $obj.SamAccountName -ErrorAction Stop }
        } else {
          if ($objectType -eq 'User') { Disable-ADAccount -Identity $obj.SamAccountName -ErrorAction Stop
          } else { Disable-ADAccount -Identity $obj.SamAccountName -ErrorAction Stop }
        }
        $successCount++
        Debug-Log "${Action}d $objectType '$name' in AD" -Type "Success"
      }
    } catch {
    $failCount++
    $errors += "${name}: $($_.Exception.Message)"
    Debug-Log "Failed to $Action $objectType '$name': $($_.Exception.Message)" -Type "Problem"
  }

  ## Show results
  $msg = "Successfully ${Action}d $successCount account(s)"
  if ($failCount -gt 0) {
    $msg += "`n`nFailed: $failCount"
    if ($errors.Count -gt 0 -and $errors.Count -le 5) { $msg += "`n`nErrors:`n" + ($errors -join "`n") }
  }

  Show-Modal $(if ($failCount -eq 0) { "Success" } else { "$Action Complete" }) $msg
  ## Refresh UI
  Refresh-Data -domain $Script:CurrentDomain
  Build-Tree -domain $Script:CurrentDomain
  if ($Script:FilterStatusLabel) { Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel }
  Debug-Log "Bulk $Action completed - $successCount succeeded, $failCount failed" -Type "Success"
}
}

## ----------{ Common tab builders / helper functions for UI building- re-usable across dialogs }----------
function Add-LabelAndField {
  <#
  .SYNOPSIS
  Adds a label and text field pair, incrementing Y position
  #>

  param(
    [Parameter(Mandatory)]$View,
    [Parameter(Mandatory)][ref]$Y,
    [Parameter(Mandatory)][string]$Label,
    [Parameter(Mandatory)][string]$FieldName,
    [Parameter(Mandatory)]$State,
    [string]$Value = "",
    [int]$LabelX = 2,
    [int]$FieldX = 20,
    [int]$Width = 60,
    [bool]$ReadOnly = $false,
    [bool]$IsTextField = $true
  )

  $lbl   = [Terminal.Gui.Label]::new($Label)
  $lbl.X = $LabelX
  $lbl.Y = $Y.Value
  $View.Add($lbl)

  if ($IsTextField) {
    $State.$FieldName = [Terminal.Gui.TextField]::new($Value)
    $State.$FieldName.X = $FieldX
    $State.$FieldName.Y = $Y.Value
    $State.$FieldName.Width = $Width
    $State.$FieldName.ReadOnly = $ReadOnly
  } else {
    $State.$FieldName = [Terminal.Gui.Label]::new($Value)
    $State.$FieldName.X = $FieldX
    $State.$FieldName.Y = $Y.Value
  }
  $View.Add($State.$FieldName)
  $Y.Value += 1
}

function Add-SectionHeader {
  <#
  .SYNOPSIS
  Adds a section header and increments Y
  #>

  param(
    [Parameter(Mandatory)]$View,
    [Parameter(Mandatory)][ref]$Y,
    [Parameter(Mandatory)][string]$Text,
    [int]$X = 2,
    [int]$SpaceBefore = 1,
    [int]$SpaceAfter = 2
  )

  $Y.Value += $SpaceBefore
  $lbl      = [Terminal.Gui.Label]::new($Text)
  $lbl.X    = $X
  $lbl.Y    = $Y.Value
  $View.Add($lbl)
  $Y.Value += $SpaceAfter
}

## New Properties Dialog - New AND Improved!
function New-PropertiesDialog {
  param(
    [Parameter(Mandatory=$true)]
    [string]$Title,
    [int]$Width = 100,
    [int]$Height = 40,
    [Parameter(Mandatory=$true)]
    [array]$Tabs,
    [Parameter(Mandatory=$true)]
    [object]$Data,
    [scriptblock]$OnApply,
    [scriptblock]$OnOK,
    [bool]$IncludeSearchTab = $true,
    [hashtable]$SearchTabConfig = @{}
  )

  try {
    Debug-Log "New-PropertiesDialog called for: $Title" -Type "Tracing"

    ## Capture functions for closures
    $debugLogFunc  = ${function:Debug-Log}
    $showModalFunc = ${function:Show-Modal}
    ## Shared state
    $sharedState = @{}
    ## Auto-add search tab if requested
    if ($IncludeSearchTab) {
      Debug-Log "Adding search tab to tabs array" -Type "Tracing"
      $searchTab = New-SearchTab -Data $Data -Config $SearchTabConfig
      $Tabs += $searchTab
    }

    ## Create buttons
    $btnOK     = [Terminal.Gui.Button]::new("OK")
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnApply  = [Terminal.Gui.Button]::new("Apply")

    ## Button handlers with captured functions
    $btnOK.add_Clicked({
      & $debugLogFunc "OK clicked" -Type "Insight"
      if ($OnOK) { & $OnOK $Data $sharedState }
      [Terminal.Gui.Application]::RequestStop()
    }.GetNewClosure())

    $btnCancel.add_Clicked({
      & $debugLogFunc "Cancel clicked" -Type "Insight"
      [Terminal.Gui.Application]::RequestStop()
    }.GetNewClosure())

    $btnApply.add_Clicked({
      & $debugLogFunc "Apply clicked" -Type "Insight"
      if ($OnApply) {
        try {
          & $OnApply $Data $sharedState
        } catch {
          & $debugLogFunc "Apply failed: $($_.Exception.Message)" -Type "Problem"
          [Terminal.Gui.Application]::MainLoop.Invoke({
            & $showModalFunc "Error" "Failed to apply changes:`n$($_.Exception.Message)"
          })
        }
      }
    }.GetNewClosure())

    ## Create dialog with buttons
    Debug-Log "Creating dialog with buttons" -Type "Tracing"
    $dialog = [Terminal.Gui.Dialog]::new($Title, $Width, $Height, $btnOK, $btnCancel, $btnApply)

    ## Create TabView
    Debug-Log "Creating TabView" -Type "Tracing"
    $tabView = [Terminal.Gui.TabView]::new()
    $tabView.X = 0
    $tabView.Y = 0
    $tabView.Width  = [Terminal.Gui.Dim]::Fill()
    $tabView.Height = [Terminal.Gui.Dim]::Fill(1)

    ## Build each tab
    foreach ($tabDef in $Tabs) {
      Debug-Log "Creating tab: $($tabDef.Name)" -Type "Insight"

      $tab = [Terminal.Gui.TabView+Tab]::new()
      $tab.Text = [NStack.ustring]::Make($tabDef.Name)

      $view = [Terminal.Gui.View]::new()
      $view.X = 0; $view.Y = 0
      $view.Width = [Terminal.Gui.Dim]::Fill()
      $view.Height = [Terminal.Gui.Dim]::Fill()

      ## Call builder
      if ($tabDef.Builder) {
        try {
          & $tabDef.Builder $view $Data $sharedState
          Debug-Log "Builder completed for tab: $($tabDef.Name)" -Type "Tracing"
        } catch {
          Debug-Log "Error building tab $($tabDef.Name): $($_.Exception.Message)" -Type "Problem"
          Debug-Log "Error line: $($_.InvocationInfo.ScriptLineNumber)" -Type "Problem"
          Debug-Log "Stack: $($_.ScriptStackTrace)" -Type "Problem"
          throw
        }
      }
      $tab.View = $view
      $tabView.AddTab($tab, $false)
    }

    Debug-Log "Adding TabView to dialog" -Type "Tracing"
    $dialog.Add($tabView)
    Debug-Log "All tabs added, running dialog" -Type "Success"
    [Terminal.Gui.Application]::Run($dialog)
    Debug-Log "Dialog closed normally" -Type "Insight"
  } catch {
    Debug-Log "Exception in New-PropertiesDialog: $($_.Exception.Message)" -Type "Problem"
    Debug-Log "Stack: $($_.ScriptStackTrace)" -Type "Problem"
    Show-Modal "Error" "Failed to display dialog:`n$($_.Exception.Message)"
  }
}

## ----------{ Common tab builders - re-usable across dialogs }----------
function New-SearchTab {
  <#
  .SYNOPSIS
  Creates a standard Search/Lookup tab for any object type
  #>
  param(
    [object]$Data,
    [hashtable]$Config
  )

  ## Determine object type
  $objectType = if ($Config.ObjectType) {
    $Config.ObjectType
  } elseif ($Data.ObjectClass -eq 'user')     { 'User'
  } elseif ($Data.ObjectClass -eq 'group')    { 'Group'
  } elseif ($Data.ObjectClass -eq 'computer') { 'Computer'
  } else {'Object'
  }

  $searchTypes = $Config.SearchTypes ?? @("$objectType", "Group", "User", "Computer", "OU")

  return @{
    Name = "Search/Lookup"
    Builder = {
      param($view, $data, $state)
      $y = 1
      ## Domain field
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Domain:" -FieldName 'txtSearchDomain' -State $state -Value $Script:CurrentDomain -Width 30
      $y += 1
      ## Name field
      $nameLabel = "Name:"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label $nameLabel -FieldName 'txtSearchName' -State $state -Value $data.Name -Width 30
      $y += 1
      ## Search type dropdown
      $lblType = [Terminal.Gui.Label]::new("Type:")
      $lblType.X = 2; $lblType.Y = $y
      $view.Add($lblType)

      $state.cmbSearchType = [Terminal.Gui.ComboBox]::new()
      $state.cmbSearchType.X = 15; $state.cmbSearchType.Y = $y; $state.cmbSearchType.Width = 20
      $state.cmbSearchType.SetSource($searchTypes)
      $state.cmbSearchType.SelectedItem = 0
      $view.Add($state.cmbSearchType)
      $y += 2

      ## Filter field (top right)
      $lblFilter = [Terminal.Gui.Label]::new("Filter Results:")
      $lblFilter.X = 50; $lblFilter.Y = 1
      $view.Add($lblFilter)

      $state.txtSearchFilter = [Terminal.Gui.TextField]::new("")
      $state.txtSearchFilter.X = 67; $state.txtSearchFilter.Y = 1; $state.txtSearchFilter.Width = 20
      $view.Add($state.txtSearchFilter)

      ## Filter handler
      $state.txtSearchFilter.add_TextChanged({
        if ($state.currentSearchOutputLines) { $search = $state.txtSearchFilter.Text.ToString().Trim()
          if ($search) { $state.txtSearchOutput.Text = ($state.currentSearchOutputLines | Where-Object { $_ -match "(?i)$search" }) -join "`n"
          } else { $state.txtSearchOutput.Text = $state.currentSearchOutputLines -join "`n"
          }
        }
      }.GetNewClosure())

      ## Results label
      $lblResult = [Terminal.Gui.Label]::new("Properties:")
      $lblResult.X = 2; $lblResult.Y = $y
      $view.Add($lblResult)
      $y += 1

      ## Results text view
      $state.txtSearchOutput          = [Terminal.Gui.TextView]::new()
      $state.txtSearchOutput.X        = 2; $state.txtSearchOutput.Y = $y
      $state.txtSearchOutput.Width    = [Terminal.Gui.Dim]::Fill(2)
      $state.txtSearchOutput.Height   = [Terminal.Gui.Dim]::Fill(1)
      $state.txtSearchOutput.ReadOnly = $true
      $state.txtSearchOutput.WordWrap = $false
      $view.Add($state.txtSearchOutput)

      ## Object-specific checkboxes
      script:Set-ObjectCheckboxes -View $view -State $state -Data $data -ObjectType $objectType -Mode 'Create'

      ## Search button
      $btnSearch   = [Terminal.Gui.Button]::new("Search")
      $btnSearch.X = 50; $btnSearch.Y = 3
      $view.Add($btnSearch)

      ## Auto-populate
      [Terminal.Gui.Application]::MainLoop.Invoke({
        if ($data) {
          $lines = @()
          $data.PSObject.Properties | ForEach-Object {
            $value = if ($_.Value -is [array]) {
              $_.Value -join ', '
            } elseif ($null -eq $_.Value) {
              ''
            } else {
              $_.Value.ToString()
            }
            $lines += "$($_.Name.PadRight(25)): $value"
          }
          $state.txtSearchOutput.Text = $lines -join "`n"
          $state.currentSearchOutputLines = $lines
          ## Update checkboxes based on object type
          if ($Script:SetObjectCheckboxes_Func) {
            & $Script:SetObjectCheckboxes_Func -State $state -Data $data -ObjectType $objectType -Mode 'Update'
          }
        }
      }.GetNewClosure())
    }
  }
}

## ----------{ Helper funcitons for UI building }----------
function Add-LabelAndField {
  <#
  .SYNOPSIS
  Adds a label and text field pair, incrementing Y position
  #>

  param(
    [Parameter(Mandatory)]$View,
    [Parameter(Mandatory)][ref]$Y,
    [Parameter(Mandatory)][string]$Label,
    [Parameter(Mandatory)][string]$FieldName,
    [Parameter(Mandatory)]$State,
    [string]$Value = "",
    [int]$LabelX = 2,
    [int]$FieldX = 15,
    [int]$Width = 60,
    [bool]$ReadOnly = $false,
    [bool]$IsTextField = $true
  )

  $lbl = [Terminal.Gui.Label]::new($Label)
  $lbl.X = $LabelX
  $lbl.Y = $Y.Value
  $View.Add($lbl)

  if ($IsTextField) {
    $State.$FieldName = [Terminal.Gui.TextField]::new($Value)
    $State.$FieldName.X = $FieldX
    $State.$FieldName.Y = $Y.Value
    $State.$FieldName.Width = $Width
    $State.$FieldName.ReadOnly = $ReadOnly
  } else {
    $State.$FieldName = [Terminal.Gui.Label]::new($Value)
    $State.$FieldName.X = $FieldX
    $State.$FieldName.Y = $Y.Value
  }
  $View.Add($State.$FieldName)
  $Y.Value += 1
}

function Add-SectionHeader {
  <#
  .SYNOPSIS
  Adds a section header and increments Y
  #>

  param(
    [Parameter(Mandatory)]$View,
    [Parameter(Mandatory)][ref]$Y,
    [Parameter(Mandatory)][string]$Text,
    [int]$X = 2,
    [int]$SpaceBefore = 1,
    [int]$SpaceAfter = 2
  )

  $Y.Value += $SpaceBefore
  $lbl   = [Terminal.Gui.Label]::new($Text)
  $lbl.X = $X
  $lbl.Y = $Y.Value
  $View.Add($lbl)
  $Y.Value += $SpaceAfter
}

function Show-GroupPropertiesDialog {
  param($group)

  if (-not $group) {
    Debug-Log "Group object is null" -Type "Warning"
    return
  }
  Debug-Log "Show-GroupPropertiesDialog starting for: $($group.Name)" -Type "Insight"

  ## ----------{ General Tab }---------
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $data, $state)

      ## Use $data instead of $group
      $grp = $data

      ## Capture functions for this builder's closures
      $showAuditLogFunc = ${function:Show-AuditLogDialog}
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Group Information"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Group name:" -FieldName 'txtName' -State $state -Value ($grp.Name ?? "") -IsTextField $false
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Description:" -FieldName 'txtDescription' -State $state -Value ($grp.Description ?? "")

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Group Details"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Type:" -FieldName 'lblType' -State $state -Value ($grp.Type ?? "Security") -IsTextField $false
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Scope:" -FieldName 'lblScope' -State $state -Value ($grp.Scope ?? "Global") -IsTextField $false

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Contact Information"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Email:" -FieldName 'txtEmail' -State $state -Value ($grp.Email ?? "")
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Managed by:" -FieldName 'txtManagedBy' -State $state -Value ($grp.ManagedBy ?? "")
      $y += 1

      ## Audit Log button
      $btnAuditLog = [Terminal.Gui.Button]::new("View Audit Log...")
      $btnAuditLog.X = 2; $btnAuditLog.Y = $y
      $btnAuditLog.add_Clicked({
        & $showAuditLogFunc -Object $grp -ObjectType 'Group'
      }.GetNewClosure())
      $view.Add($btnAuditLog)
    }
  }

  ## ----------{ Members Tab }---------
  $membersTab = @{
    Name = "Members"
    Builder = {
      param($view, $data, $state)

      ## Use $data instead of $group
      $grp = $data

      ## CAPTURE FUNCTIONS FOR THIS BUILDER'S CLOSURES
      $debugLogFunc = ${function:Debug-Log}
      $demoMode     = $Script:DemoMode
      $allUsers     = $Script:Users
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Group Members:")
      $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      $state.lstMembers = [Terminal.Gui.ListView]::new()
      $state.lstMembers.X = 2; $state.lstMembers.Y = $y
      $state.lstMembers.Width = [Terminal.Gui.Dim]::Fill(2)
      $state.lstMembers.Height = 25

      $state.memberList = @()
      if ($demoMode) {
        $state.memberList = $allUsers | Where-Object { $_.Groups -contains $grp.Name } | Select-Object -ExpandProperty Name | Sort-Object
      } else {
        try {
          $members = Get-ADGroupMember -Identity $grp.Name -ErrorAction Stop
          $state.memberList = $members | Select-Object -ExpandProperty Name | Sort-Object
        } catch {
          & $debugLogFunc "Failed to get group members: $($_.Exception.Message)" -Type "Warning"
          $state.memberList = @()
        }
      }

      if ($state.memberList.Count -gt 0) {
        $state.lstMembers.SetSource($state.memberList)
      } else {
        $state.lstMembers.SetSource(@("(No members)"))
      }
      $view.Add($state.lstMembers)
    }
  }

  ## ----------{ Report Tab }---------
  $reportTab = @{
    Name = "Report"
    Builder = {
      param($view, $data, $state)

      ## Use $data instead of $group
      $grp = $data

      ## Capture all functions and data at builder level
      $debugLogFunc  = ${function:Debug-Log}
      $showModalFunc = ${function:Show-Modal}
      $allGroups     = $Script:Groups   ## capture the groups array
      $allUsers      = $Script:Users    ## capture the users array
      $demoMode      = $Script:DemoMode ## capture the demo mode flag

      ## Get detailed member information
      $memberDetails = @()
      if ($demoMode) {
        $members = $allUsers | Where-Object { $_.Groups -contains $grp.Name }
        foreach ($member in $members) {
          $memberDetails += [PSCustomObject]@{
            Name = $member.Name
            SamAccountName = $member.SamAccountName
            Email = $member.EmailAddress
            Department = $member.Department
            Title = $member.Title
            Enabled = $member.Enabled
            Status = if ($member.Enabled) { "Enabled" } else { "Disabled" }
          }
        }
      } else {
        try {
          $members = Get-ADGroupMember -Identity $grp.Name -ErrorAction Stop
          foreach ($member in $members) {
            if ($member.objectClass -eq 'user') {
              $userDetails = Get-ADUser -Identity $member.SamAccountName -Properties EmailAddress,Department,Title,Enabled -ErrorAction SilentlyContinue
              if ($userDetails) {
                $memberDetails += [PSCustomObject]@{
                  Name           = $userDetails.Name
                  SamAccountName = $userDetails.SamAccountName
                  Email          = $userDetails.EmailAddress
                  Department     = $userDetails.Department
                  Title          = $userDetails.Title
                  Enabled        = $userDetails.Enabled
                  Status         = if ($userDetails.Enabled) { "Enabled" } else { "Disabled" }
                }
              }
            }
          }
        } catch {
          & $debugLogFunc "Failed to get member details: $($_.Exception.Message)" -Type "Warning"
        }
      }

      $y = 1
      $lblHeader = [Terminal.Gui.Label]::new("Membership Report: $($grp.Name)")
      $lblHeader.X = 2; $lblHeader.Y = $y; $view.Add($lblHeader); $y += 2

      $totalMembers = $memberDetails.Count
      $enabledMembers = ($memberDetails | Where-Object { $_.Enabled -eq $true }).Count
      $disabledMembers = ($memberDetails | Where-Object { $_.Enabled -eq $false }).Count

      $lblStats = [Terminal.Gui.Label]::new("Total Members: $totalMembers (Enabled: $enabledMembers, Disabled: $disabledMembers)")
      $lblStats.X = 2; $lblStats.Y = $y; $view.Add($lblStats); $y += 2

      $lstReport = [Terminal.Gui.ListView]::new()
      $lstReport.X = 2; $lstReport.Y = $y
      $lstReport.Width = [Terminal.Gui.Dim]::Fill(2)
      $lstReport.Height = 22

      $displayItems = @()
      foreach ($member in ($memberDetails | Sort-Object Name)) {
        $statusIcon = if ($member.Enabled) { "✓" } else { "⊗" }
        $displayItems += "[$statusIcon] $($member.Name) | $($member.Department ?? 'N/A') | $($member.Title ?? 'N/A') | $($member.Email ?? 'N/A')"
      }

      if ($displayItems.Count -eq 0) { $displayItems = @("(No members)") }
      $lstReport.SetSource($displayItems)
      $view.Add($lstReport)
      $y = 26

      $btnExportFull = [Terminal.Gui.Button]::new("Export Full Report")
      $btnExportFull.X = 2; $btnExportFull.Y = $y
      $btnExportFull.add_Clicked({
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filename = "group_members_$($grp.Name)_$timestamp.csv"
        try {
          $memberDetails | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
          & $showModalFunc "Export Complete" "Membership report exported to:`n`n$filename"
          & $debugLogFunc "Exported group membership report to $filename" -Type "Success"
        } catch {
          & $showModalFunc "Export Failed" "Failed to export report:`n`n$($_.Exception.Message)"
        }
      }.GetNewClosure())
      $view.Add($btnExportFull)

      $btnCompare = [Terminal.Gui.Button]::new("Compare with Group...")
      $btnCompare.X = 25; $btnCompare.Y = $y
      $btnCompare.add_Clicked({
        ## Re-capture functions for this nested dialog
        $debugLogFunc2  = $debugLogFunc
        $showModalFunc2 = $showModalFunc
        $compareDlg = [Terminal.Gui.Dialog]::new("Compare Groups", 60, 24)

        $lblInfo   = [Terminal.Gui.Label]::new("Select group to compare with $($grp.Name):")
        $lblInfo.X = 2; $lblInfo.Y = 1
        $compareDlg.Add($lblInfo)

        ## Search box
        $lblSearch   = [Terminal.Gui.Label]::new("Search:")
        $lblSearch.X = 2; $lblSearch.Y = 3
        $compareDlg.Add($lblSearch)

        $txtSearch   = [Terminal.Gui.TextField]::new("")
        $txtSearch.X = 11; $txtSearch.Y = 3; $txtSearch.Width = 45
        $compareDlg.Add($txtSearch)

        ## Get all other groups
        $allOtherGroups = @()
        if ($demoMode) {
          & $debugLogFunc2 "Demo mode - getting groups from allGroups (count: $($allGroups.Count))" -Type "Tracing"
          & $debugLogFunc2 "Current group name: $($grp.Name)" -Type "Tracing"
          $allOtherGroups = @($allGroups | Where-Object { $_.Name -ne $grp.Name } | Sort-Object Name | Select-Object -ExpandProperty Name)
          & $debugLogFunc2 "Filtered other groups count: $($allOtherGroups.Count)" -Type "Tracing"
        } else {
          try {
            $allOtherGroups = @(Get-ADGroup -Filter * | Where-Object { $_.Name -ne $grp.Name } | Sort-Object Name | Select-Object -ExpandProperty Name)
          } catch {
            $allOtherGroups = @()
          }
        }
        & $debugLogFunc2 "Found $($allOtherGroups.Count) other groups for comparison" -Type "Tracing"

        ## Filtered groups list (starts with all groups)
        $filteredGroups = $allOtherGroups
        $lstGroups      = [Terminal.Gui.ListView]::new()
        $lstGroups.X    = 2; $lstGroups.Y = 5; $lstGroups.Width = 54; $lstGroups.Height = 14

        if ($filteredGroups.Count -gt 0) {
          $lstGroups.SetSource($filteredGroups)
        } else {
          $lstGroups.SetSource(@("(No other groups found)"))
        }
        $compareDlg.Add($lstGroups)

        ## Search-as-you-type handler
        $txtSearch.add_TextChanged({
          $searchText = $txtSearch.Text.ToString().Trim()
          if ([string]::IsNullOrWhiteSpace($searchText)) {
            $filteredGroups = $allOtherGroups ## No filter, show all groups
          } else {
            $filteredGroups = @($allOtherGroups | Where-Object { $_ -like "*$searchText*" }) ## Filter groups by search text
          }

          if ($filteredGroups.Count -gt 0) {
            $lstGroups.SetSource($filteredGroups)
          } else {
            $lstGroups.SetSource(@("(No matches found)"))
          }

          $lstGroups.SelectedItem = 0
          $lstGroups.SetNeedsDisplay()
        }.GetNewClosure())

        $btnSelect = [Terminal.Gui.Button]::new("Compare")
        $btnSelect.add_Clicked({
          & $debugLogFunc2 "Compare button clicked, filteredGroups count: $($filteredGroups.Count), selected: $($lstGroups.SelectedItem)" -Type "Tracing"

          if ($filteredGroups.Count -eq 0 -or $filteredGroups[0] -eq "(No matches found)" -or $filteredGroups[0] -eq "(No other groups found)") {
            & $showModalFunc2 "No Groups" "No groups available for comparison"
            return
          }
          if ($lstGroups.SelectedItem -lt 0) {
            & $showModalFunc2 "No Selection" "Please select a group to compare with"
            return
          }
          ## Get the selected group name
          $compareGroupName = $filteredGroups[$lstGroups.SelectedItem]
          & $debugLogFunc2 "Comparing $($grp.Name) with $compareGroupName" -Type "Tracing"
          [Terminal.Gui.Application]::RequestStop()

          ## Get members of current group
          $group1Members = @($memberDetails | Select-Object -ExpandProperty SamAccountName)

          ## Get members of comparison group
          $group2Members = @()
          if ($demoMode) {
            $group2Members = @($allUsers | Where-Object { $_.Groups -contains $compareGroupName } | Select-Object -ExpandProperty SamAccountName)
          } else {
            try {
              $group2Members = @(Get-ADGroupMember -Identity $compareGroupName -ErrorAction Stop | Select-Object -ExpandProperty SamAccountName)
            } catch {
              & $debugLogFunc2 "Failed to get members of $compareGroupName : $($_.Exception.Message)" -Type "Problem"
              $group2Members = @()
            }
          }

          $inBoth       = @($group1Members | Where-Object { $group2Members -contains $_ })
          $onlyInGroup1 = @($group1Members | Where-Object { $group2Members -notcontains $_ })
          $onlyInGroup2 = @($group2Members | Where-Object { $group1Members -notcontains $_ })

          $resultMsg = "Comparison: $($grp.Name) vs $compareGroupName`n`n"
          $resultMsg += "In both groups: $($inBoth.Count)`n"
          $resultMsg += "Only in $($grp.Name): $($onlyInGroup1.Count)`n"
          $resultMsg += "Only in ${compareGroupName}: $($onlyInGroup2.Count)"

          & $showModalFunc2 "Group Comparison" $resultMsg
        }.GetNewClosure())
        $compareDlg.AddButton($btnSelect)
        $btnCancel = [Terminal.Gui.Button]::new("Cancel")
        $btnCancel.add_Clicked({
          [Terminal.Gui.Application]::RequestStop()
        }.GetNewClosure())
        $compareDlg.AddButton($btnCancel)
        [Terminal.Gui.Application]::Run($compareDlg)
      }.GetNewClosure())
      $view.Add($btnCompare)
    }
  }

  ## ----------{ Apply Logic }---------
  $applyLogic = {
    param($data, $state)

    ## Use $data instead of $group
    $grp = $data

    ## Capture funcitons for this closure
    $showModalFunc = ${function:Show-Modal}
    $demoMode      = $Script:DemoMode

    try {
      $changesMade = $false
      if ($state.txtDescription) {
        $newDescription = $state.txtDescription.Text.ToString().Trim()
        if ($newDescription -ne $grp.Description) {
          if ($demoMode) {
            $grp.Description = $newDescription
          } else {
            Set-UnifiedObject -ObjectType Group -Object $grp -Properties @{ Description = $newDescription }
          }
          $changesMade = $true
        }
      }
      if ($state.txtEmail) {
        $newEmail = $state.txtEmail.Text.ToString().Trim()
        if ($newEmail -ne $grp.Email) {
          if ($demoMode) {
            $grp.Email = $newEmail
          } else {
            Set-UnifiedObject -ObjectType Group -Object $grp -Properties @{ mail = $newEmail; Email = $newEmail }
          }
          $changesMade = $true
        }
      }
      if ($state.txtManagedBy) {
        $newManagedBy = $state.txtManagedBy.Text.ToString().Trim()
        if ($newManagedBy -ne $grp.ManagedBy) {
          if ($demoMode) {
            $grp.ManagedBy = $newManagedBy
          } else {
            Set-UnifiedObject -ObjectType Group -Object $grp -Properties @{ ManagedBy = $newManagedBy }
          }
          $changesMade = $true
        }
      }
      if ($changesMade) {
        & $showModalFunc "Success" "Changes applied successfully"
      } else {
        & $showModalFunc "Info" "No changes to apply"
      }
    } catch {
      & $showModalFunc "Error" "Failed to apply changes:`n$($_.Exception.Message)"
    }
  }

  ## ----------{ Create Dialog }---------
  $tabs = @($generalTab, $membersTab, $reportTab)
  New-PropertiesDialog -Title "Group Properties - $($group.Name)" -Width 100 -Height 40 -Tabs $tabs -Data $group -OnApply $applyLogic -IncludeSearchTab $true -SearchTabConfig @{ObjectType='Group'}
}

# ----------{ Show Computer Properties }----------
function Show-ComputerPropertiesDialog {
  param([string]$computerName)

  Debug-Log "Showing computer properties for: $computerName" -Type "Insight"
  $computer = $Script:Computers | Where-Object { $_.Name -eq $computerName } | Select-Object -First 1
  if (-not $computer) {
    Debug-Log "Computer NOT found in Script:Computers for name: $computerName" -Type "Insight"
    Show-Modal "Not Found" "Computer '$computerName' not found"
    return
  }
  Debug-Log "Computer found: $($computer.Name)" -Type "Insight"

  ## ----------{ General Tab }---------
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $data, $state)
      $comp = $data
      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Computer Information"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Name:" -FieldName 'txtName' -State $state -Value ($comp.Name ?? "") -FieldX 25
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "DNS Host Name:" -FieldName 'txtDNS' -State $state -Value ($comp.DNSHostName ?? "") -FieldX 25 -IsTextField $false
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Domain:" -FieldName 'txtDomain' -State $state -Value ($comp.Domain ?? "") -FieldX 25 -IsTextField $false
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Description:" -FieldName 'txtDescription' -State $state -Value ($comp.Description ?? "") -FieldX 25 -Width 65
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Operating System"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "OS:" -FieldName 'txtOS' -State $state -Value ($comp.OperatingSystem ?? "") -FieldX 25 -IsTextField $false
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Version:" -FieldName 'txtOSVer' -State $state -Value ($comp.OperatingSystemVersion ?? "") -FieldX 25 -IsTextField $false
      if ($comp.OperatingSystemServicePack) { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Service Pack:" -FieldName 'txtSP' -State $state -Value $comp.OperatingSystemServicePack -FieldX 25 -IsTextField $false }
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Network"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "IPv4 Address:" -FieldName 'txtIPv4' -State $state -Value ($comp.IPv4Address ?? "") -FieldX 25 -IsTextField $false
      if ($comp.IPv6Address) { Add-LabelAndField -View $view -Y ([ref]$y) -Label "IPv6 Address:" -FieldName 'txtIPv6' -State $state -Value $comp.IPv6Address -FieldX 25 -IsTextField $false }
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Location:" -FieldName 'txtLocation' -State $state -Value ($comp.Location ?? "") -FieldX 25 -Width 65
    }
  }

  ## ----------{ Account Tab }---------
  $accountTab = @{
    Name = "Account"
    Builder = {
      param($view, $data, $state)
      $comp = $data
      $formatDateFunc = ${function:Format-DateSafe}
      $showAuditLogFunc = ${function:Show-AuditLogDialog}
      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Account Status"
      $state.chkEnabled = [Terminal.Gui.CheckBox]::new("Computer Account Enabled")
      $state.chkEnabled.X=4; $state.chkEnabled.Y=$y; $state.chkEnabled.Checked=($comp.Enabled??$true)
      $view.Add($state.chkEnabled); $y+=1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Password Settings"
      $state.chkPasswordExpired = [Terminal.Gui.CheckBox]::new("Password Expired")
      $state.chkPasswordExpired.X=4; $state.chkPasswordExpired.Y=$y; $state.chkPasswordExpired.Checked=($comp.PasswordExpired??$false); $state.chkPasswordExpired.Enabled=$false
      $view.Add($state.chkPasswordExpired); $y+=1
      $state.chkPasswordNeverExpires = [Terminal.Gui.CheckBox]::new("Password never expires")
      $state.chkPasswordNeverExpires.X=4; $state.chkPasswordNeverExpires.Y=$y; $state.chkPasswordNeverExpires.Checked=($comp.PasswordNeverExpires??$false); $state.chkPasswordNeverExpires.Enabled=$false
      $view.Add($state.chkPasswordNeverExpires); $y+=1
      $state.chkCannotChangePassword = [Terminal.Gui.CheckBox]::new("Cannot change password")
      $state.chkCannotChangePassword.X=4; $state.chkCannotChangePassword.Y=$y; $state.chkCannotChangePassword.Checked=($comp.CannotChangePassword??$false); $state.chkCannotChangePassword.Enabled=$false
      $view.Add($state.chkCannotChangePassword); $y+=1
      $state.chkPasswordNotRequired = [Terminal.Gui.CheckBox]::new("Password not required")
      $state.chkPasswordNotRequired.X=4; $state.chkPasswordNotRequired.Y=$y; $state.chkPasswordNotRequired.Checked=($comp.PasswordNotRequired??$false); $state.chkPasswordNotRequired.Enabled=$false
      $view.Add($state.chkPasswordNotRequired); $y+=2
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Account Details" -SpaceBefore 0
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "SAM Account:" -FieldName 'txtSAM' -State $state -Value ($comp.SamAccountName ?? "") -FieldX 25 -IsTextField $false
      if ($comp.PasswordLastSet) {
        $dateStr = & $formatDateFunc $comp.PasswordLastSet 'yyyy-MM-dd HH:mm'
        $lbl = [Terminal.Gui.Label]::new("Password last set: $dateStr")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
      if ($comp.LastLogonDate) {
        $dateStr = & $formatDateFunc $comp.LastLogonDate 'yyyy-MM-dd HH:mm'
        $lbl = [Terminal.Gui.Label]::new("Last logon: $dateStr")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
      if ($comp.logonCount) {
        $lbl = [Terminal.Gui.Label]::new("Logon count: $($comp.logonCount)")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
      if ($comp.AccountExpirationDate) {
        $dateStr = & $formatDateFunc $comp.AccountExpirationDate 'yyyy-MM-dd'
        $lbl = [Terminal.Gui.Label]::new("Account expires: $dateStr")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=2
      }
      $y += 1
      $btnAuditLog = [Terminal.Gui.Button]::new("View Audit Log...")
      $btnAuditLog.X = 2; $btnAuditLog.Y = $y
      $btnAuditLog.add_Clicked({ & $showAuditLogFunc -Object $comp -ObjectType 'Computer' }.GetNewClosure())
      $view.Add($btnAuditLog)
    }
  }

  ## ----------{ Security Tab }---------
  $securityTab = @{
    Name = "Security"
    Builder = {
      param($view, $data, $state)
      $comp = $data
      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Delegation Settings"
      $state.chkTrustedForDelegation = [Terminal.Gui.CheckBox]::new("Trusted for delegation")
      $state.chkTrustedForDelegation.X=4; $state.chkTrustedForDelegation.Y=$y; $state.chkTrustedForDelegation.Checked=($comp.TrustedForDelegation??$false); $state.chkTrustedForDelegation.Enabled=$false
      $view.Add($state.chkTrustedForDelegation); $y+=1
      $state.chkTrustedToAuth = [Terminal.Gui.CheckBox]::new("Trusted to authenticate for delegation")
      $state.chkTrustedToAuth.X=4; $state.chkTrustedToAuth.Y=$y; $state.chkTrustedToAuth.Checked=($comp.TrustedToAuthForDelegation??$false); $state.chkTrustedToAuth.Enabled=$false
      $view.Add($state.chkTrustedToAuth); $y+=1
      $state.chkAccountNotDelegated = [Terminal.Gui.CheckBox]::new("Account not delegated")
      $state.chkAccountNotDelegated.X=4; $state.chkAccountNotDelegated.Y=$y; $state.chkAccountNotDelegated.Checked=($comp.AccountNotDelegated??$false); $state.chkAccountNotDelegated.Enabled=$false
      $view.Add($state.chkAccountNotDelegated); $y+=1
      $state.chkNoPreAuth = [Terminal.Gui.CheckBox]::new("Does not require Kerberos preauthentication")
      $state.chkNoPreAuth.X=4; $state.chkNoPreAuth.Y=$y; $state.chkNoPreAuth.Checked=($comp.DoesNotRequirePreAuth??$false); $state.chkNoPreAuth.Enabled=$false
      $view.Add($state.chkNoPreAuth); $y+=1
      $state.chkUseDES = [Terminal.Gui.CheckBox]::new("Use DES encryption types")
      $state.chkUseDES.X=4; $state.chkUseDES.Y=$y; $state.chkUseDES.Checked=($comp.UseDESKeyOnly??$false); $state.chkUseDES.Enabled=$false
      $view.Add($state.chkUseDES); $y+=2
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Kerberos Encryption" -SpaceBefore 0
      if ($comp.KerberosEncryptionType) {
        $kerbTypes = $comp.KerberosEncryptionType -join ', '
        $lbl = [Terminal.Gui.Label]::new("Encryption types: $kerbTypes")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
      if ($comp.'msDS-SupportedEncryptionTypes') {
        $lbl = [Terminal.Gui.Label]::new("Supported encryption: $($comp.'msDS-SupportedEncryptionTypes')")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Identifiers"
      $sid = $comp.SID ?? $comp.objectSid
      if ($sid) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "SID:" -FieldName 'txtSID' -State $state -Value $sid.ToString() -FieldX 25 -IsTextField $false
      }
      if ($comp.ObjectGUID) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Object GUID:" -FieldName 'txtGUID' -State $state -Value $comp.ObjectGUID.ToString() -FieldX 25 -IsTextField $false
      }
    }
  }

  ## ----------{ Member Of Tab }---------
  $memberOfTab = @{
    Name = "Member Of"
    Builder = {
      param($view, $data, $state)
      $comp = $data
      $showModalFunc = ${function:Show-Modal}
      $showEditGroupFunc = ${function:Show-EditGroupMembershipDialog}
      $demoMode = $Script:DemoMode
      $y = 1
      $lbl = [Terminal.Gui.Label]::new("Group Memberships:")
      $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $state.lstGroups = [Terminal.Gui.ListView]::new()
      $state.lstGroups.X = 2; $state.lstGroups.Y = $y
      $state.lstGroups.Width = [Terminal.Gui.Dim]::Fill(2)
      $state.lstGroups.Height = 20
      $state.groupList = @()
      if ($comp.MemberOf) { $state.groupList = $comp.MemberOf | ForEach-Object { if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ } } }
      if ($state.groupList.Count -gt 0) {
        $state.lstGroups.SetSource($state.groupList)
      } else {
        $state.lstGroups.SetSource(@("(No group memberships)"))
      }
      $view.Add($state.lstGroups)
      $pseudoUser = [PSCustomObject]@{
        Name = $comp.Name
        SamAccountName = $comp.SamAccountName
        MemberOf = $comp.MemberOf
        Groups = $state.groupList
      }
      $btnAdd = [Terminal.Gui.Button]::new("Add to Group...")
      $btnAdd.X = 2; $btnAdd.Y = 23
      $btnAdd.add_Clicked({
        & $showEditGroupFunc -User $pseudoUser -OnUpdate {
          $refreshedGroups = @()
          if ($comp.MemberOf) {
            $refreshedGroups = $comp.MemberOf | ForEach-Object { if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ } }
          }
          if ($refreshedGroups.Count -gt 0) {
            $state.lstGroups.SetSource($refreshedGroups)
          } else {
            $state.lstGroups.SetSource(@("(No group memberships)"))
          }
          $state.groupList = $refreshedGroups
        }
      }.GetNewClosure())
      $view.Add($btnAdd)
      $btnRemove = [Terminal.Gui.Button]::new("Remove from Group")
      $btnRemove.X = 22; $btnRemove.Y = 23
      $btnRemove.add_Clicked({
        $selectedIndex = $state.lstGroups.SelectedItem
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $state.groupList.Count) {
          $selectedGroup = $state.groupList[$selectedIndex]
          if ($selectedGroup -eq "(No group memberships)") {
            & $showModalFunc "Info" "No group selected"
            return
          }
          $confirmDlg = & $showModalFunc "Confirm Removal" "Remove $($comp.Name) from group '$selectedGroup'?" -YesNo
          if ($confirmDlg -eq 0) {
            try {
              if ($demoMode) {
                $comp.MemberOf = $comp.MemberOf | Where-Object { $_ -notmatch "CN=$selectedGroup," }
              } else {
                Remove-ADGroupMember -Identity $selectedGroup -Members $comp.SamAccountName -Confirm:$false
              }
              $updatedGroups = @()
              if ($comp.MemberOf) {
                $updatedGroups = $comp.MemberOf | ForEach-Object {
                  if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
                }
              }
              if ($updatedGroups.Count -gt 0) {
                $state.lstGroups.SetSource($updatedGroups)
              } else {
                $state.lstGroups.SetSource(@("(No group memberships)"))
              }
              $state.groupList = $updatedGroups
              & $showModalFunc "Success" "Successfully removed $($comp.Name) from group '$selectedGroup'"
            } catch {
              & $showModalFunc "Error" "Failed to remove from group:`n$($_.Exception.Message)"
            }
          }
        } else {
          & $showModalFunc "Info" "Please select a group to remove"
        }
      }.GetNewClosure())
      $view.Add($btnRemove)
    }
  }

  ## ----------{ SPNs Tab }---------
  $spnTab = @{
    Name = "SPNs"
    Builder = {
      param($view, $data, $state)
      $comp = $data
      $y = 1
      $lbl = [Terminal.Gui.Label]::new("Service Principal Names:")
      $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2
      $state.lstSPNs = [Terminal.Gui.ListView]::new()
      $state.lstSPNs.X = 2; $state.lstSPNs.Y = $y
      $state.lstSPNs.Width = [Terminal.Gui.Dim]::Fill(2)
      $state.lstSPNs.Height = [Terminal.Gui.Dim]::Fill(2)
      $spnList = @()
      if ($comp.ServicePrincipalNames) {
        $spnList = $comp.ServicePrincipalNames
      } elseif ($comp.servicePrincipalName) {
        $spnList = $comp.servicePrincipalName
      }
      if ($spnList.Count -gt 0) {
        $state.lstSPNs.SetSource($spnList)
      } else {
        $state.lstSPNs.SetSource(@("(No SPNs configured)"))
      }
      $view.Add($state.lstSPNs)
    }
  }

  ## ----------{ LAPS Tab (with password masking) }---------
  $lapsTab = @{
    Name = "LAPS"
    Builder = {
      param($view, $data, $state)

      $comp = $data
      $demoMode = $Script:DemoMode
      $formatDateFunc = ${function:Format-DateSafe}
      $showModalFunc = ${function:Show-Modal}
      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Local Administrator Password Solution"

      $lapsEnabled = $false
      $lapsPassword = ""
      $lapsExpiry = ""
      $lapsAccountName = "Administrator"
      $lapsType = ""

      if (-not $demoMode) {
        ## Production Mode - Query AD
        try {
          ## Try Windows LAPS first (new)
          $lapsData = Get-ADComputer -Identity $comp.Name -Properties 'msLAPS-Password','msLAPS-PasswordExpirationTime','msLAPS-AccountName','ms-Mcs-AdmPwd','ms-Mcs-AdmPwdExpirationTime' -ErrorAction SilentlyContinue

          ## Check Windows LAPS (new)
          if ($lapsData.'msLAPS-Password') {
            $lapsEnabled = $true
            $lapsType = "Windows LAPS"
            $lapsPassword = $lapsData.'msLAPS-Password'
            $lapsAccountName = if ($lapsData.'msLAPS-AccountName') { $lapsData.'msLAPS-AccountName' } else { "Administrator" }
            $lapsExpiry = if ($lapsData.'msLAPS-PasswordExpirationTime') {
              try {
                [DateTime]::FromFileTimeUtc($lapsData.'msLAPS-PasswordExpirationTime').ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
              } catch {
                "Unknown"
              }
            } else { "Unknown" }
          }
          ## Check Legacy LAPS (old)
          elseif ($lapsData.'ms-Mcs-AdmPwd') {
            $lapsEnabled = $true
            $lapsType = "Legacy LAPS"
            $lapsPassword = $lapsData.'ms-Mcs-AdmPwd'
            $lapsExpiry = if ($lapsData.'ms-Mcs-AdmPwdExpirationTime') {
              [DateTime]::FromFileTime($lapsData.'ms-Mcs-AdmPwdExpirationTime').ToString('yyyy-MM-dd HH:mm:ss')
            } else { "Unknown" }
          }
        } catch {
          ## Error querying - will show as not enabled
        }

      } else {
        ## DEMO MODE - Check computer object properties

        ## Check Windows LAPS (new) - including alternative property names
        if ($comp.'msLAPS-Password' -or $comp.LAPSPassword) {
          $lapsEnabled = $true
          $lapsType = "Windows LAPS"
          $lapsPassword = $comp.'msLAPS-Password' ?? $comp.LAPSPassword
          $lapsAccountName = $comp.'msLAPS-AccountName' ?? $comp.LAPSAccountName ?? "Administrator"

          ## Handle expiration time
          if ($comp.'msLAPS-PasswordExpirationTime') {
            try {
              $lapsExpiry = [DateTime]::FromFileTimeUtc($comp.'msLAPS-PasswordExpirationTime').ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
            } catch {
              $lapsExpiry = "Unknown"
            }
          } elseif ($comp.LAPSPasswordExpiration) {
            $lapsExpiry = & $formatDateFunc $comp.LAPSPasswordExpiration 'yyyy-MM-dd HH:mm:ss'
          } else {
            $lapsExpiry = "Unknown"
          }
        }
        ## Check Legacy LAPS (old)
        elseif ($comp.'ms-Mcs-AdmPwd') {
          $lapsEnabled = $true
          $lapsType = "Legacy LAPS"
          $lapsPassword = $comp.'ms-Mcs-AdmPwd'
          if ($comp.'ms-Mcs-AdmPwdExpirationTime') {
            try {
              $lapsExpiry = [DateTime]::FromFileTime($comp.'ms-Mcs-AdmPwdExpirationTime').ToString('yyyy-MM-dd HH:mm:ss')
            } catch {
              $lapsExpiry = "Unknown"
            }
          } else {
            $lapsExpiry = "Unknown"
          }
        }
      }
      if ($lapsEnabled) {
        $lbl = [Terminal.Gui.Label]::new("LAPS Status: ✓ Enabled ($lapsType)")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=2
        $lbl = [Terminal.Gui.Label]::new("Password Expires: $lapsExpiry")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
        ## Additional LAPS info if available
        if ($comp.LAPSPasswordLastSet) {
          $lastSet = & $formatDateFunc $comp.LAPSPasswordLastSet 'yyyy-MM-dd HH:mm:ss'
          $lbl = [Terminal.Gui.Label]::new("Password Last Set: $lastSet")
          $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
        }
        $y += 1

        ## View Password Button
        $btnViewPassword = [Terminal.Gui.Button]::new("👁 View Password...")
        $btnViewPassword.X = 4
        $btnViewPassword.Y = $y
        $btnViewPassword.add_Clicked({
          $details = @"
LAPS Password Details
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Computer: $($comp.Name)
LAPS Type: $lapsType
Account Name: $lapsAccountName

Password:
$lapsPassword

Password Expires: $lapsExpiry

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⚠ SECURITY WARNING ⚠

This password provides full local administrator
access to this computer. Handle with extreme care
and store securely.

Do not share via email or unsecured channels.
Log access to this password per your organization's
security policy.
"@
          & $showModalFunc "LAPS Password" $details
        }.GetNewClosure())
        $view.Add($btnViewPassword)

        ## Copy Password Button
        $btnCopyPassword = [Terminal.Gui.Button]::new("📋 Copy Password")
        $btnCopyPassword.X = 28
        $btnCopyPassword.Y = $y
        $btnCopyPassword.add_Clicked({
          try {
            if ($IsWindows) { Set-Clipboard -Value $lapsPassword
            } elseif ($IsLinux) { $lapsPassword | xclip -selection clipboard
            } elseif ($IsMacOS) { $lapsPassword | pbcopy
            }
            & $showModalFunc "Success" "LAPS password copied to clipboard"
          } catch {
            & $showModalFunc "Error" "Failed to copy to clipboard:`n$($_.Exception.Message)"
          }
        }.GetNewClosure())
        $view.Add($btnCopyPassword)

        $y += 2
        $lbl = [Terminal.Gui.Label]::new("⚠ This password provides full local administrator access")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
        $lbl = [Terminal.Gui.Label]::new("  Handle with appropriate security controls")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
        $lbl = [Terminal.Gui.Label]::new("  Log all password access per security policy")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)

      } else {
        $lbl = [Terminal.Gui.Label]::new("LAPS Status: ✗ Not Enabled")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=2
        $lbl = [Terminal.Gui.Label]::new("This computer does not have LAPS configured or you")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
        $lbl = [Terminal.Gui.Label]::new("do not have permissions to view the password.")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl)
      }
    }
  }

  ## ----------{ BitLocker Tab }---------
  $bitlockerTab = @{
    Name = "BitLocker"
    Builder = {
      param($view, $data, $state)
      $comp = $data
      $showModalFunc = ${function:Show-Modal}
      $formatDateFunc = ${function:Format-DateSafe}
      $rawBitLockerRecovery = $Script:rawBitLockerRecovery
      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "BitLocker Recovery Information"
      $computerDN = $comp.DistinguishedName
      $recoveryKeys = @()
      if ($rawBitLockerRecovery) { $recoveryKeys = $rawBitLockerRecovery | Where-Object { $_.DN -match [regex]::Escape($computerDN) } }
      if ($recoveryKeys.Count -eq 0 -and $comp.BitLockerRecoveryKey) {
        $recoveryKeys = @([PSCustomObject]@{
          Name             = "Recovery Key"
          RecoveryPassword = $comp.BitLockerRecoveryKey
          RecoveryGuid     = "N/A"
          VolumeGuid       = "N/A"
          Created          = $comp.Created ?? (Get-Date)
        })
      }
      if ($recoveryKeys.Count -gt 0) {
        $lbl = [Terminal.Gui.Label]::new("Found $($recoveryKeys.Count) recovery key(s) for this computer")
        $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 2
        $state.lstRecoveryKeys = [Terminal.Gui.ListView]::new()
        $state.lstRecoveryKeys.X = 4; $state.lstRecoveryKeys.Y = $y
        $state.lstRecoveryKeys.Width = [Terminal.Gui.Dim]::Fill(2)
        $state.lstRecoveryKeys.Height = 8
        $keyList = @($recoveryKeys | ForEach-Object {
          $created = & $formatDateFunc $_.Created 'yyyy-MM-dd HH:mm'
          "$($_.Name) - Created: $created"
        })
        $state.lstRecoveryKeys.SetSource($keyList)
        $view.Add($state.lstRecoveryKeys)
        $y += 9
        $btnView = [Terminal.Gui.Button]::new("View Selected Key...")
        $btnView.X = 4; $btnView.Y = $y
        $btnView.add_Clicked({
          $selectedIndex = $state.lstRecoveryKeys.SelectedItem
          if ($selectedIndex -ge 0 -and $selectedIndex -lt $recoveryKeys.Count) {
            $key = $recoveryKeys[$selectedIndex]
            $createdStr = & $formatDateFunc $key.Created 'yyyy-MM-dd HH:mm:ss'
            $details = @"
BitLocker Recovery Key Details
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Computer: $($comp.Name)
Key Name: $($key.Name)

Recovery Password:
$($key.RecoveryPassword)

Recovery GUID: $($key.RecoveryGuid)
Volume GUID:   $($key.VolumeGuid)

Created: $createdStr

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⚠ SECURITY WARNING ⚠

This recovery password provides full access to
encrypted data on this volume. Handle with extreme
care and store securely.

Do not share via email or unsecured channels.
"@
            & $showModalFunc "BitLocker Recovery Key" $details
          } else {
            & $showModalFunc "Info" "Please select a recovery key to view"
          }
        }.GetNewClosure())
        $view.Add($btnView)
        $btnCopy = [Terminal.Gui.Button]::new("Copy to Clipboard")
        $btnCopy.X = 28; $btnCopy.Y = $y
        $btnCopy.add_Clicked({
          $selectedIndex = $state.lstRecoveryKeys.SelectedItem
          if ($selectedIndex -ge 0 -and $selectedIndex -lt $recoveryKeys.Count) {
            $key = $recoveryKeys[$selectedIndex]
            try {
              if ($IsWindows) { Set-Clipboard -Value $key.RecoveryPassword
              } elseif ($IsLinux) { $key.RecoveryPassword | xclip -selection clipboard
              } elseif ($IsMacOS) { $key.RecoveryPassword | pbcopy
              }
              & $showModalFunc "Success" "Recovery password copied to clipboard"
            } catch {
              & $showModalFunc "Error" "Failed to copy to clipboard:`n$($_.Exception.Message)"
            }
          } else {
            & $showModalFunc "Info" "Please select a recovery key to copy"
          }
        }.GetNewClosure())
        $view.Add($btnCopy)
        $y += 2
        $lbl = [Terminal.Gui.Label]::new("⚠ Recovery passwords provide full access to encrypted volumes")
        $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 1
        $lbl = [Terminal.Gui.Label]::new("  Handle with appropriate security controls")
        $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl)
      } else {
        $lbl = [Terminal.Gui.Label]::new("No BitLocker recovery keys found for this computer")
        $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 2
        $lbl = [Terminal.Gui.Label]::new("This may indicate:")
        $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 1
        $bullets = @(
          "• BitLocker is not enabled on this computer",
          "• Recovery keys are not backed up to Active Directory",
          "• You do not have permissions to view recovery keys"
        )
        foreach ($bullet in $bullets) {
          $lbl = [Terminal.Gui.Label]::new("  $bullet")
          $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 1
        }
      }
    }
  }

  ## ----------{ Advanced Tab }---------
  $advancedTab = @{
    Name = "Advanced"
    Builder = {
      param($view, $data, $state)
      $comp = $data
      $formatDateFunc = ${function:Format-DateSafe}
      $y = 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Distinguished Name:" -FieldName 'txtDN' -State $state -Value ($comp.DistinguishedName ?? "") -LabelX 2 -FieldX 2 -Width 90 -ReadOnly $true
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Canonical Name:" -FieldName 'txtCN' -State $state -Value ($comp.CanonicalName ?? "") -LabelX 2 -FieldX 2 -Width 90 -ReadOnly $true
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Timestamps"
      $created = & $formatDateFunc ($comp.Created ?? $comp.whenCreated) 'yyyy-MM-dd HH:mm:ss'
      $lbl = [Terminal.Gui.Label]::new("Created: $created")
      $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      $modified = & $formatDateFunc ($comp.Modified ?? $comp.whenChanged) 'yyyy-MM-dd HH:mm:ss'
      $lbl = [Terminal.Gui.Label]::new("Modified: $modified")
      $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      if ($comp.LastBootUpTime) {
        $bootTime = & $formatDateFunc $comp.LastBootUpTime 'yyyy-MM-dd HH:mm:ss'
        $lbl = [Terminal.Gui.Label]::new("Last Boot: $bootTime")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Update Sequence Numbers"
      if ($comp.uSNCreated) {
        $lbl = [Terminal.Gui.Label]::new("USN Created: $($comp.uSNCreated)")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
      if ($comp.uSNChanged) {
        $lbl = [Terminal.Gui.Label]::new("USN Changed: $($comp.uSNChanged)")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Additional Properties"
      if ($comp.isCriticalSystemObject) {
        $lbl = [Terminal.Gui.Label]::new("✓ Critical System Object")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
      if ($comp.ProtectedFromAccidentalDeletion) {
        $lbl = [Terminal.Gui.Label]::new("✓ Protected from Accidental Deletion")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
      if ($comp.PrimaryGroup) {
        $lbl = [Terminal.Gui.Label]::new("Primary Group: $($comp.PrimaryGroup)")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
    }
  }

  ## ----------{ Apply Logic }---------
  $applyLogic = {
    param($data, $state)
    $comp = $data
    $showModalFunc = ${function:Show-Modal}
    $demoMode = $Script:DemoMode
    try {
      $changesMade = $false
      if ($state.txtDescription) {
        $newDescription = $state.txtDescription.Text.ToString().Trim()
        if ($newDescription -ne $comp.Description) {
          if ($demoMode) {
            $comp.Description = $newDescription
          } else {
            Set-UnifiedObject -ObjectType Computer -Object $comp -Properties @{ Description = $newDescription }
            $comp.Description = $newDescription
          }
          $changesMade = $true
        }
      }
      if ($state.txtLocation) {
        $newLocation = $state.txtLocation.Text.ToString().Trim()
        if ($newLocation -ne $comp.Location) {
          if ($demoMode) {
            $comp.Location = $newLocation
          } else {
            Set-UnifiedObject -ObjectType Computer -Object $comp -Properties @{ Location = $newLocation }
            $comp.Location = $newLocation
          }
          $changesMade = $true
        }
      }
      if ($changesMade) {
        & $showModalFunc "Success" "Computer changes applied successfully"
      } else {
        & $showModalFunc "Info" "No changes to apply"
      }
    } catch {
      & $showModalFunc "Error" "Failed to apply computer changes:`n$($_.Exception.Message)"
    }
  }

  ## ----------{ Create Dialog }---------
  $tabs = @($generalTab, $accountTab, $securityTab, $memberOfTab, $spnTab, $lapsTab, $bitlockerTab, $advancedTab)
  New-PropertiesDialog -Title "Computer Properties - $($computer.Name)" -Width 100 -Height 40 -Tabs $tabs -Data $computer -OnApply $applyLogic -IncludeSearchTab $true -SearchTabConfig @{ObjectType='Computer'}
}

function Show-GPODetailsDialog {
  param($GPO)

  $gpoName = if ($GPO.DisplayName) { $GPO.DisplayName } else { $GPO.Name }
  Debug-Log "Showing details for GPO: $gpoName" -Type "Insight"
  ## Build details text
  $details = @"
Group Policy Object Details
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Name: $gpoName

Description:
$(if ($GPO.Description) { $GPO.Description } else { "(No description)" })

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Path: $(if ($GPO.GPCFileSysPath) { $GPO.GPCFileSysPath } else { "N/A" })

Version Number: $(if ($GPO.VersionNumber) { $GPO.VersionNumber } else { "N/A" })

Functionality Version: $(if ($GPO.GPCFunctionalityVersion) { $GPO.GPCFunctionalityVersion } else { "N/A" })

Flags: $(if ($GPO.Flags) { $GPO.Flags } else { "N/A" })

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Machine Extensions:
$(if ($GPO.GPCMachineExtensionNames) { $GPO.GPCMachineExtensionNames } else { "(None)" })

User Extensions:
$(if ($GPO.GPCUserExtensionNames) { $GPO.GPCUserExtensionNames } else { "(None)" })

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Distinguished Name:
$(if ($GPO.DN) { $GPO.DN } else { "N/A" })

Created: $(if ($GPO.Created) { $GPO.Created } else { "Unknown" })
Modified: $(if ($GPO.Modified) { $GPO.Modified } else { "Unknown" })
"@

  ## Create dialog
  $dialog = [Terminal.Gui.Dialog]::new()
  $dialog.Title = "GPO Details - $gpoName"
  $dialog.Width = 100
  $dialog.Height = 35

  $txtDetails = [Terminal.Gui.TextView]::new()
  $txtDetails.X = 1
  $txtDetails.Y = 1
  $txtDetails.Width = [Terminal.Gui.Dim]::Fill(1)
  $txtDetails.Height = [Terminal.Gui.Dim]::Fill(3)
  $txtDetails.ReadOnly = $true
  $txtDetails.Text = $details
  $dialog.Add($txtDetails)

  ## Buttons
  $btnCopy = [Terminal.Gui.Button]::new("Copy Details")
  $btnCopy.X = 2
  $btnCopy.Y = [Terminal.Gui.Pos]::Bottom($txtDetails) + 1
  $btnCopy.add_Clicked({
    try {
      if ($IsWindows) { Set-Clipboard -Value $details
      } elseif ($IsLinux) { $details | xclip -selection clipboard
      } elseif ($IsMacOS) { $details | pbcopy
      }
      Show-Modal "Success" "GPO details copied to clipboard"
    } catch {
      Show-Modal "Error" "Failed to copy to clipboard:`n$($_.Exception.Message)"
    }
  }.GetNewClosure())
  $dialog.Add($btnCopy)

  $btnClose = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = 20
  $btnClose.Y = [Terminal.Gui.Pos]::Bottom($txtDetails) + 1
  $btnClose.add_Clicked({
    $dialog.RequestStop()
  })
  $dialog.Add($btnClose)

  [Terminal.Gui.Application]::Run($dialog)
}

## ----------{ List dialog helpers - For GPO, DNS, Trusts, etc. }----------
function New-ListDialog {
  <#
  .SYNOPSIS
  Creates a filterable list dialog with standard buttons

  .PARAMETER Title
  Dialog title

  .PARAMETER Items
  Array of items to display

  .PARAMETER FormatItem
  Scriptblock to format each item for display

  .PARAMETER OnView
  Scriptblock when View Details is clicked

  .PARAMETER OnExport
  Scriptblock when Export is clicked

  .PARAMETER OnRefresh
  Scriptblock when Refresh is clicked

  .PARAMETER Width
  Dialog width (default 120)

  .PARAMETER Height
  Dialog height (default 40)
  #>

  param(
    [Parameter(Mandatory)]
    [string]$Title,
    [Parameter(Mandatory)]
    [array]$Items,
    [scriptblock]$FormatItem = { param($item) $item.ToString() },
    [scriptblock]$OnView,
    [scriptblock]$OnExport,
    [scriptblock]$OnRefresh,
    [int]$Width = 120,
    [int]$Height = 40,
    [string]$FilterHelp = "(Filter by name or description)"
  )

  $dialog = [Terminal.Gui.Dialog]::new()
  $dialog.Title  = $Title
  $dialog.Width  = $Width
  $dialog.Height = $Height
  $y = 1

  ## Summary
  $lblSummary = [Terminal.Gui.Label]::new("Total Items: $($Items.Count)")
  $lblSummary.X = 2; $lblSummary.Y = $y
  $dialog.Add($lblSummary)
  $y += 2

  ## Filter
  $lblFilter = [Terminal.Gui.Label]::new("Filter:")
  $lblFilter.X = 2; $lblFilter.Y = $y
  $dialog.Add($lblFilter)

  $txtFilter = [Terminal.Gui.TextField]::new("")
  $txtFilter.X = 10; $txtFilter.Y = $y; $txtFilter.Width = 40
  $dialog.Add($txtFilter)

  $lblFilterHelp = [Terminal.Gui.Label]::new($FilterHelp)
  $lblFilterHelp.X = 52; $lblFilterHelp.Y = $y
  $dialog.Add($lblFilterHelp)
  $y += 2

  ## List label
  $lblList = [Terminal.Gui.Label]::new("Items:")
  $lblList.X = 2; $lblList.Y = $y
  $dialog.Add($lblList)
  $y += 1

  ## List view
  $lstItems = [Terminal.Gui.ListView]::new()
  $lstItems.X = 2; $lstItems.Y = $y
  $lstItems.Width = [Terminal.Gui.Dim]::Fill(2)
  $lstItems.Height = 20

  ## Format items
  $displayList = $Items | ForEach-Object { & $FormatItem $_ }
  $lstItems.SetSource($displayList)
  $dialog.Add($lstItems)
  $y += 21

  ## Filter functionality
  $originalList = $displayList
  $txtFilter.add_TextChanged({
    $filterText = $txtFilter.Text.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($filterText)) {
      $lstItems.SetSource($originalList)
    } else {
      $filtered = $originalList | Where-Object { $_ -match "(?i)$filterText" }
      if ($filtered) {
        $lstItems.SetSource($filtered)
      } else {
        $lstItems.SetSource(@("(No matches)"))
      }
    }
  }.GetNewClosure())

  ## Buttons
  $btnX = 2
  if ($OnView) {
    $btnView = [Terminal.Gui.Button]::new("View Details...")
    $btnView.X = $btnX; $btnView.Y = $y
    $btnView.add_Clicked({
      $selectedIndex = $lstItems.SelectedItem
      $selectedText = $lstItems.Source.ToList()[$selectedIndex]
      if ($selectedText -eq "(No matches)") {
        Show-Modal "Info" "No item selected"
        return
      }
      ## Find actual item
      $selectedItem = $null
      for ($i = 0; $i -lt $Items.Count; $i++) {
        $formatted = & $FormatItem $Items[$i]
        if ($formatted -eq $selectedText) {
          $selectedItem = $Items[$i]
          break
        }
      }
      if ($selectedItem) {
        & $OnView $selectedItem
      } else {
        Show-Modal "Error" "Could not find selected item"
      }
    }.GetNewClosure())
    $dialog.Add($btnView)
    $btnX += 20
  }
  if ($OnExport) {
    $btnExport = [Terminal.Gui.Button]::new("Export List...")
    $btnExport.X = $btnX; $btnExport.Y = $y
    $btnExport.add_Clicked({
      & $OnExport $Items
    }.GetNewClosure())
    $dialog.Add($btnExport)
    $btnX += 20
  }
  if ($OnRefresh) {
    $btnRefresh = [Terminal.Gui.Button]::new("Refresh")
    $btnRefresh.X = $btnX; $btnRefresh.Y = $y
    $btnRefresh.add_Clicked({
      $dialog.RequestStop()
      & $OnRefresh
    }.GetNewClosure())
    $dialog.Add($btnRefresh)
    $btnX += 14
  }

  $btnClose = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = $btnX; $btnClose.Y = $y
  $btnClose.add_Clicked({ $dialog.RequestStop() })
  $dialog.Add($btnClose)
  [Terminal.Gui.Application]::Run($dialog)
}

## ----------{ Cleaned up GPO list dialog }----------
function Show-GPOListDialog {
  param([string]$Domain)

  if (-not $Domain) { $Domain = $Script:CurrentDomain }
  Debug-Log "Opening GPO list for domain: $Domain" -Type "Insight"
  ## Get GPOs
  $gpoResult = Test-GPOHealth -Domain $Domain
  $gpos      = $gpoResult.GPOs

  if (-not $gpos -or $gpos.Count -eq 0) {
    Show-Modal "No GPOs" "No Group Policy Objects found for domain: $Domain"
    return
  }
  Debug-Log "Found $($gpos.Count) GPOs" -Type "Insight"
  ## Format function
  $formatGPO = {
    param($gpo)
    $name = if ($gpo.DisplayName) { $gpo.DisplayName } else { $gpo.Name }
    $version = if ($gpo.VersionNumber) { " (v$($gpo.VersionNumber))" } else { "" }
    "$name$version"
  }
  ## View handler
  $onView = {
    param($gpo)
    Show-GPODetailsDialog -GPO $gpo
  }
  ## Export handler
  $onExport = {
    param($items)
    try {
      $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
      $filename   = "gpo_list_${Domain}_$timestamp.csv"
      $exportData = $items | ForEach-Object {
        [PSCustomObject]@{
          Name        = if ($_.DisplayName) { $_.DisplayName } else { $_.Name }
          Description = $_.Description
          Path        = $_.GPCFileSysPath
          Version     = $_.VersionNumber
          Created     = $_.Created
          Modified    = $_.Modified
          DN          = $_.DN
        }
      }
      $exportData | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
      Show-Modal "Success" "Exported $($items.Count) GPOs to:`n$filename"
      Debug-Log "Exported GPO list to $filename" -Type "Success"
    } catch {
      Show-Modal "Error" "Failed to export GPO list:`n$($_.Exception.Message)"
      Debug-Log "Export failed: $($_.Exception.Message)" -Type "Problem"
    }
  }
  ## Refresh handler
  $onRefresh = { Show-GPOListDialog -Domain $Domain }
  ## Show dialog
  New-ListDialog -Title "Group Policy Objects - $Domain" -Items $gpos -FormatItem $formatGPO -OnView $onView -OnExport $onExport -OnRefresh $onRefresh -FilterHelp "(Filter by GPO name or description)"
}

## ----------[ Context Menu Handler }----------
function Show-ContextMenu {
  param(
    [array]$menuItems,
    [object]$obj,
    [string]$objType
  )

  Debug-Log "Showing context menu for $($obj.Name) (Type: $objType)" -Type "Insight"

  ## Create dialog
  $contextDialog   = [Terminal.Gui.Dialog]::new("Actions", 30, ($menuItems.Count + 4))
  $contextDialog.X = [Terminal.Gui.Pos]::Center()
  $contextDialog.Y = [Terminal.Gui.Pos]::Center()

  ## Create list view
  $listView = [Terminal.Gui.ListView]::new()
  $listView.SetSource($menuItems)
  $listView.X = 0
  $listView.Y = 0
  $listView.Width = [Terminal.Gui.Dim]::Fill()
  $listView.Height = [Terminal.Gui.Dim]::Fill(2)
  $contextDialog.Add($listView)

  ## Handle selection
  $listView.add_OpenSelectedItem({
  $selected = $menuItems[$listView.SelectedItem]
  Debug-Log "Menu item selected: $selected" -Type "Insight"
  [Terminal.Gui.Application]::RequestStop()

  if ($selected -ne "---") {
    switch ($selected) {
      "Properties" {
        switch ($objType) {
      'user' {
        Debug-Log "Showing user properties for $($obj.Name)" -Type "Insight"
        Show-UserPropertiesDialog -user $obj
      }
      'group' {
        Debug-Log "Showing group properties for $($obj.Name)" -Type "Insight"
        Show-GroupPropertiesDialog -group $obj
      }
      'computer' {
        Debug-Log "Showing computer properties for $($obj.Name)" -Type "Insight"
        Show-ComputerPropertiesDialog -computerName $obj.Name
      }
      'dc' {
        Debug-Log "Showing DC properties for $($obj.Name)" -Type "Insight"
        Show-DCPropertiesDialog -dc $obj
      }
      'ou' {
        Debug-Log "Showing OU properties for $($obj.Name)" -Type "Insight"
        Show-OUPropertiesDialog -ouname $obj.Name
      }
    }
  }
      "Reset Password"    { Show-ResetPasswordDialog -userName $obj.Name                                          }
      "Disable Account"   { Toggle-UserAccount -userName $obj.Name -disable $true                                 }
      "Enable Account"    { Toggle-UserAccount -userName $obj.Name -disable $false                                }
      "Unlock Account"    { Unlock-UserAccount -userName $obj.Name                                                }
      "Disable"           { Toggle-ComputerAccount -computerName $obj.Name -disable $true                         }
      "Enable"            { Toggle-ComputerAccount -computerName $obj.Name -disable $false                        }
      "Move to OU..."     { Invoke-ObjectOperation -Objects @($selectedNode.Tag) -Operation 'Move'                }
      "Delete"            { Invoke-ObjectOperation -Objects @($selectedNode.Tag) -Operation 'Delete'              }
      "Add Member..."     { Show-EditGroupMembershipDialog -groupName $obj.Name                                   }
      "Remove Member..."  { Show-EditGroupMembershipDialog -groupName $obj.Name                                   }
      "Check Replication" { Check-DCReplication -dcName $obj.Name                                                 }
      "Refresh"           { Refresh-Data -Domain $Script:CurrentDomain -RebuildTree -ShowModal -ShowLoadingDialog }
    }
  }
})

  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
  $contextDialog.AddButton($btnCancel)
  [Terminal.Gui.Application]::Run($contextDialog)
}

function Invoke-BulkMove {
  if ($Script:SelectedObjects.Count -eq 0) {
    Show-Modal "No Selection" "No objects selected"
    return
  }

  $objects = @()
  foreach ($objName in $Script:SelectedObjects) {
    $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
    $obj = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
    if (-not $obj) { $obj = $Script:Groups | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1 }
    if ($obj) { $objects += $obj }
  }
  if ($objects.Count -gt 0) {
    Invoke-ObjectOperation -Objects $objects -Operation 'Move' -IsBulk
    $Script:SelectedObjects = @()
    $Script:SelectionMode = $false
    Show-SelectionPanel -Parent $win
  }
}

## Debug helper - verify demo data loaded correctly. Not being called isn't critical, just needs to exist for when needed
function Test-DemoData {
  Debug-Log ("========== Demo Data Check ==========") -Type "Insight"
  Debug-Log ("Global:Users count: $($Script:Users.Count)") -Type "Insight"
  Debug-Log ("Global:DCs count: $($Script:DCs.Count)") -Type "Insight"
  Debug-Log ("") -Type "Insight"
  Debug-Log ("Users in memory:") -Type "Insight"
  foreach ($u in $Script:Users) {
    $locked = if ($u.Locked) { "🔒" } else { "" }
    $disabled = if ($u.Disabled) { "⊗" } else { "○" }
    Debug-Log ("  $disabled$locked $($u.Name) - Groups: $($u.Groups -join ', ')") -Type "Insight"
  }

  ## Call this after loading demo data to verify:
  ## Test-DemoData
}

## ----------{ Helper Functions }----------
function Show-ResetPasswordDialog {
  param([string]$userName)

  $dlg             = [Terminal.Gui.Dialog]::new("Reset Password - $userName", 60, 12)
  $lbl             = [Terminal.Gui.Label]::new("New Password:"); $lbl.X=2; $lbl.Y=1; $dlg.Add($lbl)
  $txtPwd          = [Terminal.Gui.TextField]::new(""); $txtPwd.X=18; $txtPwd.Y=1; $txtPwd.Width=35; $txtPwd.Secret=$true; $dlg.Add($txtPwd)
  $lblConfirm      = [Terminal.Gui.Label]::new("Confirm Password:"); $lblConfirm.X=2; $lblConfirm.Y=3; $dlg.Add($lblConfirm)
  $txtConfirm      = [Terminal.Gui.TextField]::new(""); $txtConfirm.X=18; $txtConfirm.Y=3; $txtConfirm.Width=35; $txtConfirm.Secret=$true; $dlg.Add($txtConfirm)
  $chkMustChange   = [Terminal.Gui.CheckBox]::new("User must change password at next logon")
  $chkMustChange.X = 2; $chkMustChange.Y=5; $chkMustChange.Checked=$true; $dlg.Add($chkMustChange)

  $btnOK = [Terminal.Gui.Button]::new("OK")
  $btnOK.add_Clicked({
    $password1 = $txtPwd.Text.ToString()
    $password2 = $txtConfirm.Text.ToString()
    if ($password1 -ne $password2) {
      Show-Modal "Error" "Passwords do not match!"
      return
    }
    if ($password1.Length -lt 8) {
      Show-Modal "Error" "Password must be at least 8 characters!"
      return
    }
    try {
      if ($Script:DemoMode) {
        Debug-Log (" Password reset for $userName (demo mode)") -Type "Insight"
        Show-Modal "Success" "Password reset successfully (demo mode)"
      } else {
        $secPwd = ConvertTo-SecureString -String $password1 -AsPlainText -Force
        Set-ADAccountPassword -Identity $userName -NewPassword $secPwd -Reset -ErrorAction Stop
        if ($chkMustChange.Checked) {
          Set-UnifiedObject -ObjectType User -Object $userName -Properties @{ ChangePasswordAtLogon = $true }
        }
        Show-Modal "Success" "Password reset successfully"
      }
      [Terminal.Gui.Application]::RequestStop()
    } catch {
      $errMsg = $_.Exception.Message
      Show-Modal "Error" "Failed to reset password:`n$errMsg"
    }
  }).GetNewClosure()
  $dlg.AddButton($btnOK)
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
  $dlg.AddButton($btnCancel)
  [Terminal.Gui.Application]::Run($dlg)
}

function Toggle-UserAccount {
  param([string]$userName, [bool]$disable)

  $action = if ($disable) { "disable" } else { "enable" }
  $result = Show-Modal "Confirm" "Are you sure you want to $action account:`n$userName?" -YesNo
  if ($result -eq 0) {
    try {
      if ($Script:DemoMode) {
        $user = $Script:Users | Where-Object { $_.Name -eq $userName } | Select-Object -First 1
        if ($user) {
          $user.Disabled = $disable
          Debug-Log (" Account $userName $action`d (demo mode)") -Type "Insight"
        }
      } else {
        if ($disable) {
          Disable-ADAccount -Identity $userName -ErrorAction Stop
        } else {
          Enable-ADAccount -Identity $userName -ErrorAction Stop
        }
      }
      Show-Modal "Success" "Account $action`d successfully"
      Refresh-Data -Domain $Script:CurrentDomain -RebuildTree -ShowModal -ShowLoadingDialog
    } catch {
      $errMsg = $_.Exception.Message
      Show-Modal "Error" "Failed to $action account:`n$errMsg"
    }
  }
}

## DSA-TUI Batch Operations Module v1.0 - Select multiple objects and perform bulk actions
## ----------{ Global Selection State }----------
$Script:SelectedObjects = @()
$Script:SelectionMode   = $false

## ----------{ Toggle Selection Mode }----------
function Toggle-SelectionMode {
  $Script:SelectionMode = -not $Script:SelectionMode

  if ($Script:SelectionMode) {
    Debug-Log (" Selection mode ENABLED") -Type "Insight"
    Show-Modal "Selection Mode" "Selection mode enabled!`n`nClick objects to select/deselect them.`nPress Ctrl+A to select all.`nPress Ctrl+D to deselect all."
  } else {
    Debug-Log (" Selection mode DISABLED") -Type "Insight"
    $Script:SelectedObjects = @()
    Build-Tree -domain $Script:CurrentDomain
    Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
  }
}

function Manage-Selection {
  <#
  .SYNOPSIS
  Select all or deselect all objects

  .PARAMETER Action
 'SelectAll' to select all objects, 'DeselectAll' to clear selection

  .EXAMPLE
  Manage-Selection -Action 'SelectAll'

  .EXAMPLE
  Manage-Selection -Action 'DeselectAll'
  #>

  param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('SelectAll', 'DeselectAll')]
    [string]$Action
  )

  if ($Action -eq 'SelectAll') {
    ## Check if selection mode is enabled
    if (-not $Script:SelectionMode) {
      Show-Modal "Selection Mode" "Enable selection mode first (Ctrl+S)"
      return
    }

    ## Clear existing selection
    $Script:SelectedObjects = @()

    ## Get all users from tree
    foreach ($user in $Script:Users) {
      $statusIcon = if ($user.Disabled) { "⊗" } else { "○" }
      $displayName = "(U) $statusIcon $($user.Name)"
      $Script:SelectedObjects += $displayName
    }
    Debug-Log "Selected all users ($($Script:SelectedObjects.Count))" -Type "Insight"
    Show-SelectionPanel -Parent $win
    Show-Modal "Selected All" "Selected $($Script:SelectedObjects.Count) users"
  }
  if ($Action -eq 'DeselectAll') {
    $Script:SelectedObjects = @()
    Debug-Log "Deselected all objects" -Type "Insight"
    Show-SelectionPanel -Parent $win
  }
}

## ----------{ Bulk Disable/Enable }----------
function Invoke-BulkDisableEnable {
  param([bool]$disable)

  ## Guard: nothing selected
  if ($Script:SelectedObjects.Count -eq 0) {
    Show-Modal "No Selection" "No objects selected. Select objects first."
    return
  }
  $action = if ($disable) { "disable" } else { "enable" }
  ## Confirm
  $result = Show-Modal "Confirm Bulk Action" "Are you sure you want to $action $($Script:SelectedObjects.Count) user account(s)?" -YesNo
  if ($result -ne 0) { return }

  $successCount = 0
  $failCount    = 0
  $errors       = @()

  foreach ($objName in $Script:SelectedObjects) {
    $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
    try {
      if ($Script:DemoMode) {
        $user = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
        if ($user) {
          $user.Disabled = $disable
          $successCount++
          Debug-Log (" $action`d $cleanName (demo mode)") -Type "Insight"
        }
      }
      else {
        if ($disable) { Disable-ADAccount -Identity $cleanName -ErrorAction Stop }
        else { Enable-ADAccount -Identity $cleanName -ErrorAction Stop }
        $successCount++
        Debug-Log (" $action`d $cleanName in AD") -Type "Insight"
      }
    }
    catch {
      $failCount++
      $errors += "$cleanName`: $($_.Exception.Message)"
      Debug-Log (" Failed to $action $cleanName`: $_") -Type "Insight"
    }
  }

  ## Results message
  $msg = "Successfully $action`d $successCount account(s)"
  if ($failCount -gt 0) {
    $msg += "`n`nFailed: $failCount"
    if ($errors.Count -gt 0 -and $errors.Count -le 5) { $msg += "`n`nErrors:`n" + ($errors -join "`n") }
  }
  Show-Modal "Bulk Action Complete" $msg
  ## Refresh UI
  Build-Tree -Domain $Script:CurrentDomain
  Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
  ## Clear selection + refresh panel
  $Script:SelectedObjects = @()
  $Script:SelectionMode   = $false
  Show-SelectionPanel -Parent $win
  ## Force redraw (prevents stale list artefacts)
  [Terminal.Gui.Application]::Refresh()
}

## ----------{ Bulk Add to Group }----------
function Invoke-BulkAddToGroup {
  if ($Script:SelectedObjects.Count -eq 0) {
    Show-Modal "No Selection" "No objects selected. Select objects first."
    return
  }

  $dlg       = [Terminal.Gui.Dialog]::new("Bulk Add to Group", 70, 18)
  $lblInfo   = [Terminal.Gui.Label]::new("Add $($Script:SelectedObjects.Count) user(s) to group:")
  $lblInfo.X = 2; $lblInfo.Y=1; $dlg.Add($lblInfo)

  ## Get list of groups
  $groupList = if ($Script:DemoMode) {
    $allGroups = @()
    foreach ($u in $Script:Users) { $allGroups += $u.Groups }
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
  $confirm = Show-Modal "Confirm Bulk Add" "Add $($Script:SelectedObjects.Count) user(s) to group:`n$targetGroup?" -YesNo

  if ($confirm -eq 0) {
    $successCount = 0
    $failCount    = 0

   foreach ($objName in $Script:SelectedObjects) {
     $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
      try {
        if ($Script:DemoMode) {
          $user = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
         if ($user -and $user.Groups -notcontains $targetGroup) {
            $user.Groups += $targetGroup
            $successCount++
            Debug-Log (" Added $cleanName to $targetGroup (demo mode)") -Type "Insight"
          }
        } else {
          Add-ADGroupMember -Identity $targetGroup -Members $cleanName -ErrorAction Stop
          $successCount++
          Debug-Log (" Added $cleanName to $targetGroup in AD") -Type "Insight"
        }
      } catch {
        $failCount++
        Debug-Log (" Failed to add $cleanName`: $_") -Type "Insight"
      }
    }
    Show-Modal "Bulk Add Complete" "Successfully added $successCount user(s)`nFailed: $failCount"
    ## Refresh tree
    Build-Tree -domain $Script:CurrentDomain
    Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
    [Terminal.Gui.Application]::RequestStop()
    }
  }).GetNewClosure()
  $dlg.AddButton($btnAdd)
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
  $dlg.AddButton($btnCancel)
  [Terminal.Gui.Application]::Run($dlg)
}

## ----------{ Bulk Add to Group }----------
function Invoke-BulkAddToGroup {
  if ($Script:SelectedObjects.Count -eq 0) {
    Show-Modal "No Selection" "No objects selected. Select objects first."
    return
  }

  $dlg       = [Terminal.Gui.Dialog]::new("Bulk Add to Group", 70, 18)
  $lblInfo   = [Terminal.Gui.Label]::new("Add $($Script:SelectedObjects.Count) user(s) to group:")
  $lblInfo.X = 2; $lblInfo.Y=1; $dlg.Add($lblInfo)

  ## Get list of groups
  $groupList = if ($Script:DemoMode) {
  $allGroups = @()
  foreach ($u in $Script:Users) { $allGroups += $u.Groups }
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
  $confirm     = Show-Modal "Confirm Bulk Add" "Add $($Script:SelectedObjects.Count) user(s) to group:`n$targetGroup?" -YesNo
  if ($confirm -eq 0) {
    $successCount = 0
    $failCount    = 0
    foreach ($objName in $Script:SelectedObjects) {
      $cleanName = $objName -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', ''
      try {
        if ($Script:DemoMode) {
          $user = $Script:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
          if ($user -and $user.Groups -notcontains $targetGroup) {
            $user.Groups += $targetGroup
            $successCount++
            Debug-Log (" Added $cleanName to $targetGroup (demo mode)") -Type "Insight"
          }
        } else {
          Add-ADGroupMember -Identity $targetGroup -Members $cleanName -ErrorAction Stop
          $successCount++
          Debug-Log (" Added $cleanName to $targetGroup in AD") -Type "Insight"
        }
      } catch {
        $failCount++
        Debug-Log (" Failed to add $cleanName`: $_") -Type "Insight"
      }
    }
    Show-Modal "Bulk Add Complete", "Successfully added $successCount user(s)`nFailed: $failCount"
    ## Refresh tree
    Build-Tree -domain $Script:CurrentDomain
    Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
    [Terminal.Gui.Application]::RequestStop()
  }
  }).GetNewClosure()
  $dlg.AddButton($btnAdd)
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
  $dlg.AddButton($btnCancel)
  [Terminal.Gui.Application]::Run($dlg)
}

## ----------{ Selection Panel (Create + Update) }----------
function Show-SelectionPanel {
  param(
    [Terminal.Gui.View]$Parent,
    [int]$PanelWidth  = 40,
    [int]$PanelHeight = 5
  )
  ## If panel already exists, just update it
  if ($Script:SelectionPanel -and $Script:SelectionPanel.Tag) {
    $panel       = $Script:SelectionPanel
    $lblCount    = $panel.Tag.CountLabel
    $lstSelected = $panel.Tag.ListView
  }
  else {
    ## ----------{ Create panel }----------
    $panel = [Terminal.Gui.FrameView]::new("Selected Objects")
    $panel.Width  = $PanelWidth
    $panel.Height = $PanelHeight
    $panel.X = [Terminal.Gui.Pos]::AnchorEnd($PanelWidth)

    ## Dynamic: Place below filter panel with 1 line gap
    $filterPanelRef = $Parent.Subviews | Where-Object { $_.GetType().Name -eq 'FrameView' -and $_.Text -eq 'Filters' } | Select-Object -First 1
    if ($filterPanelRef) {
      $panel.Y = [Terminal.Gui.Pos]::Bottom($filterPanelRef) + 1
    } else {
      $panel.Y = 26  # Fallback if filter panel not found
    }

    ## Count label
    $lblCount = [Terminal.Gui.Label]::new("0 objects selected")
    $lblCount.X = 1
    $lblCount.Y = 0
    $panel.Add($lblCount)
    ## ListView
    $lstSelected = [Terminal.Gui.ListView]::new(@())
    $lstSelected.X = 1
    $lstSelected.Y = 1
    $lstSelected.Width  = [Terminal.Gui.Dim]::Fill(2)
    $lstSelected.Height = [Terminal.Gui.Dim]::Fill(6)
    $panel.Add($lstSelected)
    ## Store references
    $panel | Add-Member -MemberType NoteProperty -Name Tag -Value @{
      CountLabel = $lblCount
      ListView   = $lstSelected
    } -Force
    ## ----------{ Batch action buttons }----------
    $yPos = [Terminal.Gui.Pos]::Bottom($lstSelected) + 1
    $btnBulkDisable = [Terminal.Gui.Button]::new("Disable")
    $btnBulkDisable.X = 2
    $btnBulkDisable.Y = $yPos
    $btnBulkDisable.add_Clicked({ Invoke-BulkDisableEnable -disable $true }).GetNewClosure()
    $panel.Add($btnBulkDisable)
    $btnBulkEnable = [Terminal.Gui.Button]::new("Enable")
    $btnBulkEnable.X = 15
    $btnBulkEnable.Y = $yPos
    $btnBulkEnable.add_Clicked({ Invoke-BulkDisableEnable -disable $false }).GetNewClosure()
    $panel.Add($btnBulkEnable)
    $btnBulkMove = [Terminal.Gui.Button]::new("Move")
    $btnBulkMove.X = 27
    $btnBulkMove.Y = $yPos
    $btnBulkMove.add_Clicked({ Invoke-BulkMove }).GetNewClosure()
    $panel.Add($btnBulkMove)
    ## Attach once
    $Parent.Add($panel)
    $Script:SelectionPanel = $panel
  }
  ## ----------{ Update contents }----------
  $count         = $Script:SelectedObjects.Count
  $lblCount.Text = "$count object(s) selected"
  $displayNames  = $Script:SelectedObjects | ForEach-Object { $_ -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', '' }
  $lstSelected.SetSource($displayNames)
  $panel.SetNeedsDisplay()
}

## ----------{ Add / Remove group members aka edit }----------
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

  Debug-Log "Edit group membership for user $($User.Name)" -Type "Insight"
  ## Get current user groups
  $currentGroups = @()
  if ($User.Groups) {
    $currentGroups = $User.Groups
  } elseif ($User.MemberOf) {
    $currentGroups = $User.MemberOf | ForEach-Object { if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ } }
  }

  ## Get all available groups
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

  ## Create dialog
  $dlg = [Terminal.Gui.Dialog]::new("Edit Group Membership - $($User.Name)", 90, 35)
  $lblInfo = [Terminal.Gui.Label]::new("Check/uncheck groups and click Apply to save changes:")
  $lblInfo.X = 2
  $lblInfo.Y = 1
  $dlg.Add($lblInfo)

  ## Create scrollable view for checkboxes
  $scrollView = [Terminal.Gui.ScrollView]::new()
  $scrollView.X = 2
  $scrollView.Y = 3
  $scrollView.Width = [Terminal.Gui.Dim]::Fill(2)
  $scrollView.Height = [Terminal.Gui.Dim]::Fill(5)
  $scrollView.ShowVerticalScrollIndicator = $true

  ## Track checkboxes and their corresponding groups
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

    ## Store reference
    $checkboxes[$groupName] = $chk
    $scrollView.Add($chk)
    $y++
  }

  ## Set content size for scrolling
  $scrollView.ContentSize = [Terminal.Gui.Size]::new(80, $y)
  $dlg.Add($scrollView)

  ## Buttons
  $btnApply = [Terminal.Gui.Button]::new("Apply Changes")
  $btnApply.X = 2
  $btnApply.Y = [Terminal.Gui.Pos]::Bottom($scrollView) + 1
  $btnApply.add_Clicked({
    ## Determine what changed
    $toAdd = @()
    $toRemove = @()

    foreach ($groupName in $checkboxes.Keys) {
      $isCurrentlyMember = $currentGroups -contains $groupName
      $isChecked = $checkboxes[$groupName].Checked

      if ($isChecked -and -not $isCurrentlyMember) {
        ## Need to add
        $toAdd += $groupName
      } elseif (-not $isChecked -and $isCurrentlyMember) {
        ## Need to remove
        $toRemove += $groupName
      }
    }

    if ($toAdd.Count -eq 0 -and $toRemove.Count -eq 0) {
      Show-Modal "No Changes" "No group membership changes were made"
      return
    }

    ## Show confirmation
    $changeMsg = ""
    if ($toAdd.Count -gt 0) { $changeMsg += "Add to $($toAdd.Count) group(s):`n  " + ($toAdd -join "`n  ") + "`n`n" }
    if ($toRemove.Count -gt 0) { $changeMsg += "Remove from $($toRemove.Count) group(s):`n  " + ($toRemove -join "`n  ") }
    $confirm = Show-Modal "Confirm Changes" "Apply these changes for $($User.Name)?`n`n$changeMsg" -YesNo
    ## Not "Yes"
    if ($confirm -ne 0) { return }

    ## Apply changes
    $addedCount = 0
    $removedCount = 0
    $errors = @()

    ## Add to groups
    foreach ($groupName in $toAdd) {
      try {
        if ($Script:DemoMode) {
         if (-not $User.Groups) { $User.Groups = @() }
           $User.Groups += $groupName
           Debug-Log "Added $($User.Name) to group $groupName (demo)" -Type "Success"
           $addedCount++
         } else {
           Add-ADGroupMember -Identity $groupName -Members $User.SamAccountName -ErrorAction Stop
           Debug-Log "Added $($User.SamAccountName) to group $groupName" -Type "Success"
           $addedCount++
         }
      } catch {
        $errors += "Failed to add to ${$groupName}: $($_.Exception.Message)"
        Debug-Log "Failed to add: $($_.Exception.Message)" -Type "Problem"
      }
    }

    ## Remove from groups
    foreach ($groupName in $toRemove) {
      try {
        if ($Script:DemoMode) {
          $User.Groups = $User.Groups | Where-Object { $_ -ne $groupName }
          Debug-Log "Removed $($User.Name) from group $groupName (demo)" -Type "Success"
          $removedCount++
        } else {
          Remove-ADGroupMember -Identity $groupName -Members $User.SamAccountName -Confirm:$false -ErrorAction Stop
          Debug-Log "Removed $($User.SamAccountName) from group $groupName" -Type "Success"
          $removedCount++
        }
      } catch {
        $errors += "Failed to remove from ${$groupName}: $($_.Exception.Message)"
        Debug-Log "Failed to remove: $($_.Exception.Message)" -Type "Problem"
      }
    }

    ## Show results
    $resultMsg = "Changes applied:"
    if ($addedCount -gt 0) { $resultMsg += "`nAdded to $addedCount group(s)" }
    if ($removedCount -gt 0) { $resultMsg += "`nRemoved from $removedCount group(s)" }
    if ($errors.Count -gt 0) {
      $resultMsg += "`n`nErrors:`n" + ($errors -join "`n")
      Show-Modal "Partial Success" $resultMsg
    } else {
      Show-Modal "Success" $resultMsg
    }

    ## Call OnUpdate callback if provided
    if ($OnUpdate) { & $OnUpdate }

    ## Close the dialog
    [Terminal.Gui.Application]::RequestStop()
  }).GetNewClosure()
  $dlg.Add($btnApply)

  ## Cancel button
  $btnCancel   = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.X = [Terminal.Gui.Pos]::Right($btnApply) + 2
  $btnCancel.Y = [Terminal.Gui.Pos]::Bottom($scrollView) + 1
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
  $dlg.Add($btnCancel)

  ## Run the dialog
  [Terminal.Gui.Application]::Run($dlg)
}

function Invoke-ForceDCReplication {
  param(
    [Parameter(Mandatory)]
    [string]$DcName
  )

  $confirm = Show-Modal -Title "Force Replication" -Msg "This will force Active Directory replication on:`n$DcName`n`nProceed?" -YesNo
  ## 0 = Yes, 1 = No
  if ($confirm -ne 0) { return }
  try {
    Start-Process -FilePath repadmin.exe -ArgumentList "/syncall","/AdeP",$DcName -NoNewWindow -Wait
    Show-Modal "Replication" "Forced replication completed for $DcName"
  }
  catch {
    Show-Modal "Replication Error" $_.Exception.Message
  }
}

## ----------{ Main Get-AD Health function }---------
function Show-ADHealthDialog {
  <#
  .SYNOPSIS
  Comprehensive Active Directory health check with tabbed interface for DSA-TUI

  .DESCRIPTION
  Performs multiple health checks on Active Directory environment with tabbed UI:
  - System Info
  - Domain Controller Status
  - Replication Health
  - DNS Records
  - SYSVOL/NETLOGON Shares
  - FSMO Roles
  - Trust Relationships
  - DFS Status
  - Group Policy Objects

  .PARAMETER Domain
  The domain to check. Defaults to $Script:CurrentDomain

  .EXAMPLE
  Show-ADHealthDialog

  .EXAMPLE
  Show-ADHealthDialog -Domain "example.com"
  #>

  [CmdletBinding()]
  param(
    [Parameter(Position=0)]
    [string]$Domain
  )

  Debug-Log "Show-ADHealthDialog called" -Type "Insight"

  ## ----------{ OS & Tool detection }---------

  ## Detect OS
  $osInfo = @{
    IsWindows = $false
    IsLinux   = $false
    IsMacOS   = $false
    OSName    = "Unknown"
  }

  if ($PSVersionTable.PSVersion.Major -ge 6) {
    ## PowerShell Core 6+
    $osInfo.IsWindows = $IsWindows
    $osInfo.IsLinux = $IsLinux
    $osInfo.IsMacOS = $IsMacOS
  } else {
    ## Windows PowerShell 5.1
    $osInfo.IsWindows = $true
  }

  if ($osInfo.IsWindows) { $osInfo.OSName = "Windows $($PSVersionTable.PSVersion)" }
  elseif ($osInfo.IsLinux) { $osInfo.OSName = "Linux" }
  elseif ($osInfo.IsMacOS) { $osInfo.OSName = "macOS" }
  Debug-Log "Detected OS: $($osInfo.OSName)" -Type "Insight"

  ## Check tool availability
  $tools = Test-ToolsAvailability
  Debug-Log "Tool availability checked" -Type "Insight"

  ## Check AD module
  $hasADModule = $null -ne ($Script:HasActiveDirectory)
  Debug-Log "ActiveDirectory module available: $hasADModule" -Type "Insight"

  if (-not $hasADModule -and -not $Script:DemoMode) {
    Show-Modal "AD Module Missing" "ActiveDirectory PowerShell module is not available.`n`nOS: $($osInfo.OSName)`n`nInstall RSAT (Windows) or realmd/sssd (Linux/macOS) to use this feature."
    return
  }

  ## Determine domain to check
  if (-not $Domain) {
    if ($Script:CurrentDomain) {
      $Domain = $Script:CurrentDomain
    } else {
      if (-not $Script:DemoMode) {
        try {
          $Domain = (Get-ADDomain -Current LocalComputer).DNSRoot
        } catch {
          Show-Modal "Error" "Could not determine domain. Please specify domain in demo mode or ensure AD module is available."
          return
        }
      } else {
        $Domain = "example.com"
      }
    }
  }

  Debug-Log "Checking AD Health for domain: $Domain" -Type "Insight"

  ## Store tools and OS info in script scope for helper functions
  $Script:ADHealthTools  = $tools
  $Script:ADHealthOS     = $osInfo
  $Script:ADHealthDomain = $Domain


  ## Function captures:
  $getFSMOFunc      = ${function:Get-FSMOStatusText}
  $getTrustTestFunc = ${function:Get-TrustTestingTabContent}
  $getStatsFunc     = ${function:Get-DomainStatisticsText}

  ## Create main dialog
  $dialog = [Terminal.Gui.Dialog]::new("AD Health Check - $Domain", 140, 35)

  ## Create TabView
  $tabView = [Terminal.Gui.TabView]::new()
  $tabView.X = 0
  $tabView.Y = 0
  $tabView.Width  = [Terminal.Gui.Dim]::Fill()
  $tabView.Height = [Terminal.Gui.Dim]::Fill(2)

  ## Tab definitions for AD health check
  $tabs = @(
    @{ Name = "System Info"         ; Generator = { Get-SystemInfoText } }
    @{ Name = "Domain Controllers"  ; Generator = { Get-DCStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "Replication"         ; Generator = { Get-ReplicationStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "DNS Records"         ; Generator = { Get-DNSStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "SYSVOL/NETLOGON"     ; Generator = { Get-SysvolStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "FSMO Roles"          ; Generator = { Get-FSMOStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "Trust Relationships" ; Generator = { Get-TrustStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "Trust Testing"       ; Generator = { Get-TrustTestingTabContent -Domain $Script:ADHealthDomain } }
    @{ Name = "Statistics"          ; Generator = { Get-DomainStatisticsText -Domain $Script:ADHealthDomain } }
    @{ Name = "DFS Status"          ; Generator = { Get-DFSStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "Group Policy"        ; Generator = { Get-GPOStatusText -Domain $Script:ADHealthDomain } }
  )

  ## Create tabs
  foreach ($tabDef in $tabs) {
    $tab             = [Terminal.Gui.TabView+Tab]::new()
    $tab.Text        = $tabDef.Name
    $tab.View        = [Terminal.Gui.View]::new()
    $tab.View.Width  = [Terminal.Gui.Dim]::Fill()
    $tab.View.Height = [Terminal.Gui.Dim]::Fill()

    $text = & $tabDef.Generator
    $textView   = [Terminal.Gui.TextView]::new()
    $textView.X = 1
    $textView.Y = 0
    $textView.Width    = [Terminal.Gui.Dim]::Fill(1)
    $textView.Height   = [Terminal.Gui.Dim]::Fill()
    $textView.ReadOnly = $true
    $textView.Text     = if ($text) { $text } else { "" }
    $tab.View.Add($textView)
    $tabView.AddTab($tab, $false)
  }
  $dialog.Add($tabView)

  ## ----------{ CTRL+F Search }---------
  ## Add key handler for Ctrl+F
  $dialog.add_KeyPress({
    param($e)

    if (($e.KeyEvent.Key -band [Terminal.Gui.Key]::CtrlMask) -and
        (([int]$e.KeyEvent.Key -band 0xFF) -eq [byte][char]'f')) {
      $currentTab = $tabView.SelectedTab
      $textView   = $currentTab.View.Subviews[0]
      if ($textView -and $textView -is [Terminal.Gui.TextView]) {
        ## Simple search prompt
        $searchText = Show-InputDialog -Title "Search" -Label "Search for:" -Width 50
        if ($searchText) {
          $content = $textView.Text.ToString()
          $index   = $content.IndexOf($searchText, [StringComparison]::OrdinalIgnoreCase)
          if ($index -ge 0) {
            ## Calculate line number
            $lines = $content.Substring(0, $index) -split "`n"
            $lineNum = $lines.Count - 1
            ## Move cursor to that line
            try {
              $textView.CursorPosition = [Terminal.Gui.Point]::new(0, $lineNum)
              $textView.SetNeedsDisplay()
            } catch { }
            Show-Modal "Search" "Found at line $($lineNum + 1)"
          } else {
            Show-Modal "Search" "Text '$searchText' not found"
          }
        }
      }
      $e.Handled = $true
    }
  }.GetNewClosure())

  ## ----------{ Capture Variables for Closures }---------
  $capturedTabs = $tabs
  $capturedTabView = $tabView
  $capturedDomain = $Domain

  ## ----------{ Buttons }---------
  $btnRefresh = [Terminal.Gui.Button]::new("Refresh")
  $btnRefresh.X = 2
  $btnRefresh.Y = [Terminal.Gui.Pos]::AnchorEnd(1)
  $btnRefresh.add_Clicked({
    Debug-Log "Refreshing current tab..." -Type "Insight"
    $currentTab = $capturedTabView.SelectedTab
    $tabName = $currentTab.Text.ToString()
    ## Find matching tab definition
    $tabDef = $capturedTabs | Where-Object { $_.Name -eq $tabName } | Select-Object -First 1
    if ($tabDef) {
      ## Refresh tools check if System Info
      if ($tabName -eq "System Info") { $Script:ADHealthTools = Test-ToolsAvailability }
      $newText = & $tabDef.Generator
      if ($newText) {
        $currentTab.View.Subviews[0].Text = $newText
        $currentTab.View.Subviews[0].SetNeedsDisplay()
      }
    }
  }.GetNewClosure())
  $dialog.Add($btnRefresh)

  $btnExport = [Terminal.Gui.Button]::new("Export Report")
  $btnExport.X = 15
  $btnExport.Y = [Terminal.Gui.Pos]::AnchorEnd(1)
  $btnExport.add_Clicked({
    Debug-Log "Exporting AD health report..." -Type "Insight"

    ## Build report from all tabs
    $report = @"
AD HEALTH REPORT
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Domain: $capturedDomain
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

"@

    foreach ($tabDef in $capturedTabs) {
      $report += "`n$($tabDef.Name.ToUpper())`n"
      $report += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n"
      $report += & $tabDef.Generator
      $report += "`n"
    }

    $report += "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n"
    $report += "End of Report`n"

    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $defaultName = "ADHealth_$($capturedDomain)_$timestamp.txt"
    $savePath = Show-FileBrowserDialog -Mode 'Save' -StartDir "." -DefaultName $defaultName

    if ($savePath) {
      try {
        $report | Out-File -FilePath $savePath -Encoding UTF8
        Debug-Log "Report exported to $savePath" -Type "Success"
        Show-Modal "Export Complete" "AD Health Report saved to:`n`n$savePath"
      } catch {
        Debug-Log "Failed to export report: $($_.Exception.Message)" -Type "Problem"
        Show-Modal "Export Failed" "Failed to save report:`n`n$($_.Exception.Message)"
      }
    }
  }.GetNewClosure())
  $dialog.Add($btnExport)

  ## Live search typing:
  ## Search label
  $lblSearch = [Terminal.Gui.Label]::new("Search:")
  $lblSearch.X = 36
  $lblSearch.Y = [Terminal.Gui.Pos]::AnchorEnd(1)
  $dialog.Add($lblSearch)

  ## Search text field
  $txtSearch = [Terminal.Gui.TextField]::new("")
  $txtSearch.X = 44
  $txtSearch.Y = [Terminal.Gui.Pos]::AnchorEnd(1)
  $txtSearch.Width = 30
  $dialog.Add($txtSearch)

  ## Store original content for each tab (for filtering)
  $tabOriginalContent = @{}

  ## Search handler - filters current tab as user types
  $txtSearch.add_TextChanged({
    try {
      $searchTerm = $txtSearch.Text.ToString().Trim()
      $currentTab = $capturedTabView.SelectedTab
      $textView = $currentTab.View.Subviews[0]

      if (-not $textView -or $textView -isnot [Terminal.Gui.TextView]) { return }

      $tabName = $currentTab.Text.ToString()

      ## Store original content on first search
      if (-not $tabOriginalContent.ContainsKey($tabName)) {
        $tabOriginalContent[$tabName] = $textView.Text.ToString()
      }

      ## Get original content
      $originalContent = $tabOriginalContent[$tabName]

      if ([string]::IsNullOrWhiteSpace($searchTerm)) {
        ## No search term - show original content
        $textView.Text = $originalContent
        $textView.SetNeedsDisplay()
        return
      }

      ## Filter lines containing search term (case-insensitive)
      $lines = $originalContent -split "`n"
      $filteredLines = $lines | Where-Object {
        $_ -match [regex]::Escape($searchTerm)
      }

      if ($filteredLines.Count -eq 0) {
        $filteredContent = "(No matches found for '$searchTerm')"
      } else {
        $filteredContent = $filteredLines -join "`n"
      }

      $textView.Text = $filteredContent
      $textView.SetNeedsDisplay()
    } catch {
      Debug-Log "Search error: $($_.Exception.Message)" -Type "Problem"
    }
  }.GetNewClosure())

  ## Clear search when tab changes
  $capturedTabView.add_SelectedTabChanged({
    try {
      $txtSearch.Text = ""
      ## Restore original content if we had filtered
      $currentTab = $capturedTabView.SelectedTab
      $tabName = $currentTab.Text.ToString()
      if ($tabOriginalContent.ContainsKey($tabName)) {
        $textView = $currentTab.View.Subviews[0]
        if ($textView -and $textView -is [Terminal.Gui.TextView]) {
          $textView.Text = $tabOriginalContent[$tabName]
          $textView.SetNeedsDisplay()
        }
      }
    } catch { }
  }.GetNewClosure())

  ## Close button
  $btnClose   = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = [Terminal.Gui.Pos]::AnchorEnd(10)
  $btnClose.Y = [Terminal.Gui.Pos]::AnchorEnd(1)
  $btnClose.add_Clicked({
    Debug-Log "AD Health dialog closed" -Type "Insight"
    [Terminal.Gui.Application]::RequestStop()
  }.GetNewClosure())
  $dialog.Add($btnClose)

  ## Run dialog
  Debug-Log "Running AD Health dialog" -Type "Insight"
  [Terminal.Gui.Application]::Run($dialog)
}

function Check-DCReplication {
  [CmdletBinding()]

  param(
    [Parameter(Mandatory)]
    [string]$dcName
  )

  try {
    $partners = Get-ADReplicationPartnerMetadata -Target $dcName -Scope Server -ErrorAction Stop
    $text = ($partners | ForEach-Object {
      $status = if ($_.ConsecutiveReplicationFailures -gt 0) { "❌ FAIL ($($_.ConsecutiveReplicationFailures))"            }
                else { "✅ OK" }

@"
Partner   : $($_.Partner)
Status    : $status
Last Sync : $($_.LastReplicationSuccess)
Transport : $($_.InterSiteTransportProtocol)
─────────────────────────────────────
"@

    }) -join "`n"

    $btnForce = [Terminal.Gui.Button]::new("Force Replication")
    $btnClose = [Terminal.Gui.Button]::new("Close")

    $btnForce.add_Clicked({ Invoke-ForceDCReplication -DcName $dcName }).GetNewClosure()
    $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
    $dialog = [Terminal.Gui.Dialog]::new("Replication – $dcName", 90, 30, $btnForce, $btnClose)

    $textView = [Terminal.Gui.TextView]::new(1,1,86,22)
    $textView.Text = $text
    $textView.ReadOnly = $true
    $dialog.Add($textView)
    [Terminal.Gui.Application]::Run($dialog)
  }
  catch {
    Show-Modal "Replication Error" $_.Exception.Message
  }
}

function Show-OUPropertiesDialog {
  param($ou)

  Debug-Log "Showing OU properties dialog" -Type "Insight"
  if (-not $ou) {
    Debug-Log "OU object is null" -Type "Warning"
    return
  }

  $ouName = $ou.Name ?? $ou.Text ?? "Unknown"
  Debug-Log "OU name resolved to: $ouName" -Type "Insight"

  ## If not found, create a basic one
  if (-not $ou.Name) {
    $ou = @{
      Name = $ouName
      Path = ""
      Description = ""
    }
  }

  ## ----------{ General Tab }---------
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $ou, $state)

      $y = 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Name:" -FieldName 'txtName' -State $state -Value ($ou.Name ?? "") -Width 50
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Description:" -FieldName 'txtDesc' -State $state -Value ($ou.Description ?? "") -Width 50
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Path:" -FieldName 'txtPath' -State $state -Value ($ou.Path ?? "") -Width 50 -ReadOnly $true
      $y += 1
      $objectCount = ($Script:Users | Where-Object { $_.OU -contains $ou.Name }).Count
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Contains:" -FieldName 'lblCount' -State $state -Value "$objectCount objects" -IsTextField $false
    }
  }

  ## ----------{ Trust Testing Tab }---------
  $trustTestTab = @{
    Name = "Trust Tests"
    Builder = {
      param($view, $data, $state)
      ## Capture functions
      $debugLogFunc = ${function:Debug-Log}
      $showModalFunc = ${function:Show-Modal}
      $y = 1
      ## Header
      $lblHeader = [Terminal.Gui.Label]::new("Test trust relationships with partner domains")
      $lblHeader.X = 2
      $lblHeader.Y = $y
      $view.Add($lblHeader)
      $y += 2
      ## Get trust relationships from data
      $trusts = if ($data -and $data.ContainsKey('Trusts')) {
        $data['Trusts']
      } else {
        @()
      }
      if ($trusts.Count -eq 0) {
        $lblNoTrusts = [Terminal.Gui.Label]::new("No trust relationships configured")
        $lblNoTrusts.X = 2
        $lblNoTrusts.Y = $y
        $view.Add($lblNoTrusts)
        return
      }
      ## Trust list
      $lblTrusts = [Terminal.Gui.Label]::new("Available Trusts:")
      $lblTrusts.X = 2
      $lblTrusts.Y = $y
      $view.Add($lblTrusts)
      $y += 1
      $lstTrusts = [Terminal.Gui.ListView]::new()
      $lstTrusts.X = 2
      $lstTrusts.Y = $y
      $lstTrusts.Width = [Terminal.Gui.Dim]::Fill(2)
      $lstTrusts.Height = 10
      $trustNames = $trusts | ForEach-Object {
        $partner = if ($_.Partner) { $_.Partner } else { $_.Name }
        $direction = if ($_.Direction) { $_.Direction } else { "Unknown" }
        "$partner ($direction)"
      }
      $lstTrusts.SetSource([string[]]$trustNames)
      $view.Add($lstTrusts)
      $y += 11
      ## Results area
      $lblResults = [Terminal.Gui.Label]::new("Test Results:")
      $lblResults.X = 2
      $lblResults.Y = $y
      $view.Add($lblResults)
      $y += 1
      $txtResults = [Terminal.Gui.TextView]::new()
      $txtResults.X = 2
      $txtResults.Y = $y
      $txtResults.Width = [Terminal.Gui.Dim]::Fill(2)
      $txtResults.Height = [Terminal.Gui.Dim]::Fill(4)
      $txtResults.ReadOnly = $true
      $view.Add($txtResults)
      ## Test button
      $btnTest = [Terminal.Gui.Button]::new("Test Selected Trust")
      $btnTest.X = 2
      $btnTest.Y = [Terminal.Gui.Pos]::AnchorEnd(2)
      $btnTest.add_Clicked({
        if ($lstTrusts.SelectedItem -lt 0) {
          & $showModalFunc "No Selection" "Please select a trust to test"
          return
        }
        $selectedTrust = $trusts[$lstTrusts.SelectedItem]
        $partner = if ($selectedTrust.Partner) { $selectedTrust.Partner } else { $selectedTrust.Name }
        & $debugLogFunc "Testing trust connection to $partner" -Type "Insight"
        ## Build test results
        $results = @()
        $results += "Testing trust relationship with: $partner"
        $results += "Direction: $($selectedTrust.Direction)"
        $results += "Type: $($selectedTrust.Type)"
        $results += ""
        $results += "═══════════════════════════════════════════"
        $results += ""
        try {
          if ($Script:DemoMode -or (-not (Get-Command nltest -ErrorAction SilentlyContinue))) {
            ## Demo mode or nltest not available
            $results += "⚠ Demo Mode or nltest not available"
            $results += ""
            $results += "To test trust in production, use:"
            $results += ""
            $results += "  nltest /server:$env:COMPUTERNAME /sc_query:$partner"
            $results += ""
            $results += "Or use PowerShell:"
            $results += "  Test-ComputerSecureChannel -Credential (Get-Credential) -Server $partner"
            $results += ""
            $results += "Or use netdom:"
            $results += "  netdom trust $($Script:CurrentDomain) /domain:$partner /verify"
            $results += ""
            $results += "═══════════════════════════════════════════"
            $results += ""
            $results += "Simulated Result: ✓ Trust appears healthy"
            $results += "  - Secure channel: OK"
            $results += "  - Authentication: OK"
            $results += "  - Trust status: Active"
          } else {
            ## Production Mode - actually test
            $results += "Running: nltest /sc_query:$partner"
            $results += ""
            $nlTestOutput = nltest /sc_query:$partner 2>&1
            $results += $nlTestOutput -join "`n"
            $results += ""
            $results += "═══════════════════════════════════════════"
            $results += ""

            if ($LASTEXITCODE -eq 0) {
              $results += "✓ Trust test successful"
            } else {
              $results += "✗ Trust test failed (exit code: $LASTEXITCODE)"
            }
          }
        } catch {
          $results += "✗ Error testing trust:"
          $results += $_.Exception.Message
        }
        $txtResults.Text = $results -join "`n"
        & $debugLogFunc "Trust test completed for $partner" -Type "Success"
      }.GetNewClosure())
      $view.Add($btnTest)
    }
  }

  ## ----------{ Statistics Tab }---------
  $statisticsTab = @{
    Name = "Statistics"
    Builder = {
      param($view, $ou, $state)

      ## Calculate statistics
      $ouName = $ou.Name
      $usersInOU     = $Script:Users | Where-Object { $_.OU -contains $ouName }
      $enabledUsers  = ($usersInOU   | Where-Object { $_.Enabled -eq $true }).Count
      $disabledUsers = ($usersInOU   | Where-Object { $_.Disabled -eq $true }).Count
      $lockedUsers   = ($usersInOU   | Where-Object { $_.LockedOut -eq $true }).Count

      $groupsInOU = if ($Script:Groups) {
        $Script:Groups | Where-Object { $_.OU -contains $ouName }
      } else { @() }

      $computersInOU = if ($Script:Computers) {
        $Script:Computers | Where-Object { $_.OU -contains $ouName }
      } else { @() }

      $enabledComputers  = ($computersInOU | Where-Object { $_.Enabled -eq $true }).Count
      $disabledComputers = ($computersInOU | Where-Object { $_.Disabled -eq $true }).Count

      $nestedOUs = if ($Script:rawOUs) {
        $Script:rawOUs | Where-Object { $_.Path -match [regex]::Escape($ouName) -and $_.Name -ne $ouName }
      } else { @() }

      $totalObjects = $usersInOU.Count + $groupsInOU.Count + $computersInOU.Count
      ## Display statistics
      $y = 1
      $lblHeader = [Terminal.Gui.Label]::new("OU: $ouName")
      $lblHeader.X = 2; $lblHeader.Y = $y; $view.Add($lblHeader); $y += 2

      ## Users section
      $lblUsers = [Terminal.Gui.Label]::new("═══ USERS ═══")
      $lblUsers.X = 2; $lblUsers.Y = $y; $view.Add($lblUsers); $y += 1
      $lblUserTotal = [Terminal.Gui.Label]::new("Total Users:       $($usersInOU.Count)")
      $lblUserTotal.X = 4; $lblUserTotal.Y = $y; $view.Add($lblUserTotal); $y += 1
      $lblUserEnabled = [Terminal.Gui.Label]::new("  Enabled:         $enabledUsers")
      $lblUserEnabled.X = 4; $lblUserEnabled.Y = $y; $view.Add($lblUserEnabled); $y += 1
      $lblUserDisabled = [Terminal.Gui.Label]::new("  Disabled:        $disabledUsers")
      $lblUserDisabled.X = 4; $lblUserDisabled.Y = $y; $view.Add($lblUserDisabled); $y += 1
      $lblUserLocked = [Terminal.Gui.Label]::new("  Locked Out:      $lockedUsers")
      $lblUserLocked.X = 4; $lblUserLocked.Y = $y; $view.Add($lblUserLocked); $y += 2

      ## Computers section
      $lblComputers = [Terminal.Gui.Label]::new("═══ COMPUTERS ═══")
      $lblComputers.X = 2; $lblComputers.Y = $y; $view.Add($lblComputers); $y += 1
      $lblCompTotal = [Terminal.Gui.Label]::new("Total Computers:   $($computersInOU.Count)")
      $lblCompTotal.X = 4; $lblCompTotal.Y = $y; $view.Add($lblCompTotal); $y += 1
      $lblCompEnabled = [Terminal.Gui.Label]::new("  Enabled:         $enabledComputers")
      $lblCompEnabled.X = 4; $lblCompEnabled.Y = $y; $view.Add($lblCompEnabled); $y += 1
      $lblCompDisabled = [Terminal.Gui.Label]::new("  Disabled:        $disabledComputers")
      $lblCompDisabled.X = 4; $lblCompDisabled.Y = $y; $view.Add($lblCompDisabled); $y += 2

      ## Groups section
      $lblGroups = [Terminal.Gui.Label]::new("═══ GROUPS ═══")
      $lblGroups.X = 2; $lblGroups.Y = $y; $view.Add($lblGroups); $y += 1
      $lblGroupTotal = [Terminal.Gui.Label]::new("Total Groups:      $($groupsInOU.Count)")
      $lblGroupTotal.X = 4; $lblGroupTotal.Y = $y; $view.Add($lblGroupTotal); $y += 2

      ## Structure section
      $lblStructure = [Terminal.Gui.Label]::new("═══ STRUCTURE ═══")
      $lblStructure.X = 2; $lblStructure.Y = $y; $view.Add($lblStructure); $y += 1
      $lblNestedOUs = [Terminal.Gui.Label]::new("Nested OUs:        $($nestedOUs.Count)")
      $lblNestedOUs.X = 4; $lblNestedOUs.Y = $y; $view.Add($lblNestedOUs); $y += 2

      ## Total
      $lblTotal = [Terminal.Gui.Label]::new("═══════════════════")
      $lblTotal.X = 2; $lblTotal.Y = $y; $view.Add($lblTotal); $y += 1

      $lblTotalObjects = [Terminal.Gui.Label]::new("TOTAL OBJECTS:     $totalObjects")
      $lblTotalObjects.X = 2; $lblTotalObjects.Y = $y; $view.Add($lblTotalObjects); $y += 2

      ## Export button
      $btnExport = [Terminal.Gui.Button]::new("Export Statistics")
      $btnExport.X = 2; $btnExport.Y = $y
      $btnExport.add_Clicked({
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filename = "ou_stats_${ouName}_$timestamp.csv"
        try {
          $stats = [PSCustomObject]@{
            OU                = $ouName
            Timestamp         = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            TotalUsers        = $usersInOU.Count
            EnabledUsers      = $enabledUsers
            DisabledUsers     = $disabledUsers
            LockedUsers       = $lockedUsers
            TotalComputers    = $computersInOU.Count
            EnabledComputers  = $enabledComputers
            DisabledComputers = $disabledComputers
            TotalGroups       = $groupsInOU.Count
            NestedOUs         = $nestedOUs.Count
            TotalObjects      = $totalObjects
          }
          $stats | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
          Show-Modal "Export Complete" "Statistics exported to:`n`n$filename"
          Debug-Log "Exported OU statistics to $filename" -Type "Success"
        } catch {
          Show-Modal "Export Failed" "Failed to export statistics:`n`n$($_.Exception.Message)"
          Debug-Log "Failed to export OU statistics: $($_.Exception.Message)" -Type "Problem"
        }
      }.GetNewClosure())
      $view.Add($btnExport)
    }
  }
  ## ----------{ Apply Logic }---------
  $applyLogic = {
    param($ou, $state)
    try {
      $changesMade = $false

      if ($state.txtDesc) {
        $newDesc = $state.txtDesc.Text.ToString().Trim()
        if ($newDesc -ne ($ou.Description ?? "")) {
          if ($Script:DemoMode) {
            $ou.Description = $newDesc
          } else {
            Set-ADOrganizationalUnit -Identity $ou.DistinguishedName -Description $newDesc
          }
          $changesMade = $true
        }
      }
      if ($changesMade) {
        Show-Modal "Success" "Changes applied successfully"
      } else {
        Show-Modal "Info" "No changes to apply"
      }
    } catch {
      Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)"
    }
  }
  ## ----------{ Create Dialog }---------
  ## Note: OUs don't get a search tab
  $tabs = @($generalTab, $statisticsTab)
  Debug-Log "Show-OUPropertiesDialog running" -Type "Insight"
  New-PropertiesDialog -Title "OU Properties - $ouName" -Width 80 -Height 28 -Tabs $tabs -Data $ou -OnApply $applyLogic -IncludeSearchTab $false
  Debug-Log "Show-OUPropertiesDialog completed" -Type "Insight"
}

## Capture Set-ObjectCheckboxes for use in closures (Needed here, outside of helper functions)
$Script:SetObjectCheckboxes_Func = ${function:Set-ObjectCheckboxes}

## Don't let users do stupid stuff
if ($DemoMode -and $PSBoundParameters.ContainsKey('Domain')) {
  Debug-Log "Invalid startup: -Domain cannot be used with -DemoMode" -Type "Problem"
  return
}

## ----------{ Program Launch starts Here }----------
## Echo basic info for debugging
Debug-Log "DemoMode: $DemoMode" -Type "Insight"
Debug-Log "Logging: $Logging" -Type "Insight"
Debug-Log "LogFile: $LogFile" -Type "Insight"

## Initialise logging if requested
if ($Logging -or $LogFile) {
  Debug-Log "Logging condition TRUE" -Type "Insight"
  if (-not $LogFile) {
    $LogFile = Join-Path $PSScriptRoot "dsa_tui_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Debug-Log "Auto-generated LogFile: $LogFile" -Type "Tracing"
  }

  if (-not [System.IO.Path]::IsPathRooted($LogFile)) { $LogFile = Join-Path (Get-Location).Path $LogFile }
  $Script:Logging    = $true
  $Script:LogFile    = $LogFile
  $Script:StatusItem = [PSCustomObject]@{ Title = "Initializing..." }

  Debug-Log "Attempting to create log at: $LogFile" -Type "Tracing"
  try {
    $Script:LogStream           = [System.IO.StreamWriter]::new($LogFile, $false)
    $Script:LogStream.AutoFlush = $true
    Debug-Log "SUCCESS: Log file created at $LogFile" -Type "Tracing"
    $Script:LogStream.WriteLine("========== DSA-TUI Log Started $(Get-Date) ==========")
    $Script:LogStream.WriteLine("DemoMode: $DemoMode")
    $Script:LogStream.WriteLine("Theme: $Theme")
    $Script:LogStream.Flush()
  } catch {
    Debug-Log "FAILED to create log: $_" -Type "Problem"
    $Script:Logging = $false
  }
} else {
  Debug-Log "Logging condition FALSE - no logging enabled" -Type "Warning"
}

## Set globals
$Script:DemoMode       = $DemoMode
$Script:ThemeMode      = $Theme
$Script:DataFileLoaded = $false
$Script:DataFilePath   = $null

## ----------{ Step 2: Module Checks & Terminal.Gui }----------
Debug-Log "Performing pre-flight module checks..." -Type "Insight"

## Check all modules ONCE at startup
## ----------{ Preflight Checks }----------

## Core modules
$Script:hasConsoleTools    = Test-Requirement -Type Module -Name 'Microsoft.PowerShell.ConsoleGuiTools' -InstallMsg 'Install-Module -Name Microsoft.PowerShell.ConsoleGuiTools -RequiredVersion 0.7.2'
$Script:HasPSWriteColor    = Test-Requirement -Type Module -Name "PSWriteColor" -InstallMsg 'Install-Module -Name PSWriteColor' -Optional
$Script:HasTerminalIcons   = Test-Requirement -Type Module -Name "Terminal-Icons" -InstallMsg 'Install-Module -Name Terminal-Icons' -Optional
$Script:hasNerdFonts       = Test-Requirement -Type Module -Name 'NerdFonts' -InstallMsg 'Install-Module -Name NerdFonts' -Optional
$Script:hasNerdFonts       = Test-Requirement -Type Module -Name 'DNSClient' -InstallMsg 'Install-Module -Name DNSClient' -Optional

## Optional modules
$Script:HasActiveDirectory = Test-Requirement -Type Module -Name "ActiveDirectory" -InstallMsg 'Install-Module -Name AzureAD' -Optional
$Script:HasPSWindowsUpdate = Test-Requirement -Type Module -Name "PSWindowsUpdate" -InstallMsg 'Install-Module -Name PSWindowsUpdate -RequiredVersion 2.2.0.3' -Optional

## Windows capabilities (RSAT tools)
$Script:HasDNSServer       = Test-Requirement -Type WindowsCapability -Name "Rsat.Dns.Tools~~~~0.0.1.0" -InstallMsg 'Add-WindowsCapability -Online -Name Rsat.Dns.Tools~~~~0.0.1.0' -Optional

## Windows Features (Server only)
$Script:HasGroupPolicy     = Test-Requirement -Type WindowsFeature -Name "GPMC" -InstallMsg 'Install-WindowsFeature GPMC' -Optional
$Script:HasDFSR            = Test-Requirement -Type WindowsFeature -Name "RSAT-DFS-Mgmt-Con" -InstallMsg "Install-WindowsFeature 'RSAT-DFS-Mgmt-Con'" -Optional

## Optional insight/log
Debug-Log "To install all RSAT features (may be overkill): 'Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability -Online'" -Type "Insight"

## Initialise icon set based on available modules
Initialise-Icons

if ($Script:LaunchReady) {
  Debug-Log "Only checking of script prerequisites requested. Now exiting..." -Type "Insight"
  return
}

## Lack of AD module defaults to demo mode
if (-not $Script:HasActiveDirectory) {
  Debug-Log "ActiveDirectory module missing. Falling back to DEMO mode..." -Type "Warning"
  $Script:DemoMode = $true
}
Debug-Log "Module availability check complete" -Type "Insight"
$Script:UseIcons = $false
if ($Script:HasTerminalIcons) { try { Write-Host '' -NoNewline; $Script:UseIcons = $true } catch {} }

## ----------{Step 3: Initialise Terminal.Gui UI }----------
Initialise-DirectoryEmoji
$windowTitle = "$($Script:ProjectName) $($Script:DirectoryEmoji) Active Directory $BuildVersion Codename: $($Script:FruitName)"

$uiComponents = Initialise-UIFramework -Theme $Theme -Title $windowTitle
$top = $uiComponents.Top
$win = $uiComponents.Window
$Script:themeData = $uiComponents.Theme

## TODO: This is where the statusbar and menu needs to move to

## ----------{Step 4: Create Status & Menu Bars }----------
Debug-Log "Creating status bar..." -Type "Insight"
$statusBar = Set-StatusBar -Initialise -ThemeData $Script:themeData
$top.Add($statusBar)

Debug-Log "Creating main menu..." -Type "Insight"
$menu = Build-MainMenu
$top.Add($menu)

Debug-Log "UI Framework ready - window visible to user" -Type "Success"
Get-Theme -Dump $Script:themeData

## ----------{ Step 7: Build Remaining UI Components }----------
Debug-Log "Creating filter panel..." -Type "Insight"
$filterPanel = Create-FilterPanel -Parent $win
$win.Add($filterPanel)

Debug-Log "Showing selection panel..." -Type "Insight"
Show-SelectionPanel -Parent $win
Debug-Log "Creating Info panel..." -Type "Insight"
Show-InfoPanel -Parent $win

## ----------{Step 5: Forest/Domain Initialization }----------
Debug-Log "Initializing forest/domain globals..." -Type "Insight"

## Tab & layout placeholders
$Script:LayoutInProgress = $false
$Script:TabRows = @()
$Script:AllTabs = @()
$Script:ActiveTab = $null
$Script:TabRowHeight = 1

if ($Script:DemoMode) {
  Debug-Log "DemoMode enabled..." -Type "Insight"

  ## Handle file import if requested
  if ($ImportDemoData) {
    if (-not $DemoDataFile) {
      Debug-Log "ImportDemoData specified but no file was provided (-DemoDataFile)." -Type "Problem"
      throw "Demo data import aborted: file not specified."
    }
    if (-not (Test-Path -LiteralPath $DemoDataFile)) {
      Debug-Log "Demo data file not found: $DemoDataFile" -Type "Problem"
      throw "Demo data import aborted: file $DemoDataFile does not exist."
    }
    Debug-Log "Importing demo data from file: $DemoDataFile" -Type "Insight"
    ## Import file - this will set forest/domain info from the CSV
    Initialise-DataSource -FilePath $DemoDataFile -Domain $null
    ## Set flags immediately after import
    $Script:DataFileLoaded = $true
    $Script:DataFilePath   = $DemoDataFile
    Show-InfoPanel -UpdateOnly
  } else {
    ## No file import - use default jukebox.example demo data
    Debug-Log "No file import - using default jukebox.example demo structure..." -Type "Insight"
    $Script:ForestName    = "jukebox.example"
    $Script:RootDomain    = "example.com"
    $Script:Domains       = @('example.com', 'example.net', 'example.org')
    $Script:CurrentDomain = $Script:RootDomain
  }
} else {
  Debug-Log "Production Mode: querying AD forest..." -Type "Insight"
  try {
    if ($Domain) {
      Debug-Log "Querying specified domain: $Domain" -Type "Insight"
      $targetDomain       = Get-ADDomain -Server $Domain -ErrorAction Stop
      $forest             = Get-ADForest -Server $targetDomain.Forest -ErrorAction Stop
    } else {
      $targetDomain       = Get-ADDomain -ErrorAction Stop
      $forest             = Get-ADForest -ErrorAction Stop
    }
    $Script:ForestName    = $forest.Name.Split('.')[0].ToUpper()
    $Script:RootDomain    = $forest.RootDomain
    $Script:Domains       = $forest.Domains
    $Script:Sites         = $forest.Sites | ForEach-Object { $_.Name }
    $Script:CurrentDomain = if ($Domain) { $Domain } else { $Script:RootDomain }
  } catch {
    Debug-Log ("Failed to query AD domain/forest: $_") -Type "Problem"
    Debug-Log ("Falling back to minimal domain info.") -Type "Warning"
    $Script:ForestName    = "DOMAIN"
    $Script:RootDomain    = if ($Domain) { $Domain } else { $env:USERDNSDOMAIN }
    $Script:Domains       = @($Script:RootDomain)
    $Script:Sites         = @()
    $Script:CurrentDomain = $Script:RootDomain
  }
}
$Script:Domain = $Script:CurrentDomain  ## Compatibility

## Initialise object arrays ONLY if file data wasn't loaded
if (-not $Script:DataFileLoaded) {
  Debug-Log "Initializing empty object arrays..." -Type "Tracing"
  $Script:CurrentDC       = $null
  $Script:Users           = @()
  $Script:Groups          = @()
  $Script:DCs             = @()
  $Script:Computers       = @()
  $Script:ADObjects       = @()
  $Script:SelectedObjects = @()
  $Script:SelectionMode   = $false
} else {
  Debug-Log "Skipping array initialization - file data already loaded" -Type "Insight"
}
## sneak sneak
Set-StatusBar "Refreshing users" -Icon Working -Percent 25

## ----------{Step 6: Load Domain Data }----------
if (-not $Script:DataFileLoaded) {
  Debug-Log "Loading domain data..." -Type "Insight"
  if ($Script:FilePathToLoad) {
    Debug-Log "File path specified: $($Script:FilePathToLoad)" -Type "Insight"
    Set-StatusBar "Loading data from file: $($Script:FilePathToLoad)..." -Icon 'Working' -Percent 10
  } else {
    Debug-Log "No file path - will use AD or demo data" -Type "Insight"
    Set-StatusBar "Loading domain data for $($Script:CurrentDomain)..." -Icon 'Working' -Percent 10
  }
  Set-StatusBar "Enumerating objects..." -Icon 'Working' -Percent 20
  Initialise-DataSource -FilePath $Script:FilePathToLoad -Domain $Script:CurrentDomain
  Set-StatusBar "Data load complete" -Icon 'Success' -Percent 50
} else {
  Debug-Log "Skipping data load - CSV data already loaded" -Type "Insight"
  Set-StatusBar "Ready" -Icon 'Success' -Percent 50
}

## Set Current DC after data is loaded
if ($Script:DCs -and $Script:DCs.Count -gt 0) {
  if (-not $Script:CurrentDC) {
    $Script:CurrentDC = $Script:DCs | Where-Object { $_.IsGlobalCatalog } | Select-Object -First 1
    if (-not $Script:CurrentDC) { $Script:CurrentDC = $Script:DCs | Select-Object -First 1 }
    $Script:CurrentDCName = $Script:CurrentDC.Name
    Debug-Log "Set current DC to: $($Script:CurrentDCName)" -Type "Insight"
  } else {
    $Script:CurrentDCName = $Script:CurrentDC.Name
    Debug-Log "CurrentDC already set to: $($Script:CurrentDCName), preserving" -Type "Tracing"
  }
} else {
  Debug-Log "No DCs available to set as CurrentDC" -Type "Warning"
  if (-not $Script:CurrentDC) {
    $Script:CurrentDC     = $null
    $Script:CurrentDCName = "(None)"
  }
}
Debug-Log "POST-LOAD: Users=$($Script:Users.Count), DCs=$($Script:DCs.Count), Computers=$($Script:Computers.Count), Group=$($Script:Groups.Count), Objects=$($Script:ADObjects.Count)" -Type "Insight"
Debug-Log "Forest/Domain initialization complete: CurrentDomain=$($Script:CurrentDomain)" -Type "Insight"

## Verify data was loaded
if ($Script:ADObjects.Count -eq 0 -and -not $Script:DemoMode) {
  Debug-Log "WARNING: No AD objects loaded in Production Mode!" -Type "Problem"
  Set-StatusBar "Warning: No objects loaded" -Icon 'Warning'
} else {
  Debug-Log "Verified: $($Script:ADObjects.Count) objects loaded successfully" -Type "Success"
}

## Now update filter panel with the current DC (already created, just updating)
Show-InfoPanel -UpdateOnly
$win.Add($infoPanel)
Debug-Log "InfoPanel created and updated with DC: $($Script:CurrentDC.Name ?? 'None')" -Type "Tracing"

## Now build the tree
Debug-Log "Initializing TreeView..." -Type "Insight"
Set-StatusBar "Building tree view..." -Icon 'Working' -Percent 60

$treeFrame = [Terminal.Gui.FrameView]::new("Active Directory Objects")
$treeFrame.X = 0
$treeFrame.Y = 0
$treeFrame.Width    = [Terminal.Gui.Dim]::Fill(42)
$treeFrame.Height   = [Terminal.Gui.Dim]::Fill()
$Script:tree        = [Terminal.Gui.TreeView]::new()

$Script:tree.X = 0
$Script:tree.Y = 0
$Script:tree.Width  = [Terminal.Gui.Dim]::Fill()
$Script:tree.Height = [Terminal.Gui.Dim]::Fill()

## Build and populate tree
Set-StatusBar "Populating tree..." -Icon 'Working' -Percent 80
Debug-Log "Building tree for domain: $($Script:CurrentDomain)" -Type "Insight"
Debug-Log "Objects available: Users=$($Script:Users.Count), Groups=$($Script:Groups.Count), Computers=$($Script:Computers.Count), DCs=$($Script:DCs.Count)" -Type "Insight"

$rootNode = Build-Tree -domain $Script:CurrentDomain
if ($null -eq $rootNode) {
  Debug-Log "FATAL - Build-Tree returned null root node" -Type "Problem"
  Set-StatusBar "Error: Tree build failed" -Icon 'Error'
} else {
  Debug-Log "Build-Tree completed, tree populated with root: $($rootNode.Text)" -Type "Success"
  Set-StatusBar "Tree built successfully" -Icon 'Success' -Percent 95
}

$treeFrame.Add($Script:tree)
$win.Add($treeFrame)
Debug-Log "TreeView created and added to window successfully" -Type "Success"
Set-StatusBar "Ready" -Icon 'Success' -Percent 100

## Global Key Handlers
$top.add_KeyPress({
  param($e)
  $handled = $false
  switch ($e.KeyEvent.Key) {
    ([Terminal.Gui.Key]::F1)  { Show-Modal "Shortcuts" "F1-Help | F2-Password | F3-New | F5-Refresh | F6-Themes | F7-Search | F8-Focus Tree | F10-Quit| F12-Context Menu" ; $handled = $true }
    ([Terminal.Gui.Key]::F2)  { Generate-RandomPassword ; $handled = $true }
    ([Terminal.Gui.Key]::F3)  { Show-NewObjectWizard    ; $handled = $true }
    ([Terminal.Gui.Key]::F5)  { Refresh-Data -domain $Script:CurrentDomain -RebuildTree ; $handled = $true }
    ([Terminal.Gui.Key]::F6)  { Show-ThemeSelector      ; $handled = $true }
    ([Terminal.Gui.Key]::F7)  { Show-ADSearchDialog     ; $handled = $true }
    ([Terminal.Gui.Key]::F8)  { if ($Script:tree) { $Script:tree.SetFocus(); Debug-Log "Tree focused via F8" -Type "Insight" } ; $handled = $true }
    ([Terminal.Gui.Key]::F10) { [Terminal.Gui.Application]::RequestStop() ; $handled = $true }
    ([Terminal.Gui.Key]::F12) {
      $selectedNode = $Script:tree.SelectedObject
      if ($selectedNode -and $selectedNode.Tag -and $selectedNode.Tag.Object) { Show-ObjectContextMenu -Object $selectedNode.Tag.Object -ObjectType $selectedNode.Tag.Type }
      $handled = $true
    }
  }
  $e.Handled = $handled
})
Set-StatusBar "Ready" -Icon 'Success'

## Debug view tree dump
if ($DebugMode -or $Logging) {
  Debug-Log "========== Full View Tree Dump ==========" -Type "Insight"
  Debug-DumpViewTree -View $top
  Debug-Log "========== End View Tree Dump ==========" -Type "Insight"
}

## Capture functions for closure
$debugLogFunc     = ${function:Debug-Log}
$buildContextFunc = ${function:Build-ContextMenuItems}
$showContextFunc  = ${function:Show-ContextMenu}

## ----------{ Step 8: Run Application }----------
Debug-Log "Starting Terminal.Gui main loop..." -Type "Success"
[Terminal.Gui.Application]::Run($top)

## ----------{ Cleanup }----------
Debug-Log "Application stopped, cleaning up..." -Type "Insight"
Set-StatusBar "Shutting down"
[Terminal.Gui.Application]::Shutdown()
Debug-Log "Application shut down cleanly" -Type "Success"
Debug-Log "End of line..." -Type "Insight"

if ($Script:LogStream) {
  $Script:LogStream.Close()
  $Script:LogStream.Dispose()
}
