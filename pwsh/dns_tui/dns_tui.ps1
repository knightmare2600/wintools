## Add F12 handler to tree
#requires -RunAsAdministrator

<#
===========================================================================================
DNS TUI — Terminal.Gui DNS Management Application
 - Inspired by easyDNS by Andreas Hepp (PHscripts.de)
 - TUI Framework re-imagined by knightmare2600 (github.com/knightmare2600)
===========================================================================================

Recent Changelog

1.0.5.04 (Function consolidation)

- Panels resized to give more prominence to the dashboard function
- The why citron menu entry was added
- Work around powershell unwrapping arrays despite being told not to in AuditLog-Dialog function
- Statusbar shortcut keys point ot DNS features now
- add support for parsing/importing a BIND9 zone (not the top priority!)
- Tools are now a pop-up modal again like dsa-tui
- Themes are not mandatory - if one isn't specified default to procomm
- Theme selection bug now fixed

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~{ TODO / COME BACK TO }~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Lemon emoji U+1F34B
$lemon = [char]0xD83C + [char]0xDF4B
Write-Host "Here’s a lemon: $lemon"

REMAINING FEATURES TO IMPLEMENT:

BUGS:
 - There is no close button on some modals, so you can't escape
 - F3 for create and a companion menu entry

#>

param(
  [switch]$Logging,
  [string]$LogFile,
  [string]$Domain, ## User can specify domain
  [ValidateSet("british", "class91", "dark", "db-1980s", "dsb", "gemstones", "intercity-swallow", "irn-bru", "light", "matrix", "ns", "network-southeast", "panam", "procomm", "scotrail", "twa", "viarail", "viarail-soft")]
  [string]$Theme,
  [Bool]$LaunchReady ## Will check prerequisites only then exit
)

## $PSBoundParameters contains only the parameters that were actually provided. But $args contains ALL command-line
## arguments, including invalid ones. Extract parameter names from command line (anything starting with -)
$providedParams = $args | Where-Object { $_ -match '^-' } | ForEach-Object {  $_ -replace '^-+', '' }

## Define valid parameter names (including aliases)
$validParams = @(
  'DemoMode', 'Logging', 'LogFile', 'Domain', 'ImportDemoData', 'DemoDataFile', 'Theme', 'LaunchReady',
  'Verbose', 'Debug', 'ErrorAction', 'WarningAction', 'WhatIf', 'Confirm'  ## Common parameters built in to pwsh - DO NOT USE!
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
  Write-Host ".\dnstui_tui.ps1 -DemoMode -Theme panam -ImportDemoData -DemoDataFile .\data.csv`n" -ForegroundColor White
  exit 1
}

## Debug when errors scroll past too fast, and Debug-Log can't capture them
#$ErrorActionPreference = 'Stop'

## -------------------------{ Global Variables }--------------------------
## Define the build version, project and code names once only - up here to ease patching. The rest in main
## execution loop, where they belong
$Script:ProjectName  = "DNS-TUI Powershell dnsmgmt.msc"
$Script:FruitName    = "Citron" ## Danish for Lemon
$Script:EnFruitName  = "Lemon"
$Script:FruitEmoji   = [char]0xD83C + [char]0xDF4B
$Script:BuildVersion = "3.2.6.11"
$Script:ThemeMode    = $Theme

## Global emojis
$Script:Icons = @{
  ## Core DNS object types
  Server       = "🖥️"   # DNS Server
  Zone         = "🌐"    # DNS Zone
  ForwardZone  = "➡️"    # Forward Lookup Zone
  ReverseZone  = "⬅️"    # Reverse Lookup Zone
  Record       = "📄"    # DNS Record

  ## DNS Record types
  A_Record     = "🔤"    # A Record (IPv4)
  AAAA_Record  = "🔢"    # AAAA Record (IPv6)
  CNAME        = "🔗"    # Alias/CNAME
  MX           = "📧"    # Mail Exchange
  NS           = "🌐"    # Name Server
  PTR          = "👉"    # Pointer Record (CHANGED!)
  TXT          = "📝"    # Text Record
  SOA          = "👑"    # Start of Authority
  SRV          = "⚙️"    # Service Record

  ## Status indicators
  Active       = "✔️"    # Active/Online
  Inactive     = "⊗"     # Inactive/Offline
  Enabled      = "●"     # Enabled
  Disabled     = "○"     # Disabled
  Error        = "✖"     # Error
  Warning      = "⚠️"    # Warning
  Success      = "✓"     # Success
  Working      = "⏳"    # In Progress

  ## DNS Operations
  Lookup       = "🔍"    # DNS Lookup
  Query        = "❓"    # DNS Query
  Cache        = "💾"    # DNS Cache
  Refresh      = "🔄"    # Refresh/Reload
  Import       = "📥"    # Import Zone
  Export       = "📤"    # Export Zone
  Connect      = "🔌"    # Connect to Server
  Disconnect   = "🔌"    # Disconnect
  Diagnostic   = "🔧"    # Diagnostic Tools

  ## Network/DNS specific
  Network      = "🌐"    # Network
  Globe        = "🌍"    # Global DNS
  Satellite    = "📡"    # DNS Resolution
  Router       = "📶"    # Network Router
  Firewall     = "🛡️"    # Firewall/Security
}

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
          if ($InstallMsg) { Debug-Log "Please run: $InstallMsg" -Type "Tracing" }
          return $false
        }
        return $true
      } else {
        if ($Optional) {
          Debug-Log "Optional module '$Name' is NOT installed." -Type "Warning"
        } else {
          Debug-Log "Module '$Name' is NOT installed." -Type "Problem"
        }
        if ($InstallMsg) { Debug-Log "Please run: $InstallMsg" -Type "Tracing" }
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
        if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Tracing" }
        return $false
      }
      if (-not $IsAdmin) {
        Debug-Log "Cannot check/install Windows Feature '$Name' — requires admin privileges." -Type "Warning"
        if ($InstallMsg) { Debug-Log "Suggested action (run elevated): $InstallMsg" -Type "Tracing" }
        return $false
      }
      try {
        Import-Module ServerManager -ErrorAction Stop
        $feature = Get-WindowsFeature $Name -ErrorAction SilentlyContinue
      } catch {
        Debug-Log "Failed to query Windows Feature '$Name': $_" -Type "Warning"
        if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Tracing" }
        return $false
      }
      if ($feature) {
        if ($feature.Installed) {
          Debug-Log "Windows Feature '$Name' is already installed." -Type "Success"
          return $true
        } else {
          Debug-Log "Windows Feature '$Name' is NOT installed." -Type "Problem"
          if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Tracing" }
          return $false
        }
      } else {
        Debug-Log "Windows Feature '$Name' not found on this system." -Type "Warning"
        if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Tracing" }
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
        if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Tracing" }
        return $false
      }
      if ($cap) {
        if ($cap.State -eq 'Installed') {
          Debug-Log "Windows Capability '$Name' is already installed." -Type "Success"
          return $true
        } else {
          Debug-Log "Windows Capability '$Name' is NOT installed." -Type "Problem"
          if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Tracing" }
          return $false
        }
      } else {
        Debug-Log "Windows Capability '$Name' not found on this system." -Type "Warning"
        if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Tracing" }
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
        if ($InstallMsg) { Debug-Log "Suggested action: $InstallMsg" -Type "Tracing" }
        return $false
      }
    }
  }
}

## ----------{ Is it a special day...? }----------
## Determines which unicode emoji to use for the AD window title.
function Initialise-DirectoryEmoji {
  param(
    [DateTime]$Date = (Get-Date)
  )

  $month = $Date.Month
  $day   = $Date.Day

  ## Default: Globe/DNS symbol
  $emoji = "🌐"

  ## ----------{ Icon Initialisation }----------
  if ($Script:HasTerminalIcons) {
    ## Nerd Font / Symbol icons (explicit Unicode escapes)
    $Script:DNSServerEmoji = "`u{F233}"   ## nf-fa-server
    $Script:ZoneEmoji      = "`u{F0AC}"   ## nf-fa-globe
    $Script:RecordEmoji    = "`u{F016}"   ## nf-fa-file
    $Script:NetworkEmoji   = "`u{F6FF}"   ## nf-fa-network-wired
    $Script:QueryEmoji     = "`u{F002}"   ## nf-fa-search
    $Script:WarningEmoji   = "`u{F071}"   ## nf-fa-exclamation_triangle
    $Script:OkEmoji        = "`u{F00C}"   ## nf-fa-check
    $Script:ErrorEmoji     = "`u{F00D}"   ## nf-fa-times
    $Script:WorkingEmoji   = "`u{F021}"   ## nf-fa-sync
  } else {
    ## ASCII fallbacks (alignment-safe)
    $Script:DNSServerEmoji = "[DNS]"
    $Script:ZoneEmoji      = "[ZONE]"
    $Script:RecordEmoji    = "[REC]"
    $Script:NetworkEmoji   = "[NET]"
    $Script:QueryEmoji     = "[?]"
    $Script:WarningEmoji   = "!"
    $Script:OkEmoji        = "OK"
    $Script:ErrorEmoji     = "X"
    $Script:WorkingEmoji   = "..."
  }

  ## Special days with DNS/Network theme
  switch ($true) {
    { $month -eq 1  -and $day -eq 1  }                   { $emoji = "🎆" ; break }                    ## 1st Jan  - New Year's Day
    { $month -eq 1  -and $day -eq 2  }                   { $emoji = "🦄" ; break }                    ## 2nd Jan  - Wild haggis Hunting
    { $month -eq 3  -and $day -eq 15 }                   { $emoji = "🌐" ; break }                    ## 15th Mar - Happy birthday .com domain!
    { $month -eq 4  -and $day -eq 9  }                   { $emoji = "🇩🇰" ; break }                     ## 9th Apr  - Danmark's liberation day
    { $month -eq 4  -and $day -eq 30 }                   { $emoji = "🌐" ; break }                    ## 30th Apr - First website went live (1993)
    { $month -eq 5  -and $day -eq 4  }                   { $emoji = "🕯️" ; break }                    ## 4th May  - Candle for Besættelsen
    { $month -eq 6  -and $day -eq 5  }                   { $emoji = "🇩🇰" ; break }                    ## 5th June  - Constitution day in Danmark
    { $month -eq 5  -and $day -eq 21 }                   { $emoji = "🇬🇱" ; break }                    ## 21st May  - Grønland Day
    { $month -eq 7  -and $day -eq 1  }                   { $emoji = "🇨🇦" ; break }                    ## 1st Jul   - Canada Day
    { $month -eq 7  -and $day -eq 4  }                   { $emoji = "🫖" ; break }                    ## 4th Jul  - Teapot (to annoy Americans)
    { $month -eq 7  -and $day -eq 29 }                   { $emoji = "🇫🇴" ; break }                    ## 29th Jul  - Faroe Islands
    { $month -eq 10 -and $day -eq 29 }                   { $emoji = "📡" ; break }                    ## 29th Oct - Internet created (ARPANET 1969)
    { $month -eq 11 -and $day -eq 9  }                   { $emoji = "🇩🇪" ; break }                    ## 9th Nov   - Fall of Berlin Wall
    { $month -eq 11 -and $day -eq 24 }                   { $emoji = "👑" ; break }                   ## 24th Nov  - En gang til Prins Knud
    { $month -eq 11 -and $day -eq 30 }                   { $emoji = "🏴󠁧󠁢󠁳󠁣󠁴󠁿" ; break }                   ## 30th Nov  - St Andrew's Day (Saltire)
    { $month -eq 12 -and ($day -eq 24 -or $day -eq 25) } { $emoji = "🎄" ; break }                   ## 24/25th Dec  - Tree for Jul / Christmas
    { $month -eq 12 -and $day -eq 31 }                   { $emoji = "🎆" ; break }                   ## 31st Dec  - New Year's Eve
  }

  Debug-Log "The emoji today is: $emoji" -Type "Insight"
  $Script:DirectoryEmoji = $emoji
}

## ----------{ Icon initialisation function }----------
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

## ----------{ Helper function to get icon }----------
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

## ----------{ Common tab builders / helper functions for UI building- re-usable across dialogs }----------
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
      $objectType = if ($SearchTabConfig.ObjectType) {
        $SearchTabConfig.ObjectType
      } elseif ($Data.ObjectClass -eq 'user') { 'User'
      } elseif ($Data.ObjectClass -eq 'group') { 'Group'
      } elseif ($Data.ObjectClass -eq 'computer') { 'Computer'
      } else { 'Object' }
      $searchTypes = $SearchTabConfig.SearchTypes ?? @("$objectType", "Group", "User", "Computer", "OU")

      $searchTab = @{
        Name = "Search/Lookup"
        Builder = {
          param($view, $data, $state)
          $y = 1
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Domain:" -FieldName 'txtSearchDomain' -State $state -Value $Script:CurrentDomain -Width 30
          $y += 1
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Name:" -FieldName 'txtSearchName' -State $state -Value $data.Name -Width 30
          $y += 1
          $lblType = [Terminal.Gui.Label]::new("Type:")
          $lblType.X = 2; $lblType.Y = $y
          $view.Add($lblType)
          $state.cmbSearchType = [Terminal.Gui.ComboBox]::new()
          $state.cmbSearchType.X = 15; $state.cmbSearchType.Y = $y; $state.cmbSearchType.Width = 20
          $state.cmbSearchType.SetSource($searchTypes)
          $state.cmbSearchType.SelectedItem = 0
          $view.Add($state.cmbSearchType)
          $y += 2
          $lblFilter = [Terminal.Gui.Label]::new("Filter Results:")
          $lblFilter.X = 50; $lblFilter.Y = 1
          $view.Add($lblFilter)
          $state.txtSearchFilter = [Terminal.Gui.TextField]::new("")
          $state.txtSearchFilter.X = 67; $state.txtSearchFilter.Y = 1; $state.txtSearchFilter.Width = 20
          $view.Add($state.txtSearchFilter)
          $state.txtSearchFilter.add_TextChanged({
            if ($state.currentSearchOutputLines) {
              $search = $state.txtSearchFilter.Text.ToString().Trim()
              if ($search) {
                $state.txtSearchOutput.Text = ($state.currentSearchOutputLines | Where-Object { $_ -match "(?i)$search" }) -join "`n"
              } else {
                $state.txtSearchOutput.Text = $state.currentSearchOutputLines -join "`n"
              }
            }
          }.GetNewClosure())

          $lblResult = [Terminal.Gui.Label]::new("Properties:")
          $lblResult.X = 2; $lblResult.Y = $y
          $view.Add($lblResult)
          $y += 1
          $state.txtSearchOutput = [Terminal.Gui.TextView]::new()
          $state.txtSearchOutput.X = 2; $state.txtSearchOutput.Y = $y
          $state.txtSearchOutput.Width = [Terminal.Gui.Dim]::Fill(2)
          $state.txtSearchOutput.Height = [Terminal.Gui.Dim]::Fill(1)
          $state.txtSearchOutput.ReadOnly = $true
          $state.txtSearchOutput.WordWrap = $false
          $view.Add($state.txtSearchOutput)

          if (${function:Set-ObjectCheckboxes}) { Set-ObjectCheckboxes -View $view -State $state -Data $data -ObjectType $objectType -Mode 'Create' }
          $btnSearch = [Terminal.Gui.Button]::new("Search")
          $btnSearch.X = 50; $btnSearch.Y = 3
          $view.Add($btnSearch)
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
              if (${function:Set-ObjectCheckboxes}) { Set-ObjectCheckboxes -State $state -Data $data -ObjectType $objectType -Mode 'Update' }
            }
          }.GetNewClosure())
        }
      }

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

    ## Create dialog with explicit width/height (NOT in constructor)
    Debug-Log "Creating dialog with buttons" -Type "Tracing"
    $dialog = [Terminal.Gui.Dialog]::new()
    $dialog.Title = $Title
    $dialog.Width = $Width
    $dialog.Height = $Height

    # Add buttons
    $dialog.AddButton($btnOK)
    $dialog.AddButton($btnCancel)
    $dialog.AddButton($btnApply)

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

## -----------------------------------{ Menus }-----------------------------------
function Build-MainMenu {
  [CmdletBinding()]
  param()

  Debug-Log "Building main menu..." -Type "Tracing"
  ## ----------{ Menu Items }----------
  # Define all menu items
  $mFile        = [Terminal.Gui.MenuItem]::new("_Exit", "Exit application (F10)", [Action]{ [Terminal.Gui.Application]::RequestStop() })
  $mNew         = [Terminal.Gui.MenuItem]::new("F_orward Zone", "Create a Forward Zone", [Action]{ Show-DNSDialog })
  $mProps       = [Terminal.Gui.MenuItem]::new("_Reverse Zone", "Create a Reverse Zone", [Action]{ Show-DNSDialog })
  $mShowRecords = [Terminal.Gui.MenuItem]::new("_Show Records", "Show Zone DNS Records", [Action]{ Show-DNSDialog })
  $mShowTools   = [Terminal.Gui.MenuItem]::new("_Tools", "Show Installed DNS tools", [Action]{ Show-Tools })
  $mRefresh     = [Terminal.Gui.MenuItem]::new("_Refresh", "Refresh DNS data (F5)", [Action]{ Show-Dashboard})
  $mTheme       = [Terminal.Gui.MenuItem]::new("_Theme", "Change theme (F6)", [Action]{ Show-ThemeSelector })
  $mDNSDialog   = [Terminal.Gui.MenuItem]::new("_Show DNS Dialog", "Show DNS Dialog", [Action]{ Show-DNSDialog })
  $mDNSLookup   = [Terminal.Gui.MenuItem]::new("DNS L_ookup", "Perform DNS Lookups", [Action]{ Show-DNSLookupDialog })
  $mDNSTUI      = [Terminal.Gui.MenuItem]::new("_DNS TUI", "DNS TUI", [Action]{ Show-DnsTui })
  $mShortcuts   = [Terminal.Gui.MenuItem]::new("_Shortcuts", "Keyboard shortcuts (F1)", [Action]{ Show-Modal "Shortcuts" "F1  - Help`nF3  - Forward Zones`nF4  - Reverse Zones`nF5  - Refresh`nF6  - Theme`nF10 - Quit" })
  $mAboutDNSTUI = [Terminal.Gui.MenuItem]::new("Abou_t", "About Project", [Action]{ Show-Modal "About" "DNS TUI`nVersion: $($Script:BuildVersion)`nGPL-3 License" })
  $mWhyCitron   = [Terminal.Gui.MenuItem]::new("_Why Citron", "About Project Codename", [Action]{ Show-CitronInfo })

  ## Create MenuBar with ONLY defined items
  $menu = [Terminal.Gui.MenuBar]::new(@(
    [Terminal.Gui.MenuBarItem]::new("_File", @($mRefresh, $mTheme, $mFile)),
    [Terminal.Gui.MenuBarItem]::new("_Zones", @($mNew, $mProps, $mShowRecords, $mDNSTUI)),
    [Terminal.Gui.MenuBarItem]::new("_Tools", @($mShowTools, $mDNSLookup, $mDNSDialog)),
    [Terminal.Gui.MenuBarItem]::new("_Help", @($mShortcuts, $mAboutDNSTUI, $mWhyCitron))
  ))

  Debug-Log "Main menu created successfully" -Type "Insight"
  return $menu
}

## ----------{ Helper funcitons for UI building }----------
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
  $script:currentPath  = (Resolve-Path $StartDir).Path
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

  ## ----------{ Update file list helper function }----------
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

## ----------{ Set up the UI }----------
function Initialise-UIFramework {
  <#
  .SYNOPSIS
  Initialise Terminal.Gui application and create main window with menu and status bars

  .DESCRIPTION
  Sets up the Terminal.Gui framework, creates the top-level application,
  main window, menu bar, and status bar, then applies the selected theme.
  This should be called FIRST before any data loading to ensure the UI is visible.

  .PARAMETER Theme
  Theme name to apply (procomm, matrix, british, etc)

  .PARAMETER Title
  Window title to display

  .EXAMPLE
  $uiComponents = Initialise-UIFramework -Theme "procomm" -Title "DNS-TUI v1.0"
  $top = $uiComponents.Top
  $win = $uiComponents.Window
  $menu = $uiComponents.Menu
  $statusBar = $uiComponents.StatusBar
  #>

  param(
    [string]$Theme = "procomm",
    [string]$Title = "DNS-TUI - Active Directory"
  )

  Debug-Log "Initializing Terminal.Gui framework..." -Type "Tracing"
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
  ## ----------{ Step 3: Create Status Bar }---------
  Debug-Log "Creating status bar..." -Type "Tracing"
  $statusBar = Set-StatusBar -Initialise
  ## ----------{ Step 4: Create Main Menu }---------
  Debug-Log "Creating main menu..." -Type "Tracing"
  $menu = Build-MainMenu
  ## ----------{ Step 5: Apply Theme }---------
  Debug-Log "Applying theme: $Theme" -Type "Tracing"
  $Script:ThemeMode = $Theme
  $themeData        = Get-Theme -mode $Theme
  if ($themeData) {
    ## Store theme data globally
    $Script:themeData = $themeData
    ## Use Apply-Theme to handle all components properly
    ## Note: Menu and StatusBar don't exist yet, so pass $null
    Apply-Theme -ThemeData $themeData -TopLevel $top -MainWindow $win -Menu $menu -StatusBar $statusBar
    Debug-Log "Theme '$Theme' applied successfully" -Type "Success"
  } else { Debug-Log "WARNING - Theme data is null, using defaults" -Type "Warning" }
  ## ----------{ Step 6: Add Components to Top - IN ORDER }---------
  ## Add menu FIRST
  Debug-Log "Adding menu to top..." -Type "Tracing"
  $top.Add($menu)
  $top.LayoutSubviews()
  $top.SetNeedsDisplay()
  ## Then status bar
  Debug-Log "Adding status bar to top..." -Type "Tracing"
  $top.Add($statusBar)
  $top.LayoutSubviews()
  $top.SetNeedsDisplay()
  ## Finally the main window
  Debug-Log "Adding window to top..." -Type "Tracing"
  $top.Add($win)
  $top.LayoutSubviews()
  $top.SetNeedsDisplay()
  Debug-Log "All components added and laid out" -Type "Success"
  ## ----------{ Step 7: Return Components }---------
  $result = @{
    Top       = $top
    Window    = $win
    Menu      = $menu
    StatusBar = $statusBar
  }
  Debug-Log "UI Framework initialization complete" -Type "Success"
  Debug-Log "UI Framework ready - window visible to user" -Type "Success"
  return $result
}

## ----------{ Statusbar }----------
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
    Debug-Log "Initializing status bar..." -Type "Tracing"

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
      @{ Key = [Terminal.Gui.Key]::F1;  Label = "~F1~ Help";         Action = { Show-Modal "Shortcuts" "F1  - Help`nF3  - Forward Zones`nF4  - Reverse Zones`nF5  - Refresh`nF6  - Theme`nF10 - Quit" } }
      @{ Key = [Terminal.Gui.Key]::F2;  Label = "~F2~ Password";      Action = { Generate-RandomPassword } }
      @{ Key = [Terminal.Gui.Key]::F3;  Label = "~F3~ Forward Zones"; Action = { Show-NewObjectWizard } }
      @{ Key = [Terminal.Gui.Key]::F5;  Label = "~F4~ Reverse Zones"; Action = { Refresh-Data -domain $Script:CurrentDomain -RebuildTree } }
      @{ Key = [Terminal.Gui.Key]::F6;  Label = "~F6~ Themes";       Action = { Show-ThemeSelector } }
      @{ Key = [Terminal.Gui.Key]::F8;  Label = "~F8~ Focus Tree";   Action = { } }
      @{ Key = [Terminal.Gui.Key]::F9;  Label = "~F9~ Menus";        Action = { } }
      @{ Key = [Terminal.Gui.Key]::F10; Label = "~F10~ Quit";        Action = { [Terminal.Gui.Application]::RequestStop() } }
      @{ Key = [Terminal.Gui.Key]::F11; Label = "~F11~ Full Screen"; Action = { } }
      @{ Key = ([Terminal.Gui.Key]::F12); Label = "~F12~ Context Menu"; Action = { } }
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

## ----------{ Dump tree for debug }---------
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

## ----------{ Show modal function to reduce clutter }----------
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
    Debug-Log "Dumping colour scheme for theme: $themeName" -Type "Tracing"
    ## Dump the ColorSchemes that are actually being used
    Debug-Log "=== Script ColorScheme ===" -Type "Tracing"
      if ($Script:ScriptCs) {
        Debug-Log "Normal    : $($Script:ScriptCs.Normal)" -Type "Tracing"
        Debug-Log "Focus     : $($Script:ScriptCs.Focus)" -Type "Tracing"
        Debug-Log "HotNormal : $($Script:ScriptCs.HotNormal)" -Type "Tracing"
        Debug-Log "HotFocus  : $($Script:ScriptCs.HotFocus)" -Type "Tracing"
        Debug-Log "Disabled  : $($Script:ScriptCs.Disabled)" -Type "Tracing"
      } else {
        Debug-Log "ScriptCs is null!" -Type "Warning"
      }
      Debug-Log "=== Main Window ColorScheme ===" -Type "Tracing"
      if ($Script:mainWindowCs) {
        Debug-Log "Normal    : $($Script:mainWindowCs.Normal)" -Type "Tracing"
        Debug-Log "Focus     : $($Script:mainWindowCs.Focus)" -Type "Tracing"
        Debug-Log "HotNormal : $($Script:mainWindowCs.HotNormal)" -Type "Tracing"
        Debug-Log "HotFocus  : $($Script:mainWindowCs.HotFocus)" -Type "Tracing"
        Debug-Log "Disabled  : $($Script:mainWindowCs.Disabled)" -Type "Tracing"
      } else {
        Debug-Log "mainWindowCs is null!" -Type "Warning"
      }
      return
    }
    ## Load Mode
    if (-not $Mode) {
      $Mode = "procomm"
      Debug-Log "No theme specified, defaulting to procomm" -Type "Insight"
    }
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
        $Script:mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::BrightYellow,[Terminal.Gui.Color]::Blue)
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

  ## Create simple dialog - NO SPINNER (Terminal.Gui 1.16 doesn't support AddTimeout)
  $dlg = [Terminal.Gui.Dialog]::new()
  $dlg.Title = "Please Wait"
  $dlg.Width = 50
  $dlg.Height = 7

  ## Message label
  $lbl = [Terminal.Gui.Label]::new()
  $lbl.X = [Terminal.Gui.Pos]::Center()
  $lbl.Y = 2
  $lbl.Text = $Message
  $dlg.Add($lbl)

  ## Working indicator (static - no animation possible)
  $indicator = [Terminal.Gui.Label]::new()
  $indicator.X = [Terminal.Gui.Pos]::Center()
  $indicator.Y = 3
  $indicator.Text = "⏳ Working..."
  $dlg.Add($indicator)

  ## Return dialog for caller to manage
  return $dlg
}

## ----------{ Danske Soda vand }----------
## This is a theme now. Danish Fruit soda based fun. Method to the madness
function Show-CitronInfo {
  $message = @"
$($Script:ProjectName) is codenamed $Script:FruitName because:

- I was drinking lemon soda when writing the code
- $($Script:FruitName) is Danish for $($Script:ENFruitName)
- Føtex sells a rather nice $($Script:FruitName) soda
- Every great project needs a forest-fruit mascot!
"@
  Show-Modal "Why $($Script:FruitName)...? $Script:FruitEmoji" $message -EasterEgg
}

## ~~~~~~~~~~~~~~~~~~~~~~~~~{ DNS FUNCTIONS BELOW HERE }~~~~~~~~~~~~~~~~~~~~~~~~~

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

  Debug-Log "Opening DNS Lookup tool" -Type "Tracing"

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
  $debugLogFunc  = ${function:Debug-Log}
  $showModalFunc = ${function:Show-Modal}
  $icons         = if ($Script:Icons) { $Script:Icons } else { @{ Success = "✔"; Error = "✖" } }

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
    & $debugLogFunc "DNS lookup form cleared" -Type "Tracing"
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
    & $debugLogFunc "DNS Lookup tool closed" -Type "Tracing"
    [Terminal.Gui.Application]::RequestStop()
  }.GetNewClosure())
  $dialog.Add($btnClose)
  ## Run dialog
  [Terminal.Gui.Application]::Run($dialog)
}

## ----------{ BIND Zone File Import with Conflict Resolution }----------
## TODO: did this function go walkies...?

## ----------{ DNS Object Context Menu - F12 }----------
function Show-DNSObjectContextMenu {
  <#
  .SYNOPSIS
  Shows and handles a context menu for DNS objects
  .PARAMETER Object
  The DNS object (Server, Zone, Record)
  .PARAMETER ObjectType
  Type of object: 'Server', 'Zone', 'ForwardZone', 'ReverseZone', 'Record'
  #>
  param(
    [Parameter(Mandatory=$true)]
    [object]$Object,
    [Parameter(Mandatory=$true)]
    [ValidateSet('Server', 'Zone', 'ForwardZone', 'ReverseZone', 'Record')]
    [string]$ObjectType
  )

  Debug-Log "Showing DNS context menu for $($Object.Name ?? $Object.ZoneName ?? 'DNS Server') (Type: $ObjectType)" -Type "Tracing"

  ## Build menu items based on object type
  $menuItems = [System.Collections.ArrayList]@()

  switch ($ObjectType) {
    'Server' {
      [void]$menuItems.Add("🔌 Connect to Server...")
      [void]$menuItems.Add("➕ Create Forward Zone...")
      [void]$menuItems.Add("➕ Create Reverse Zone...")
      [void]$menuItems.Add("📋 View All Zones & Records")
      [void]$menuItems.Add("📥 Import Zone File...")
      [void]$menuItems.Add("🔧 Diagnostic Tools...")
      [void]$menuItems.Add("🔄 Refresh Tree")
    }
    { $_ -in 'Zone', 'ForwardZone', 'ReverseZone' } {
      [void]$menuItems.Add("📄 Zone Properties")
      [void]$menuItems.Add("➕ Create Record...")
      [void]$menuItems.Add("📋 View Records")
      [void]$menuItems.Add("📥 Import Records...")
      [void]$menuItems.Add("📤 Export Zone...")
      [void]$menuItems.Add("🔄 Reload Zone")
      [void]$menuItems.Add("🗑️ Delete Zone")
      [void]$menuItems.Add("🔄 Refresh")
    }
    'Record' {
      [void]$menuItems.Add("📄 Record Properties")
      [void]$menuItems.Add("✏️ Edit Record...")
      [void]$menuItems.Add("🗑️ Delete Record")
      [void]$menuItems.Add("🔄 Refresh")
    }
  }

  ## Create context menu dialog
  $menuDlg = [Terminal.Gui.Dialog]::new()
  $menuDlg.Title = "Context Menu - $ObjectType"
  $menuDlg.Width = 50
  $menuDlg.Height = $menuItems.Count + 8

  ## Add object name label
  $objectName = switch ($ObjectType) {
    'Server' { $global:DetectedDnsServer }
    { $_ -in 'Zone', 'ForwardZone', 'ReverseZone' } { $Object.ZoneName ?? $Object.Name }
    'Record' { "$($Object.HostName) [$($Object.RecordType)]" }
  }

  $lblName = [Terminal.Gui.Label]::new("Object: $objectName")
  $lblName.X = 2
  $lblName.Y = 1
  $menuDlg.Add($lblName)

  ## Create ListView for menu items
  $lstMenu = [Terminal.Gui.ListView]::new()
  $lstMenu.X = 2
  $lstMenu.Y = 3
  $lstMenu.Width = [Terminal.Gui.Dim]::Fill(2)
  $lstMenu.Height = $menuItems.Count
  $lstMenu.SetSource($menuItems)
  $menuDlg.Add($lstMenu)

  ## Capture functions
  $showConnectFunc       = ${function:Show-ConnectDialog}
  $showCreateForwardFunc = ${function:Show-CreateForwardZoneDialog}
  $showCreateReverseFunc = ${function:Show-CreateReverseZoneDialog}
  $showDNSDialogFunc     = ${function:Show-DNSDialog}
  $showImportFunc        = ${function:Show-ImportDialog}
  $showExportFunc        = ${function:Show-ExportDialog}
  $showToolsFunc         = ${function:Show-Tools}
  $showZonePropsFunc     = ${function:Show-DNSZoneProperties}
  $showRecordPropsFunc   = ${function:Show-DNSRecordProperties}
  $showCreateRecordFunc  = ${function:Show-CreateDNSRecordDialog}
  $showModalFunc         = ${function:Show-Modal}
  $debugLogFunc          = ${function:Debug-Log}
  $buildTreeFunc         = ${function:Build-DNSZoneTree}

  ## Handle selection
  $selectedAction = $null

  $btnOK = [Terminal.Gui.Button]::new("OK")
  $btnOK.add_Clicked({
    $selectedIndex = $lstMenu.SelectedItem
    if ($selectedIndex -ge 0 -and $selectedIndex -lt $menuItems.Count) {
      $selectedAction = $menuItems[$selectedIndex]
    }
    [Terminal.Gui.Application]::RequestStop()
  }.GetNewClosure())
  $menuDlg.AddButton($btnOK)

  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $menuDlg.AddButton($btnCancel)

  ## Show menu
  [Terminal.Gui.Application]::Run($menuDlg)

  ## Execute selected action
  if ($null -eq $selectedAction) { return }

  & $debugLogFunc "Context menu action: $selectedAction" -Type "Insight"

  switch ($ObjectType) {
    'Server' {
      switch ($selectedAction) {
        "🔌 Connect to Server..." { & $showConnectFunc }
        "➕ Create Forward Zone..." { & $showCreateForwardFunc }
        "➕ Create Reverse Zone..." { & $showCreateReverseFunc }
        "📋 View All Zones & Records" { & $showDNSDialogFunc -ServerName $global:DetectedDnsServer }
        "📥 Import Zone File..." { & $showImportFunc }
        "🔧 Diagnostic Tools..." { & $showToolsFunc }
        "🔄 Refresh Tree" {
          if ($global:DNSConnectionStatus.IsConnected) {
            & $buildTreeFunc -dnsServer $global:DetectedDnsServer
          }
        }
      }
    }

    { $_ -in 'Zone', 'ForwardZone', 'ReverseZone' } {
      $zoneName = $Object.ZoneName ?? $Object.Name
      $zoneType = if ($Object.IsReverse -or $Object.IsReverseLookupZone) { 'Reverse' } else { 'Forward' }

      switch ($selectedAction) {
        "📄 Zone Properties" { & $showZonePropsFunc -ZoneName $zoneName -ZoneType $zoneType }
        "➕ Create Record..." { & $showCreateRecordFunc -ZoneName $zoneName -ServerName $global:DetectedDnsServer }
        "📋 View Records" {
          # Open DNS dialog and auto-select this zone
          & $showDNSDialogFunc -ServerName $global:DetectedDnsServer
        }
        "📥 Import Records..." { & $showImportFunc }
        "📤 Export Zone..." { & $showExportFunc }
        "🔄 Reload Zone" {
          $zoneTypeStr = if ($Object.ZoneType) { $Object.ZoneType } else { "Primary" }
          Reload-DNSZone -ZoneName $zoneName -ZoneType $zoneTypeStr
          & $showModalFunc "Success" "Zone reloaded: $zoneName"
        }
        "🗑️ Delete Zone" {
          Delete-Zone -ZoneName $zoneName
          if ($global:DNSConnectionStatus.IsConnected) {
            & $buildTreeFunc -dnsServer $global:DetectedDnsServer
          }
        }
        "🔄 Refresh" {
          if ($global:DNSConnectionStatus.IsConnected) {
            & $buildTreeFunc -dnsServer $global:DetectedDnsServer
          }
        }
      }
    }

    'Record' {
      # For records, we need the zone name too
      $zoneName = $Object.ZoneName ?? ""

      switch ($selectedAction) {
        "📄 Record Properties" {
          if ([string]::IsNullOrWhiteSpace($zoneName)) {
            & $showModalFunc "Error" "Zone name not available for this record"
            return
          }
          & $showRecordPropsFunc -Record $Object -ZoneName $zoneName -ServerName $global:DetectedDnsServer
        }
        "✏️ Edit Record..." {
          if ([string]::IsNullOrWhiteSpace($zoneName)) {
            & $showModalFunc "Error" "Zone name not available for this record"
            return
          }
          & $showRecordPropsFunc -Record $Object -ZoneName $zoneName -ServerName $global:DetectedDnsServer
        }
        "🗑️ Delete Record" {
          if ([string]::IsNullOrWhiteSpace($zoneName)) {
            & $showModalFunc "Error" "Zone name not available for this record"
            return
          }

          # Confirm
          $confirmDlg = [Terminal.Gui.Dialog]::new()
          $confirmDlg.Title = "Confirm Delete"
          $confirmDlg.Width = 60
          $confirmDlg.Height = 12

          $lblConfirm = [Terminal.Gui.Label]::new("Delete DNS record?`n`nHostname: $($Object.HostName)`nType: $($Object.RecordType)`nZone: $zoneName")
          $lblConfirm.X = 2
          $lblConfirm.Y = 1
          $confirmDlg.Add($lblConfirm)

          $confirmResult = 1
          $btnYes = [Terminal.Gui.Button]::new("Yes")
          $btnYes.add_Clicked({ $confirmResult = 0; [Terminal.Gui.Application]::RequestStop() })
          $confirmDlg.AddButton($btnYes)

          $btnNo = [Terminal.Gui.Button]::new("No")
          $btnNo.add_Clicked({ $confirmResult = 1; [Terminal.Gui.Application]::RequestStop() })
          $confirmDlg.AddButton($btnNo)

          [Terminal.Gui.Application]::Run($confirmDlg)

          if ($confirmResult -eq 0) {
            try {
              Remove-DnsServerResourceRecord -ZoneName $zoneName -InputObject $Object -ComputerName $global:DetectedDnsServer -Force -ErrorAction Stop
              & $showModalFunc "Success" "DNS record deleted!"
              & $debugLogFunc "Deleted: $($Object.HostName) [$($Object.RecordType)]" -Type "Success"

              if ($global:DNSConnectionStatus.IsConnected) {
                & $buildTreeFunc -dnsServer $global:DetectedDnsServer
              }
            } catch {
              & $showModalFunc "Error" "Failed to delete:`n$($_.Exception.Message)"
            }
          }
        }
        "🔄 Refresh" {
          if ($global:DNSConnectionStatus.IsConnected) {
            & $buildTreeFunc -dnsServer $global:DetectedDnsServer
          }
        }
      }
    }
  }
}

function Show-ExportDialog {
  Show-Modal "Export" "Export functionality not yet implemented"
}

function Show-ImportDialog {
  if (-not $global:DNSConnectionStatus.IsConnected) {
    Show-Modal "Error" "Not connected to DNS server"
    return
  }

  ## Capture functions
  $showModalFunc = ${function:Show-Modal}
  $debugLogFunc = ${function:Debug-Log}

  ## Get zones for dropdown
  $dnsZones = @()
  try {
    $dnsZones = Get-SafeDnsServerZone -DnsServerName $global:DetectedDnsServer
  } catch {
    Show-Modal "Error" "Failed to retrieve DNS zones"
    return
  }

  $importData = [PSCustomObject]@{
    FilePath        = ""
    TargetZone      = ""
    ConflictMode    = "Ask"  # Ask, Skip, Overwrite, SkipAll
    RecordsToImport = @()
  }

  $generalTab = @{
    Name = "Import Settings"
    Builder = {
      param($view, $data, $state)

      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Zone File Selection"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Zone File Path:" -FieldName 'txtFilePath' -State $state -Value "" -Width 60 -FieldX 25

      $y += 1
      $btnBrowse = [Terminal.Gui.Button]::new("Browse...")
      $btnBrowse.X = 25
      $btnBrowse.Y = $y
      $btnBrowse.add_Clicked({
        ## Simple file path input dialog
        $pathDlg = [Terminal.Gui.Dialog]::new()
        $pathDlg.Title = "Enter Zone File Path"
        $pathDlg.Width = 80
        $pathDlg.Height = 12

        $lblPath = [Terminal.Gui.Label]::new("Enter full path to BIND zone file:")
        $lblPath.X = 2; $lblPath.Y = 1
        $pathDlg.Add($lblPath)

        $txtPath = [Terminal.Gui.TextField]::new("")
        $txtPath.X = 2; $txtPath.Y = 3
        $txtPath.Width = 74
        $pathDlg.Add($txtPath)

        $btnOK = [Terminal.Gui.Button]::new("OK")
        $btnOK.add_Clicked({
          $state.txtFilePath.Text = $txtPath.Text
          [Terminal.Gui.Application]::RequestStop()
        })
        $pathDlg.AddButton($btnOK)

        $btnCancel = [Terminal.Gui.Button]::new("Cancel")
        $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
        $pathDlg.AddButton($btnCancel)

        [Terminal.Gui.Application]::Run($pathDlg)
      }.GetNewClosure())
      $view.Add($btnBrowse)
      $y += 2

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Target Zone" -SpaceBefore 0
      $lblZone = [Terminal.Gui.Label]::new("Import to Zone:")
      $lblZone.X = 2; $lblZone.Y = $y
      $view.Add($lblZone)

      $state.cmbZone = [Terminal.Gui.ComboBox]::new()
      $state.cmbZone.X = 25; $state.cmbZone.Y = $y; $state.cmbZone.Width = 40
      $state.cmbZone.SetSource(@($dnsZones | ForEach-Object { $_.ZoneName }))
      $state.cmbZone.SelectedItem = 0
      $view.Add($state.cmbZone)
      $y += 2

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Conflict Resolution" -SpaceBefore 0
      $lblConflict = [Terminal.Gui.Label]::new("When record exists:")
      $lblConflict.X = 2; $lblConflict.Y = $y
      $view.Add($lblConflict)

      $state.rbConflict = [Terminal.Gui.RadioGroup]::new()
      $state.rbConflict.X = 25
      $state.rbConflict.Y = $y
      $state.rbConflict.RadioLabels = @(
        "Ask for each conflict (safest)",
        "Skip existing records (keep current)",
        "Overwrite existing records (update all)",
        "⚠️  Skip all checks - DANGER! (force import)"
      )
      $state.rbConflict.SelectedItem = 0
      $view.Add($state.rbConflict)
      $y += 5

      $lblWarning = [Terminal.Gui.Label]::new("⚠️  Warning: 'Overwrite' and 'DANGER' modes will modify your DNS zone WITHOUT CONFIRMATION!")
      $lblWarning.X = 25; $lblWarning.Y = $y
      $view.Add($lblWarning)
    }
  }

  $previewTab = @{
    Name = "Preview"
    Builder = {
      param($view, $data, $state)

      $y = 1
      $lblPreview = [Terminal.Gui.Label]::new("Parsed Records Preview:")
      $lblPreview.X = 2; $lblPreview.Y = $y
      $view.Add($lblPreview)
      $y += 1

      $state.txtPreview = [Terminal.Gui.TextView]::new()
      $state.txtPreview.X = 2; $state.txtPreview.Y = $y
      $state.txtPreview.Width = [Terminal.Gui.Dim]::Fill(2)
      $state.txtPreview.Height = [Terminal.Gui.Dim]::Fill(3)
      $state.txtPreview.ReadOnly = $true
      $state.txtPreview.WordWrap = $false
      $state.txtPreview.Text = "(Click 'Parse File' to preview records)"
      $view.Add($state.txtPreview)

      $y = [Terminal.Gui.Pos]::AnchorEnd(2)
      $btnParse = [Terminal.Gui.Button]::new("Parse File")
      $btnParse.X = 2; $btnParse.Y = $y
      $btnParse.add_Clicked({
        $filePath = $state.txtFilePath.Text.ToString().Trim()
        if ([string]::IsNullOrWhiteSpace($filePath)) {
          & $showModalFunc "Error" "Please enter a zone file path"
          return
        }
        if (-not (Test-Path $filePath)) {
          & $showModalFunc "Error" "File not found: $filePath"
          return
        }

        try {
          $records = Parse-BINDZoneFile -FilePath $filePath
          $data.RecordsToImport = $records

          $preview = "Parsed $($records.Count) records:`n"
          $preview += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n"

          $recordTypes = $records | Group-Object RecordType | Sort-Object Name
          foreach ($group in $recordTypes) {
            $preview += "$($group.Name): $($group.Count) records`n"
          }
          $preview += "`nSample Records:`n"
          $preview += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n"

          $records | Select-Object -First 20 | ForEach-Object { $preview += "$($_.Name) [$($_.Type)] TTL=$($_.TTL) Data=$($_.Data)`n" }
          if ($records.Count -gt 20) { $preview += "`n... and $($records.Count - 20) more records" }
          $state.txtPreview.Text = $preview
          & $showModalFunc "Success" "Parsed $($records.Count) records from zone file"
        } catch {
          & $showModalFunc "Error" "Failed to parse zone file:`n$($_.Exception.Message)"
        }
      }.GetNewClosure())
      $view.Add($btnParse)
    }
  }

  $onOK = {
    param($data, $state)

    if ($data.RecordsToImport.Count -eq 0) {
      & $showModalFunc "Error" "No records to import. Click 'Parse File' first."
      return
    }

    $targetZone   = @($dnsZones)[$state.cmbZone.SelectedItem].ZoneName
    $conflictMode = @("Ask", "Skip", "Overwrite", "SkipAll")[$state.rbConflict.SelectedItem]
    & $debugLogFunc "Starting import: $($data.RecordsToImport.Count) records to $targetZone (mode: $conflictMode)" -Type "Insight"

    ## Import with conflict resolution
    Import-DNSRecordsWithConflictResolution -Records $data.RecordsToImport -ZoneName $targetZone -ServerName $global:DetectedDnsServer -ConflictMode $conflictMode
  }
  New-PropertiesDialog -Title "Import DNS Zone File" -Width 100 -Height 35 -Tabs @($generalTab, $previewTab) -Data $importData -OnOK $onOK -IncludeSearchTab $false
}

## ----------{ BIND Zone File Parser }----------

function Parse-BINDZoneFile {
  param(
    [Parameter(Mandatory)]
    [string]$FilePath
  )

  $records    = @()
  $origin     = ""
  $defaultTTL = 3600
  $content    = Get-Content -Path $FilePath -ErrorAction Stop

  foreach ($line in $content) {
    # Remove comments
    $line = $line -replace ';.*$', ''
    $line = $line.Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    # Handle directives
    if ($line -match '^\$ORIGIN\s+(.+)$') {
      $origin = $matches[1].Trim()
      continue
    }
    if ($line -match '^\$TTL\s+(\d+)$') {
      $defaultTTL = [int]$matches[1]
      continue
    }

    ## Parse record line
    ## Format: name [ttl] [class] type rdata
    $parts = $line -split '\s+', 5

    if ($parts.Count -lt 4) { continue }

    $name = $parts[0]
    $ttl = $defaultTTL
    $class = "IN"
    $type = ""
    $rdata = ""

    ## Detect format
    if ($parts[1] -match '^\d+$') {
      ## Has TTL
      $ttl = [int]$parts[1]
      if ($parts[2] -match '^(IN|CS|CH|HS)$') {
        $class = $parts[2]
        $type = $parts[3]
        $rdata = $parts[4]
      } else {
        $type = $parts[2]
        $rdata = $parts[3] + " " + $parts[4]
      }
    } elseif ($parts[1] -match '^(IN|CS|CH|HS)$') {
      ## Has class, no TTL
      $class = $parts[1]
      $type = $parts[2]
      $rdata = $parts[3] + " " + ($parts[4] ?? "")
    } else {
      ## No TTL, no class
      $type = $parts[1]
      $rdata = ($parts[2..4] -join " ").Trim()
    }
    ## Handle @ (zone apex)
    if ($name -eq '@') { $name = '' }
    ## Append origin if needed
    if ($name -notmatch '\.$' -and $origin) {
      $name = "$name.$origin"
    }
    $records += [PSCustomObject]@{
      Name       = $name
      TTL        = $ttl
      Class      = $class
      Type       = $type.ToUpper()
      Data       = $rdata.Trim()
      RecordType = $type.ToUpper()
    }
  }
  return $records
}

## ----------{ Import with Conflict Resolution }----------

function Import-DNSRecordsWithConflictResolution {
  param(
    [Parameter(Mandatory)]
    [array]$Records,
    [Parameter(Mandatory)]
    [string]$ZoneName,
    [Parameter(Mandatory)]
    [string]$ServerName,
    [string]$ConflictMode = "Ask"  ## Ask, Skip, Overwrite, SkipAll
  )

  $showModalFunc = ${function:Show-Modal}
  $debugLogFunc  = ${function:Debug-Log}

  $stats = @{
    Total       = $Records.Count
    Imported    = 0
    Skipped     = 0
    Overwritten = 0
    Failed      = 0
  }

  ## Get existing records if checking conflicts
  $existingRecords = @{}
  if ($ConflictMode -ne "SkipAll") {
    try {
      $existing = Get-DnsServerResourceRecord -ZoneName $ZoneName -ComputerName $ServerName -ErrorAction Stop
      foreach ($rec in $existing) {
        $key = "$($rec.HostName)|$($rec.RecordType)"
        if (-not $existingRecords.ContainsKey($key)) { $existingRecords[$key] = @() }
        $existingRecords[$key] += $rec
      }
    } catch {
      & $showModalFunc "Error" "Failed to retrieve existing records"
      return
    }
  }

  foreach ($record in $Records) {
    $key = "$($record.Name)|$($record.Type)"
    $exists = $existingRecords.ContainsKey($key)

    ## Handle conflict
    if ($exists -and $ConflictMode -ne "SkipAll") {
      if ($ConflictMode -eq "Skip") {
        $stats.Skipped++
        & $debugLogFunc "Skipped existing: $($record.Name) [$($record.Type)]" -Type "Insight"
        continue
      } elseif ($ConflictMode -eq "Ask") {
        ## Ask user
        $askDlg = [Terminal.Gui.Dialog]::new()
        $askDlg.Title = "Record Conflict"
        $askDlg.Width = 70
        $askDlg.Height = 14
        $msg = "Record already exists:`n`nName: $($record.Name)`nType: $($record.Type)`nNew Data: $($record.Data)`n`nOverwrite existing record?"
        $lblMsg = [Terminal.Gui.Label]::new($msg)
        $lblMsg.X = 2; $lblMsg.Y = 1
        $askDlg.Add($lblMsg)
        $userChoice = 0  ## 0=overwrite, 1=skip, 2=skip all remaining

        $btnOverwrite = [Terminal.Gui.Button]::new("Overwrite")
        $btnOverwrite.add_Clicked({ $userChoice = 0; [Terminal.Gui.Application]::RequestStop() })
        $askDlg.AddButton($btnOverwrite)

        $btnSkip = [Terminal.Gui.Button]::new("Skip")
        $btnSkip.add_Clicked({ $userChoice = 1; [Terminal.Gui.Application]::RequestStop() })
        $askDlg.AddButton($btnSkip)

        $btnSkipAll = [Terminal.Gui.Button]::new("Skip All")
        $btnSkipAll.add_Clicked({ $userChoice = 2; [Terminal.Gui.Application]::RequestStop() })
        $askDlg.AddButton($btnSkipAll)

        [Terminal.Gui.Application]::Run($askDlg)

        if ($userChoice -eq 1) {
          $stats.Skipped++
          continue
        } elseif ($userChoice -eq 2) {
          $ConflictMode = "Skip"
          $stats.Skipped++
          continue
        }
        ## else overwrite (fall through)
      }

      ## Overwrite mode - remove old record first
      if ($ConflictMode -eq "Overwrite" -or $ConflictMode -eq "Ask") {
        try {
          $oldRecord = $existingRecords[$key][0]
          Remove-DnsServerResourceRecord -InputObject $oldRecord -ZoneName $ZoneName -ComputerName $ServerName -Force -ErrorAction Stop
          $stats.Overwritten++
        } catch {
          & $debugLogFunc "Failed to remove old record: $_" -Type "Error"
          $stats.Failed++
          continue
        }
      }
    }

    ## Import record
    try {
      $params = @{
        ZoneName     = $ZoneName
        Name         = $record.Name
        ComputerName = $ServerName
        TimeToLive   = [TimeSpan]::FromSeconds($record.TTL)
        ErrorAction  = "Stop"
      }

      switch ($record.Type) {
        'A' {
          $params.A = $true
          $params.IPv4Address = $record.Data
        }
        'AAAA' {
          $params.AAAA = $true
          $params.IPv6Address = $record.Data
        }
        'CNAME' {
          $params.CName = $true
          $params.HostNameAlias = $record.Data
        }
        'MX' {
          $dataParts = $record.Data -split '\s+', 2
          $params.MX = $true
          $params.Preference = [uint16]$dataParts[0]
          $params.MailExchange = $dataParts[1]
        }
        'NS' {
          $params.NS = $true
          $params.NameServer = $record.Data
        }
        'PTR' {
          $params.Ptr = $true
          $params.PtrDomainName = $record.Data
        }
        'TXT' {
          $params.Txt = $true
          $params.DescriptiveText = $record.Data.Trim('"')
        }
        'SRV' {
          $dataParts = $record.Data -split '\s+', 4
          $params.Srv = $true
          $params.Priority = [uint16]$dataParts[0]
          $params.Weight = [uint16]$dataParts[1]
          $params.Port = [uint16]$dataParts[2]
          $params.DomainName = $dataParts[3]
        }
        default {
          & $debugLogFunc "Unsupported record type: $($record.Type)" -Type "Warning"
          $stats.Skipped++
          continue
        }
      }
      Add-DnsServerResourceRecord @params
      $stats.Imported++
    } catch {
      & $debugLogFunc "Failed to import $($record.Name) [$($record.Type)]: $_" -Type "Error"
      $stats.Failed++
    }
  }

  # Show summary
  $summary = "Import Complete!`n`n"
  $summary += "Total Records: $($stats.Total)`n"
  $summary += "Imported: $($stats.Imported)`n"
  $summary += "Overwritten: $($stats.Overwritten)`n"
  $summary += "Skipped: $($stats.Skipped)`n"
  $summary += "Failed: $($stats.Failed)"

  & $showModalFunc "Import Summary" $summary
  & $debugLogFunc "Import complete: $($stats.Imported) imported, $($stats.Skipped) skipped, $($stats.Failed) failed" -Type "Success"
}

## ----------{ DNS Dialog - Query DNS Server Directly }----------
function Show-DNSZoneProperties {
  param(
    [Parameter(Mandatory=$true)]
    [string]$ZoneName,
    [Parameter(Mandatory=$true)]
    [ValidateSet('Forward','Reverse')]
    [string]$ZoneType
  )

  try {
    $zone = Get-DnsServerZone -Name $ZoneName -ComputerName $global:DetectedDnsServer -ErrorAction Stop

    $generalTab = @{
      Name = "General"
      Builder = {
        param($view, $data, $state)

        $y = 1
        Add-SectionHeader -View $view -Y ([ref]$y) -Text "Zone Information"
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Zone Name:" -FieldName 'lblZoneName' -State $state -Value $data.ZoneName -IsTextField $false -FieldX 25
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Zone Type:" -FieldName 'lblZoneType' -State $state -Value $data.ZoneType -IsTextField $false -FieldX 25
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Status:" -FieldName 'lblStatus' -State $state -Value $data.ZoneStatus -IsTextField $false -FieldX 25
        if ($data.ReplicationScope) { Add-LabelAndField -View $view -Y ([ref]$y) -Label "Replication:" -FieldName 'lblRep' -State $state -Value $data.ReplicationScope -IsTextField $false -FieldX 25 }

        $y += 1
        Add-SectionHeader -View $view -Y ([ref]$y) -Text "Dynamic Update" -SpaceBefore 0
        $dynUpdate = if ($data.DynamicUpdate) { $data.DynamicUpdate.ToString() } else { "None" }
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Dynamic Update:" -FieldName 'lblDynamic' -State $state -Value $dynUpdate -IsTextField $false -FieldX 25
        $y += 1
        Add-SectionHeader -View $view -Y ([ref]$y) -Text "Security" -SpaceBefore 0
        $isSigned = if ($data.IsSigned) { "✓ DNSSEC Signed" } else { "✗ Not DNSSEC Signed" }
        $lbl = [Terminal.Gui.Label]::new($isSigned)
        $lbl.X = 4; $lbl.Y = $y
        $view.Add($lbl)
        $y += 2

        if ($data.ZoneFile) {
          Add-SectionHeader -View $view -Y ([ref]$y) -Text "Storage" -SpaceBefore 0
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Zone File:" -FieldName 'lblFile' -State $state -Value $data.ZoneFile -IsTextField $false -FieldX 25
        }
      }
    }

    $soaTab = @{
      Name = "SOA Record"
      Builder = {
        param($view, $data, $state)

        $y = 1
        Add-SectionHeader -View $view -Y ([ref]$y) -Text "Start of Authority (SOA) Record"

        try {
          $soa = Get-DnsServerResourceRecord -ZoneName $data.ZoneName -RRType SOA -ComputerName $global:DetectedDnsServer -ErrorAction Stop | Select-Object -First 1

          if ($soa) {
            Add-LabelAndField -View $view -Y ([ref]$y) -Label "Primary Server:" -FieldName 'txtPrimary' -State $state -Value $soa.RecordData.PrimaryServer -Width 50 -FieldX 25
            Add-LabelAndField -View $view -Y ([ref]$y) -Label "Responsible Person:" -FieldName 'txtResponsible' -State $state -Value $soa.RecordData.ResponsiblePerson -Width 50 -FieldX 25
            $y += 1
            Add-SectionHeader -View $view -Y ([ref]$y) -Text "Timers" -SpaceBefore 0
            Add-LabelAndField -View $view -Y ([ref]$y) -Label "Serial Number:" -FieldName 'txtSerial' -State $state -Value $soa.RecordData.SerialNumber.ToString() -Width 20 -FieldX 25 -ReadOnly $true
            Add-LabelAndField -View $view -Y ([ref]$y) -Label "Refresh Interval:" -FieldName 'txtRefresh' -State $state -Value $soa.RecordData.RefreshInterval.TotalSeconds.ToString() -Width 20 -FieldX 25
            Add-LabelAndField -View $view -Y ([ref]$y) -Label "Retry Interval:" -FieldName 'txtRetry' -State $state -Value $soa.RecordData.RetryDelay.TotalSeconds.ToString() -Width 20 -FieldX 25
            Add-LabelAndField -View $view -Y ([ref]$y) -Label "Expire Limit:" -FieldName 'txtExpire' -State $state -Value $soa.RecordData.ExpireLimit.TotalSeconds.ToString() -Width 20 -FieldX 25
            Add-LabelAndField -View $view -Y ([ref]$y) -Label "Minimum TTL:" -FieldName 'txtMinTTL' -State $state -Value $soa.RecordData.MinimumTimeToLive.TotalSeconds.ToString() -Width 20 -FieldX 25
          }
        }
        catch {
          $lbl = [Terminal.Gui.Label]::new("Unable to retrieve SOA record: $_")
          $lbl.X = 4; $lbl.Y = $y
          $view.Add($lbl)
        }
      }
    }

    $nsTab = @{
      Name = "Name Servers"
      Builder = {
        param($view, $data, $state)

        $y = 1
        Add-SectionHeader -View $view -Y ([ref]$y) -Text "Name Servers for this Zone"

        $state.lstNameServers = [Terminal.Gui.ListView]::new()
        $state.lstNameServers.X = 4
        $state.lstNameServers.Y = $y
        $state.lstNameServers.Width = [Terminal.Gui.Dim]::Fill(2)
        $state.lstNameServers.Height = 20

        try {
          $nsRecords = Get-DnsServerResourceRecord -ZoneName $data.ZoneName -RRType NS -ComputerName $global:DetectedDnsServer -ErrorAction Stop
          $nsList = $nsRecords | ForEach-Object { "$($_.HostName) → $($_.RecordData.NameServer)" }

          if ($nsList.Count -gt 0) {
            $state.lstNameServers.SetSource($nsList)
          } else {
            $state.lstNameServers.SetSource(@("(No name servers configured)"))
          }
        }
        catch {
          $state.lstNameServers.SetSource(@("Error loading name servers: $_"))
        }

        $view.Add($state.lstNameServers)
      }
    }

    $statsTab = @{
      Name = "Statistics"
      Builder = {
        param($view, $data, $state)

        $y = 1
        Add-SectionHeader -View $view -Y ([ref]$y) -Text "Zone Statistics"

        try {
          $records = Get-DnsServerResourceRecord -ZoneName $data.ZoneName -ComputerName $global:DetectedDnsServer -ErrorAction Stop
          $recordTypes = $records | Group-Object -Property RecordType | Sort-Object Count -Descending

          $lblTotal = [Terminal.Gui.Label]::new("Total Records: $($records.Count)")
          $lblTotal.X = 4; $lblTotal.Y = $y
          $view.Add($lblTotal)
          $y += 2

          $lblBreakdown = [Terminal.Gui.Label]::new("Breakdown by Record Type:")
          $lblBreakdown.X = 4; $lblBreakdown.Y = $y
          $view.Add($lblBreakdown)
          $y += 1

          foreach ($type in $recordTypes) {
            $lbl = [Terminal.Gui.Label]::new("  $($type.Name): $($type.Count)")
            $lbl.X = 6; $lbl.Y = $y
            $view.Add($lbl)
            $y += 1
          }
        }
        catch {
          $lbl = [Terminal.Gui.Label]::new("Unable to retrieve zone statistics: $_")
          $lbl.X = 4; $lbl.Y = $y
          $view.Add($lbl)
        }
      }
    }

    $onApply = {
      param($data, $state)
      Show-Modal "Info" "Zone property updates not yet implemented"
    }
    New-PropertiesDialog -Title "DNS Zone Properties - $ZoneName" -Width 100 -Height 40 -Tabs @($generalTab, $soaTab, $nsTab, $statsTab) -Data $zone -OnApply $onApply -IncludeSearchTab $false
  }
  catch {
    Debug-Log "Failed to load zone properties: $_" -Type "Error"
    Show-Modal "Error" "Failed to load zone properties: $_"
  }
}

## ----------{ Build DNS Zone Tree }----------
function Build-DNSZoneTree {
  param([string]$dnsServer)
  
  if (-not $dnsServer) { $dnsServer = $global:DetectedDnsServer }
  Debug-Log "Building DNS zone tree for server: $dnsServer" -Type "Insight"
  
  ## TreeView MUST already exist
  if ($null -eq $Script:tree) {
    throw "Build-DNSZoneTree failed: TreeView does not exist"
  }
  
  ## Clear existing objects safely
  try {
    $Script:tree.ClearObjects()
  } catch {
    Debug-Log "WARNING - ClearObjects failed: $_" -Type "Warning"
  }
  
  ## Get icons
  $serverIcon = Get-Icon 'Server'
  $forwardIcon = Get-Icon 'ForwardZone'
  $reverseIcon = Get-Icon 'ReverseZone'
  $zoneIcon = Get-Icon 'Zone'
  
  ## Create root node (DNS Server)
  $root = [Terminal.Gui.Trees.TreeNode]::new("$serverIcon DNS Server: $dnsServer")
  $root.Tag = @{ Type = 'Server'; Object = @{ Name = $dnsServer }; Name = $dnsServer }
  
  ## Fetch all zones
  Debug-Log "Fetching DNS zones..." -Type "Insight"
  $allZones = Get-SafeDnsServerZone -DnsServerName $dnsServer
  
  if ($null -eq $allZones) {
    $allZones = @()
    Debug-Log "No zones returned from server" -Type "Warning"
  }
  
  ## Split into forward and reverse
  $Script:ForwardZones = @($allZones | Where-Object { -not $_.IsReverse })
  $Script:ReverseZones = @($allZones | Where-Object { $_.IsReverse })
  Debug-Log "Found $($Script:ForwardZones.Count) forward zones, $($Script:ReverseZones.Count) reverse zones" -Type "Insight"
  
  ## Build forward zones container
  if ($Script:ForwardZones.Count -gt 0) {
    $forwardContainer = [Terminal.Gui.Trees.TreeNode]::new("$forwardIcon Forward Zones ($($Script:ForwardZones.Count))")
    $forwardContainer.Tag = @{ Type = 'Container'; Object = $null; Name = 'Forward Zones' }
    
    foreach ($zone in ($Script:ForwardZones | Sort-Object -Property ZoneName)) {
      ## Create zone node
      $zoneNodeText = "$zoneIcon $($zone.ZoneName)"
      if ($zone.ZoneType) { $zoneNodeText += " [$($zone.ZoneType)]" }
      
      $zoneNode = [Terminal.Gui.Trees.TreeNode]::new($zoneNodeText)
      $zoneNode.Tag = @{
        Type = 'ForwardZone'
        Object = $zone
        Name = $zone.ZoneName
      }
      
      ## Fetch and add records as children
      try {
        $records = Get-DnsServerResourceRecord -ZoneName $zone.ZoneName -ComputerName $dnsServer -ErrorAction Stop
        
        if ($records) {
          ## Group by record type for better organization
          $recordsByType = $records | Group-Object -Property RecordType | Sort-Object Name
          
          foreach ($typeGroup in $recordsByType) {
            $recordType = $typeGroup.Name
            $recordCount = $typeGroup.Count
            
            ## Get appropriate icon for record type
            $recordIcon = switch ($recordType) {
              'A'     { Get-Icon 'A_Record' }
              'AAAA'  { Get-Icon 'AAAA_Record' }
              'CNAME' { Get-Icon 'CNAME' }
              'MX'    { Get-Icon 'MX' }
              'NS'    { Get-Icon 'NS' }
              'PTR'   { Get-Icon 'PTR' }
              'TXT'   { Get-Icon 'TXT' }
              'SOA'   { Get-Icon 'SOA' }
              'SRV'   { Get-Icon 'SRV' }
              default { Get-Icon 'Record' }
            }
            
            ## Create type container node
            $typeNodeText = "$recordIcon $recordType Records ($recordCount)"
            $typeNode = [Terminal.Gui.Trees.TreeNode]::new($typeNodeText)
            $typeNode.Tag = @{ Type = 'RecordTypeContainer'; Object = $null; Name = "$recordType Records" }
            
            ## Add individual records under type
            foreach ($record in ($typeGroup.Group | Sort-Object HostName)) {
              ## Format record display
              $recordData = Format-RecordData -record $record
              $recordText = "$recordIcon $($record.HostName) → $recordData"
              
              $recordNode = [Terminal.Gui.Trees.TreeNode]::new($recordText)
              $recordNode.Tag = @{
                Type = 'Record'
                Object = $record
                Name = $record.HostName
                ZoneName = $zone.ZoneName
              }
              
              $typeNode.Children.Add($recordNode)
            }
            
            $zoneNode.Children.Add($typeNode)
          }
          
          Debug-Log "Added $($records.Count) records to zone $($zone.ZoneName)" -Type "Tracing"
        } else {
          ## Add empty indicator
          $emptyNode = [Terminal.Gui.Trees.TreeNode]::new("(No records)")
          $emptyNode.Tag = @{ Type = 'Empty'; Object = $null; Name = 'Empty' }
          $zoneNode.Children.Add($emptyNode)
        }
      } catch {
        Debug-Log "Failed to load records for zone $($zone.ZoneName): $_" -Type "Warning"
        $errorNode = [Terminal.Gui.Trees.TreeNode]::new("⚠️ Error loading records")
        $errorNode.Tag = @{ Type = 'Error'; Object = $null; Name = 'Error' }
        $zoneNode.Children.Add($errorNode)
      }
      
      $forwardContainer.Children.Add($zoneNode)
    }
    
    $root.Children.Add($forwardContainer)
    Debug-Log "Added Forward Zones container with $($Script:ForwardZones.Count) zones" -Type "Success"
  }
  
  ## Build reverse zones container
  if ($Script:ReverseZones.Count -gt 0) {
    $reverseContainer = [Terminal.Gui.Trees.TreeNode]::new("$reverseIcon Reverse Zones ($($Script:ReverseZones.Count))")
    $reverseContainer.Tag = @{ Type = 'Container'; Object = $null; Name = 'Reverse Zones' }
    
    foreach ($zone in ($Script:ReverseZones | Sort-Object -Property ZoneName)) {
      ## Create zone node
      $zoneNodeText = "$zoneIcon $($zone.ZoneName)"
      if ($zone.ZoneType) { $zoneNodeText += " [$($zone.ZoneType)]" }
      
      $zoneNode = [Terminal.Gui.Trees.TreeNode]::new($zoneNodeText)
      $zoneNode.Tag = @{
        Type = 'ReverseZone'
        Object = $zone
        Name = $zone.ZoneName
      }
      
      ## Fetch and add records as children
      try {
        $records = Get-DnsServerResourceRecord -ZoneName $zone.ZoneName -ComputerName $dnsServer -ErrorAction Stop
        
        if ($records) {
          ## Group by record type
          $recordsByType = $records | Group-Object -Property RecordType | Sort-Object Name
          
          foreach ($typeGroup in $recordsByType) {
            $recordType = $typeGroup.Name
            $recordCount = $typeGroup.Count
            
            ## Get appropriate icon
            $recordIcon = switch ($recordType) {
              'A'     { Get-Icon 'A_Record' }
              'AAAA'  { Get-Icon 'AAAA_Record' }
              'CNAME' { Get-Icon 'CNAME' }
              'MX'    { Get-Icon 'MX' }
              'NS'    { Get-Icon 'NS' }
              'PTR'   { Get-Icon 'PTR' }
              'TXT'   { Get-Icon 'TXT' }
              'SOA'   { Get-Icon 'SOA' }
              'SRV'   { Get-Icon 'SRV' }
              default { Get-Icon 'Record' }
            }
            
            ## Create type container
            $typeNodeText = "$recordIcon $recordType Records ($recordCount)"
            $typeNode = [Terminal.Gui.Trees.TreeNode]::new($typeNodeText)
            $typeNode.Tag = @{ Type = 'RecordTypeContainer'; Object = $null; Name = "$recordType Records" }
            
            ## Add individual records
            foreach ($record in ($typeGroup.Group | Sort-Object HostName)) {
              $recordData = Format-RecordData -record $record
              $recordText = "$recordIcon $($record.HostName) → $recordData"
              
              $recordNode = [Terminal.Gui.Trees.TreeNode]::new($recordText)
              $recordNode.Tag = @{
                Type = 'Record'
                Object = $record
                Name = $record.HostName
                ZoneName = $zone.ZoneName
              }
              
              $typeNode.Children.Add($recordNode)
            }
            
            $zoneNode.Children.Add($typeNode)
          }
          
          Debug-Log "Added $($records.Count) records to zone $($zone.ZoneName)" -Type "Tracing"
        } else {
          ## Add empty indicator
          $emptyNode = [Terminal.Gui.Trees.TreeNode]::new("(No records)")
          $emptyNode.Tag = @{ Type = 'Empty'; Object = $null; Name = 'Empty' }
          $zoneNode.Children.Add($emptyNode)
        }
      } catch {
        Debug-Log "Failed to load records for zone $($zone.ZoneName): $_" -Type "Warning"
        $errorNode = [Terminal.Gui.Trees.TreeNode]::new("⚠️ Error loading records")
        $errorNode.Tag = @{ Type = 'Error'; Object = $null; Name = 'Error' }
        $zoneNode.Children.Add($errorNode)
      }
      
      $reverseContainer.Children.Add($zoneNode)
    }
    
    $root.Children.Add($reverseContainer)
    Debug-Log "Added Reverse Zones container with $($Script:ReverseZones.Count) zones" -Type "Success"
  }
  
  ## Attach root to TreeView
  try {
    $Script:tree.AddObject($root)
    $Script:tree.SelectedObject = $root
    Debug-Log "Root node added to TreeView" -Type "Success"
  } catch {
    throw "Build-DNSZoneTree failed while attaching root to TreeView: $_"
  }
  
  Debug-Log "Build-DNSZoneTree completed successfully" -Type "Success"
  return $root
}

## ----------{ Production check functions }----------
## TODO: Obviously this needs changing pu for DNS and we can never cover all tols, so check for the standard ones
## on reach platform, dig, nslookup, etc etc ping and traceroute etc get grandfathered in.
## Samba and LDAP etc can go ;)
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

## ----------{ Change DC Dialog }----------
## TOOD: Make this a change DNS Server dialog instead
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

## ----------{ Clipboard Helper Functions }----------
function Copy-ToClipboard {
  param(
    [Parameter(Mandatory)]
    [ValidateSet('LDAPQuery', 'SearchResults')]
    [string]$ContentType,
    [Parameter()]
    [Terminal.Gui.TextView]$TextView
  )

  try {
    switch ($ContentType) {
      'LDAPQuery' {
        if (-not $TextView) {
          Show-Modal "Error" "TextView parameter required for LDAPQuery"
          return
        }
        $query = $TextView.Text.ToString().Trim()
        if (-not $query) {
          Show-Modal "Info" "No LDAP query to copy"
          return
        }
        Set-Clipboard -Value $query
        Debug-Log "Copied LDAP query to clipboard" -Type "Success"
        Show-Modal "Success" "LDAP query copied to clipboard"
      }
    'SearchResults' {
      if (-not $Script:lastSearchResults -or $Script:lastSearchResults.Count -eq 0) {
        Show-Modal "Info" "No search results to copy"
        return
      }
      ## Format results for clipboard
      $clipboardText = "# AD Search Results - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
      $clipboardText += "# Search Type: $($Script:lastSearchType)`n"
      $clipboardText += "# Total Results: $($Script:lastSearchResults.Count)`n"
      $clipboardText += "`n"
      ## Format 1: Simple list
      $clipboardText += "========== Simple List ==========`n"
      foreach ($obj in $Script:lastSearchResults) { $clipboardText += "$($obj.Name) [$($obj.Type)]`n" }
      ## Format 2: CSV
      $clipboardText += "`n========== CSV Format ==========`n"
      $clipboardText += "Name,Type"
      if ($Script:lastSearchResults[0].PSObject.Properties['DN']) { $clipboardText += ",DistinguishedName" }
      if ($Script:lastSearchResults[0].PSObject.Properties['Enabled']) { $clipboardText += ",Enabled" }
        $clipboardText += "`n"
        foreach ($obj in $Script:lastSearchResults) {
          $clipboardText += "`"$($obj.Name)`",`"$($obj.Type)`""
          if ($obj.PSObject.Properties['DN']) { $clipboardText += ",`"$($obj.DN)`"" }
          if ($obj.PSObject.Properties['Enabled']) { $clipboardText += ",`"$($obj.Enabled)`"" }
          $clipboardText += "`n"
        }
        ## Format 3: PowerShell array
        $clipboardText += "`n========== PowerShell Names ==========`n"
        $clipboardText += "@(`n"
        $names = $Script:lastSearchResults | ForEach-Object { '   "' + $_.Name + '"' }
        $clipboardText += ($names -join ",`n")
        $clipboardText += "`n)`n"
        Set-Clipboard -Value $clipboardText
        Debug-Log "Copied $($Script:lastSearchResults.Count) search results to clipboard" -Type "Success"
        Show-Modal "Success" "Copied $($Script:lastSearchResults.Count) results to clipboard`n`nFormats included:`n• Simple list`n• CSV`n• PowerShell array"
      }
    }
  } catch {
    Debug-Log "Failed to copy to clipboard: $($_.Exception.Message)" -Type "Problem"
    Show-Modal "Error" "Failed to copy to clipboard:`n$($_.Exception.Message)"
  }
}

function Paste-FromClipboard {
  param(
    [Parameter(Mandatory)]
    [ValidateSet('LDAPQuery')]
    [string]$ContentType,
    [Parameter()]
    [Terminal.Gui.TextView]$TextView
  )

  try {
    switch ($ContentType) {
     'LDAPQuery' {
        if (-not $TextView) {
          Show-Modal "Error" "TextView parameter required for LDAPQuery"
          return
        }
        $clipboardText = Get-Clipboard -Raw
        if (-not $clipboardText) {
          Show-Modal "Info" "Clipboard is empty"
          return
        }
      $TextView.Text = $clipboardText
      Debug-Log "Pasted LDAP query from clipboard" -Type "Success"
      [Terminal.Gui.Application]::Refresh()
      }
    }
  } catch {
    Debug-Log "Failed to paste from clipboard: $($_.Exception.Message)" -Type "Problem"
    Show-Modal "Error" "Failed to paste from clipboard:`n$($_.Exception.Message)"
  }
}

## ----------{ Common tab builder - re-usable across dialogs }----------
function Show-DnsTui {
  param(
    [string]$DNSServer = 'localhost',
    [string]$Zone = 'example.com',
    [ScriptBlock]$showPropertiesDialog,
    [ScriptBlock]$showModal,
    [ScriptBlock]$debugLog
  )

  ## -------------------------{ Fetch DNS Records }-------------------------
  $dnsRecords = @()
  try {
    $dnsRecords = Get-DnsServerResourceRecord -ComputerName $DNSServer -ZoneName $Zone -ErrorAction Stop
  } catch {
    & $showModal "Error" "Failed to fetch DNS records from $DNSServer for zone ${Zone}:`n$($_.Exception.Message)"
    return
  }

  ## -------------------------{ Transform Records for TUI }-------------------------
  $items = @()
  foreach ($rec in $dnsRecords | Sort-Object HostName) {
    $type = $rec.RecordType
    $name = $rec.HostName
    $data = switch ($type) {
      'A'     { $rec.RecordData.IPv4Address.ToString() }
      'AAAA'  { $rec.RecordData.IPv6Address.ToString() }
      'CNAME' { $rec.RecordData.HostNameAlias }
      'MX'    { "$($rec.RecordData.MailExchange) (Pref $($rec.RecordData.Preference))" }
      default { $rec.RecordData.ToString() }
    }

    ## Store a custom object with the raw record for the modal
    $items += [PSCustomObject]@{
      Name = $name
      Type = $type
      Data = $data
      Raw  = $rec
    }
  }

  if ($items.Count -eq 0) {
    & $showModal "No Records" "No DNS records found in zone $Zone"
    return
  }

  ## -------------------------{ Show TUI List }-------------------------
  $lst = [Terminal.Gui.ListView]::new()
  $lst.X = 0; $lst.Y = 0
  $lst.Width = [Terminal.Gui.Dim]::Fill()
  $lst.Height = [Terminal.Gui.Dim]::Fill()
  $loadingNode = [TreeNode]::new("Loading DNS zones...")
  $Script:tree.AddObject($loadingNode)
  $treeFrame.Add($Script:tree)

  ## -------------------------{ Key Handlers }-------------------------
  $lst.add_KeyPress({
    param($args)
    switch ($args.Key) {
      [Terminal.Gui.Key]::Enter {
        $sel = $items[$lst.SelectedItem]
        if ($null -ne $sel) {
          ## Build a detailed DNS property object for the modal
          $props = [PSCustomObject]@{
          ZoneName    = $sel.Raw.ZoneName
          HostName    = $sel.Raw.HostName
          RecordType  = $sel.Raw.RecordType
          RecordClass = $sel.Raw.RecordClass
          TTL         = $sel.Raw.TimeToLive
          Data        = switch ($sel.Raw.RecordType) {
            'A'     { $sel.Raw.RecordData.IPv4Address.ToString() }
            'AAAA'  { $sel.Raw.RecordData.IPv6Address.ToString() }
            'CNAME' { $sel.Raw.RecordData.HostNameAlias }
            'MX'    { "$($sel.Raw.RecordData.MailExchange) (Pref $($sel.Raw.RecordData.Preference))" }
            default { $sel.Raw.RecordData.ToString() }
          }
        }
        & $showPropertiesDialog $props
      }
      $args.Handled = $true
    }
    'e' {
      $sel = $items[$lst.SelectedItem]
      if ($null -ne $sel) {
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $filename = "dns_record_${Zone}_${timestamp}.csv"
        try {
          $sel.Raw | Select-Object HostName, RecordType, RecordClass, @{Name='Data';Expression={
            switch ($_.RecordType) {
              'A'     { $_.RecordData.IPv4Address.ToString() }
              'AAAA'  { $_.RecordData.IPv6Address.ToString() }
              'CNAME' { $_.RecordData.HostNameAlias }
              'MX'    { "$($_.RecordData.MailExchange) (Pref $($_.RecordData.Preference))" }
              default { $_.RecordData.ToString() }
            }
          }}, TimeToLive | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
          & $showModal "Export Complete" "DNS record exported to:`n$filename"
          & $debugLog "Exported DNS record $($sel.Name) to $filename" -Type "Success"
        } catch {
          & $showModal "Export Failed" "Failed to export DNS record:`n$($_.Exception.Message)"
        }
      }
      $args.Handled = $true
      }
    }
  })

  ## -------------------------{ Run TUI }-------------------------
  [Terminal.Gui.Application]::Init()
  $top = [Terminal.Gui.Application]::Top
  $win = [Terminal.Gui.Window]::new("DNS Zone: $Zone")
  $win.X = 0; $win.Y = 1
  $win.Width = [Terminal.Gui.Dim]::Fill()
  $win.Height = [Terminal.Gui.Dim]::Fill()
  $win.Add($lst)
  $top.Add($win)
  [Terminal.Gui.Application]::Run($top)
  [Terminal.Gui.Application]::Shutdown()
}


## DNS-TUI Batch Operations Module v1.0 - Select multiple objects and perform bulk actions
## ----------{ Global Selection State }----------
$Script:SelectedObjects = @()
$Script:SelectionMode   = $false

## --------------------------{ Core Functions }----------------------------
function Get-DNSServerDetection {
  $dnsDetection = @{
    IsLocalDNS = $false
    DnsServerName = $env:COMPUTERNAME
    DnsServiceRunning = $false
    AutoConnect = $false
  }

  try {
    $dnsService = Get-Service -Name DNS -ErrorAction SilentlyContinue
    if ($dnsService) {
      $dnsDetection.DnsServiceRunning = ($dnsService.Status -eq 'Running')
      $dnsDetection.IsLocalDNS        = $true
      $dnsDetection.AutoConnect       = $dnsDetection.DnsServiceRunning
      Debug-Log "DNS service detected on local machine: $($dnsService.Status)" Type "Insight"
    }
  } catch {
    Debug-Log "No local DNS service detected" -Type "Warning"
  }
  return $dnsDetection
}

function Invoke-SafeOperation {
  param(
    [scriptblock]$Operation,
    [string]$ErrorMessage = "Operation failed"
  )

  try {
    $startTime = Get-Date
    $result    = & $Operation
    $endTime   = Get-Date
    $duration  = ($endTime - $startTime).TotalMilliseconds

    $global:PerformanceCounters.OperationCount++
    $global:PerformanceCounters.LastOperationTime = $duration

    if ($global:PerformanceCounters.OperationCount -gt 0) {
      $global:PerformanceCounters.AverageResponseTime =
        (($global:PerformanceCounters.AverageResponseTime * ($global:PerformanceCounters.OperationCount - 1)) + $duration) /
        $global:PerformanceCounters.OperationCount
    }
    return $result
  } catch {
    $global:PerformanceCounters.ErrorCount++
    Debug-Log "$ErrorMessage : $_" -Type "Error"
    throw
  }
}

function Test-DNSServerConnection {
  param([string]$ServerName)

  if ([string]::IsNullOrWhiteSpace($ServerName)) { return $false }

  $now = Get-Date
  if ($global:DNSConnectionStatus.IsConnected -and
    $global:DNSConnectionStatus.ServerName -eq $ServerName -and
    $global:DNSConnectionStatus.LastChecked -and
    ($now - $global:DNSConnectionStatus.LastChecked).TotalSeconds -lt $global:DNSConnectionStatus.CacheValidSeconds) {
    return $true
  }

  try {
    $null = Get-DnsServerZone -ComputerName $ServerName -ErrorAction Stop | Select-Object -First 1
    $global:DNSConnectionStatus.IsConnected = $true
    $global:DNSConnectionStatus.ServerName = $ServerName
    $global:DNSConnectionStatus.LastChecked = $now
    return $true
  } catch {
    $global:DNSConnectionStatus.IsConnected = $false
    return $false
  }
}

function Get-SafeDnsServerZone {
  param(
    [string]$DnsServerName,
    [switch]$ForwardOnly,
    [switch]$ReverseOnly
  )
  try {
    $zones = Invoke-SafeOperation -Operation {
      Get-DnsServerZone -ComputerName $DnsServerName -ErrorAction Stop
    } -ErrorMessage "Failed to retrieve DNS zones"

    ## Ensure we have an array
    if ($null -eq $zones) { return @() }
    if ($zones -isnot [array]) { $zones = @($zones) }
    if ($ForwardOnly) { $zones = $zones | Where-Object { -not $_.IsReverseLookupZone } }
    if ($ReverseOnly) { $zones = $zones | Where-Object { $_.IsReverseLookupZone } }
    ## Ensure we still have data after filtering
    if ($null -eq $zones) { return @() }
    if ($zones -isnot [array]) { $zones = @($zones) }
    $zones = $zones | Select-Object ZoneName, ZoneType,
      @{Name='IsReverse'; Expression={$_.IsReverseLookupZone}},
      @{Name='RepScope'; Expression={$_.ReplicationScope}},
      @{Name='IsSigned'; Expression={$_.IsSigned}}
    return $zones
  } catch {
    Debug-Log "Error retrieving zones: $_" -Type "Error"
    return @()
  }
}

function Format-RecordData {
  param($record)

  try {
    switch ($record.RecordType) {
      "A"     { return $record.RecordData.IPv4Address.ToString() }
      "AAAA"  { return $record.RecordData.IPv6Address.ToString() }
      "CNAME" { return $record.RecordData.HostNameAlias }
      "MX"    { return "$($record.RecordData.Preference) $($record.RecordData.MailExchange)" }
      "PTR"   { return $record.RecordData.PtrDomainName }
      "TXT"   { return $record.RecordData.DescriptiveText -join " " }
      "SRV"   { return "$($record.RecordData.Priority) $($record.RecordData.Weight) $($record.RecordData.Port) $($record.RecordData.DomainName)" }
      "NS"    { return $record.RecordData.NameServer }
      "SOA"   { return $record.RecordData.PrimaryServer }
      default { return $record.RecordData.ToString() }
    }
  } catch {
    return "N/A"
  }
}

## --------------------{ DNS View functions }--------------------
function Show-ConnectDialog {
  $dialog = [Terminal.Gui.Dialog]@{
    Title = "Connect to DNS Server"
    Width = 60
    Height = 12
  }

  $label = [Terminal.Gui.Label]@{
    X = 1
    Y = 1
    Text = "DNS Server:"
  }
  $dialog.Add($label)

  $serverField = [Terminal.Gui.TextField]@{
    X = 15
    Y = 1
    Width = 40
    Text = $global:DetectedDnsServer
  }
  $dialog.Add($serverField)

  $autoDetectLabel = [Terminal.Gui.Label]@{
    X = 1
    Y = 3
    Text = "Auto-detected: $($global:DetectedDnsServer)"
  }
  $dialog.Add($autoDetectLabel)

  $btnConnect = [Terminal.Gui.Button]@{
    X = 15
    Y = 6
    Text = "Connect"
    IsDefault = $true
  }
  $btnConnect.add_Clicked({
    $serverName = $serverField.Text.ToString()
    if (Connect-ToDNSServer -ServerName $serverName) { [Terminal.Gui.Application]::RequestStop() }
  })
  $dialog.Add($btnConnect)

  $btnCancel = [Terminal.Gui.Button]@{
    X = 30
    Y = 6
    Text = "Cancel"
  }
  $btnCancel.add_Clicked({
    [Terminal.Gui.Application]::RequestStop()
  })
  $dialog.Add($btnCancel)
  [Terminal.Gui.Application]::Run($dialog)
}

function Connect-ToDNSServer {
  param([string]$ServerName)

  if ([string]::IsNullOrWhiteSpace($ServerName)) {
    Show-Modal "Error" "Please enter a DNS server name"
    return $false
  }

  $loading = Show-LoadingDialog -Message "Connecting to $ServerName..."

  try {
    Start-Sleep -Milliseconds 500  # Brief delay for loading animation
    if (Test-DNSServerConnection -ServerName $ServerName) {
      $global:DetectedDnsServer = $ServerName
      Debug-Log "Connected to DNS server: $ServerName" -Type "Success"
      if ($loading.Timer) { $loading.Timer.Dispose() }
      [Terminal.Gui.Application]::RequestStop()
      Show-Dashboard
      return $true
    } else {
      throw "Cannot connect to DNS server"
    }
  } catch {
    if ($loading.Timer) { $loading.Timer.Dispose() }
    [Terminal.Gui.Application]::RequestStop()
    Debug-Log "Connection failed: $_" -Type "Error"
    Show-Modal "Connection Failed" "Failed to connect to DNS server: $ServerName`n`n$_"
    return $false
  }
}

function Reload-DNSZone {
  param(
    [string]$ZoneName,
    [string]$ZoneType
  )

  if ([string]::IsNullOrWhiteSpace($ZoneName)) {
    Show-Modal "Error" "Zone name is required"
    return $false
  }
  $result = Show-Modal "Reload Zone", "Reload zone '$ZoneName' from source?`n`nThis will:`n- Secondary zones: Transfer from master`n- File-backed zones: Reload from file" -YesNo
  if ($result -ne 0) { return $false }
    try {
      Debug-Log "Reloading zone: $ZoneName (Type: $ZoneType)" Type "Insight"
      if ($ZoneType -eq "Secondary") {
        ## Force zone transfer from master
        Sync-DnsServerZone -Name $ZoneName -ComputerName $global:DetectedDnsServer -Force -ErrorAction Stop
        Debug-Log "Zone transfer completed for: $ZoneName" -Type "Success"
        Show-Modal "Success" "Zone '$ZoneName' reloaded from master server"
      } else {
        ## Reload from file for primary zones
        Update-DnsServerZone -Name $ZoneName -ComputerName $global:DetectedDnsServer -Force -ErrorAction Stop
        Debug-Log "Zone reloaded from file: $ZoneName" -Type "Success"
        Show-Modal "Success" "Zone '$ZoneName' reloaded from file"
      }
      return $true
    } catch {
      Debug-Log "Failed to reload zone $ZoneName : $_" -Type "Error"
      Show-Modal "Error" "Failed to reload zone: $_"
      return $false
  }
}

## ----------{ DNS Record Properties Dialog - Full Edit Suite }----------

function Show-DNSRecordProperties {
  param(
    [Parameter(Mandatory)]
    $Record,
    [Parameter(Mandatory)]
    [string]$ZoneName,
    [Parameter(Mandatory)]
    [string]$ServerName
  )

  Debug-Log "Opening properties for DNS record: $($Record.HostName) [$($Record.RecordType)]" -Type "Tracing"

  ## Build tabs based on record type
  $tabs = @()

  ## General tab - varies by record type
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $data, $state)

      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Record Information"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Hostname:" -FieldName 'txtHostName' -State $state -Value ($data.HostName ?? "@") -Width 50 -FieldX 25
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Record Type:" -FieldName 'lblRecordType' -State $state -Value $data.RecordType -IsTextField $false -FieldX 25
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Zone:" -FieldName 'lblZone' -State $state -Value $ZoneName -IsTextField $false -FieldX 25
      $y += 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Record Data" -SpaceBefore 0

      ## Type-specific fields
      switch ($data.RecordType) {
        'A' {
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "IPv4 Address:" -FieldName 'txtIPv4' -State $state -Value ($data.RecordData.IPv4Address ?? "") -Width 30 -FieldX 25
        }
        'AAAA' {
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "IPv6 Address:" -FieldName 'txtIPv6' -State $state -Value ($data.RecordData.IPv6Address ?? "") -Width 50 -FieldX 25
        }
        'CNAME' {
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Alias Target:" -FieldName 'txtAlias' -State $state -Value ($data.RecordData.HostNameAlias ?? "") -Width 50 -FieldX 25
        }
        'MX' {
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Mail Server:" -FieldName 'txtMailServer' -State $state -Value ($data.RecordData.MailExchange ?? "") -Width 50 -FieldX 25
          $y += 1
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Priority:" -FieldName 'txtPriority' -State $state -Value ($data.RecordData.Preference ?? "10") -Width 10 -FieldX 25
        }
        'NS' {
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Name Server:" -FieldName 'txtNameServer' -State $state -Value ($data.RecordData.NameServer ?? "") -Width 50 -FieldX 25
        }
        'PTR' {
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "PTR Domain:" -FieldName 'txtPtrDomain' -State $state -Value ($data.RecordData.PtrDomainName ?? "") -Width 50 -FieldX 25
        }
        'SRV' {
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Target Host:" -FieldName 'txtSrvTarget' -State $state -Value ($data.RecordData.DomainName ?? "") -Width 50 -FieldX 25
          $y += 1
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Port:" -FieldName 'txtSrvPort' -State $state -Value ($data.RecordData.Port ?? "0") -Width 10 -FieldX 25
          $y += 1
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Priority:" -FieldName 'txtSrvPriority' -State $state -Value ($data.RecordData.Priority ?? "0") -Width 10 -FieldX 25
          $y += 1
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Weight:" -FieldName 'txtSrvWeight' -State $state -Value ($data.RecordData.Weight ?? "0") -Width 10 -FieldX 25
        }
        'TXT' {
          $lblTxt = [Terminal.Gui.Label]::new("Text Data:")
          $lblTxt.X = 2; $lblTxt.Y = $y
          $view.Add($lblTxt)
          $y += 1

          $state.txtTextData = [Terminal.Gui.TextView]::new()
          $state.txtTextData.X = 2; $state.txtTextData.Y = $y
          $state.txtTextData.Width = [Terminal.Gui.Dim]::Fill(2)
          $state.txtTextData.Height = 8
          $state.txtTextData.Text = ($data.RecordData.DescriptiveText -join "`n")
          $view.Add($state.txtTextData)
        }
      }
    }
  }

  # Advanced tab
  $advancedTab = @{
    Name = "Advanced"
    Builder = {
      param($view, $data, $state)

      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Time To Live (TTL)"

      $ttlSeconds = if ($data.TimeToLive) { $data.TimeToLive.TotalSeconds } else { 3600 }
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "TTL (seconds):" -FieldName 'txtTTL' -State $state -Value $ttlSeconds.ToString() -Width 20 -FieldX 25

      $y += 1
      $lblHelp = [Terminal.Gui.Label]::new("Common values: 300 (5 min), 3600 (1 hour), 86400 (1 day)")
      $lblHelp.X = 25; $lblHelp.Y = $y
      $view.Add($lblHelp)

      $y += 2
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Timestamps" -SpaceBefore 0

      if ($data.Timestamp) {
        $lbl = [Terminal.Gui.Label]::new("Timestamp: $($data.Timestamp.ToString('yyyy-MM-dd HH:mm:ss'))")
        $lbl.X = 4; $lbl.Y = $y
        $view.Add($lbl)
        $y += 1
      }

      if ($data.TimeToLive) {
        $lbl = [Terminal.Gui.Label]::new("Current TTL: $($data.TimeToLive)")
        $lbl.X = 4; $lbl.Y = $y
        $view.Add($lbl)
        $y += 1
      }

      $y += 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Record Options" -SpaceBefore 0

      $state.chkAging = [Terminal.Gui.CheckBox]::new("Enable aging for this record")
      $state.chkAging.X = 4; $state.chkAging.Y = $y
      $state.chkAging.Checked = $false  # Would need to check DnsServer settings
      $view.Add($state.chkAging)
      $y += 2

      $lblInfo = [Terminal.Gui.Label]::new("Note: Some advanced options require direct DNS server access")
      $lblInfo.X = 4; $lblInfo.Y = $y
      $view.Add($lblInfo)
    }
  }

  $tabs += $generalTab
  $tabs += $advancedTab

  # Apply logic
  $onApply = {
    param($data, $state)

    $showModalFunc = ${function:Show-Modal}
    $debugLogFunc = ${function:Debug-Log}

    try {
      & $debugLogFunc "Applying changes to DNS record..." -Type "Tracing"

      # Build parameters based on record type
      $updateParams = @{
        ZoneName = $ZoneName
        ComputerName = $ServerName
        ErrorAction = "Stop"
      }

      # Get old record to remove
      $oldRecord = Get-DnsServerResourceRecord -ZoneName $ZoneName -Name $data.HostName -RRType $data.RecordType -ComputerName $ServerName -ErrorAction Stop | Select-Object -First 1

      if (-not $oldRecord) {
        & $showModalFunc "Error" "Could not find original record to update"
        return
      }

      # Clone the record
      $newRecord = $oldRecord.Clone()

      # Update based on type
      switch ($data.RecordType) {
        'A' {
          $newIP = $state.txtIPv4.Text.ToString().Trim()
          if ($newIP -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
            $newRecord.RecordData.IPv4Address = [System.Net.IPAddress]::Parse($newIP)
          } else {
            & $showModalFunc "Error" "Invalid IPv4 address format"
            return
          }
        }
        'AAAA' {
          $newIP = $state.txtIPv6.Text.ToString().Trim()
          if (-not [string]::IsNullOrWhiteSpace($newIP)) {
            $newRecord.RecordData.IPv6Address = [System.Net.IPAddress]::Parse($newIP)
          }
        }
        'CNAME' {
          $newAlias = $state.txtAlias.Text.ToString().Trim()
          if (-not [string]::IsNullOrWhiteSpace($newAlias)) {
            $newRecord.RecordData.HostNameAlias = $newAlias
          }
        }
        'MX' {
          $newMailServer = $state.txtMailServer.Text.ToString().Trim()
          $newPriority = $state.txtPriority.Text.ToString().Trim()
          if (-not [string]::IsNullOrWhiteSpace($newMailServer)) {
            $newRecord.RecordData.MailExchange = $newMailServer
          }
          if ($newPriority -match '^\d+$') {
            $newRecord.RecordData.Preference = [uint16]$newPriority
          }
        }
        'NS' {
          $newNS = $state.txtNameServer.Text.ToString().Trim()
          if (-not [string]::IsNullOrWhiteSpace($newNS)) {
            $newRecord.RecordData.NameServer = $newNS
          }
        }
        'PTR' {
          $newPtr = $state.txtPtrDomain.Text.ToString().Trim()
          if (-not [string]::IsNullOrWhiteSpace($newPtr)) {
            $newRecord.RecordData.PtrDomainName = $newPtr
          }
        }
        'SRV' {
          $newTarget = $state.txtSrvTarget.Text.ToString().Trim()
          $newPort = $state.txtSrvPort.Text.ToString().Trim()
          $newPriority = $state.txtSrvPriority.Text.ToString().Trim()
          $newWeight = $state.txtSrvWeight.Text.ToString().Trim()

          if (-not [string]::IsNullOrWhiteSpace($newTarget)) { $newRecord.RecordData.DomainName = $newTarget }
          if ($newPort -match '^\d+$') { $newRecord.RecordData.Port = [uint16]$newPort }
          if ($newPriority -match '^\d+$') { $newRecord.RecordData.Priority = [uint16]$newPriority }
          if ($newWeight -match '^\d+$') { $newRecord.RecordData.Weight = [uint16]$newWeight }
        }
        'TXT' {
          $newText = $state.txtTextData.Text.ToString().Trim()
          if (-not [string]::IsNullOrWhiteSpace($newText)) {
            $newRecord.RecordData.DescriptiveText = $newText
          }
        }
      }

      # Update TTL if changed
      $newTTL = $state.txtTTL.Text.ToString().Trim()
      if ($newTTL -match '^\d+$') {
        $newRecord.TimeToLive = [TimeSpan]::FromSeconds([int]$newTTL)
      }

      # Remove old and add new
      Remove-DnsServerResourceRecord -InputObject $oldRecord @updateParams -Force
      Add-DnsServerResourceRecord -InputObject $newRecord @updateParams

      & $debugLogFunc "DNS record updated successfully" -Type "Success"
      & $showModalFunc "Success" "DNS record updated successfully!"

    } catch {
      & $debugLogFunc "Failed to update DNS record: $($_.Exception.Message)" -Type "Error"
      & $showModalFunc "Error" "Failed to update DNS record:`n`n$($_.Exception.Message)"
    }
  }
  New-PropertiesDialog -Title "DNS Record Properties - $($Record.HostName) [$($Record.RecordType)]" -Width 100 -Height 35 -Tabs $tabs -Data $Record -OnApply $onApply -IncludeSearchTab $false
}

## ----------{ Create DNS Record Dialog }----------

function Show-CreateDNSRecordDialog {
  param(
    [Parameter(Mandatory)]
    [string]$ZoneName,
    [Parameter(Mandatory)]
    [string]$ServerName
  )

  Debug-Log "Opening create record dialog for zone: $ZoneName" -Type "Tracing"

  $recordData = [PSCustomObject]@{
    HostName    = ""
    RecordType  = "A"
    IPv4Address = ""
    IPv6Address = ""
    Alias       = ""
    MailServer  = ""
    Priority    = "10"
    NameServer  = ""
    PtrDomain   = ""
    SrvTarget   = ""
    SrvPort     = "0"
    SrvPriority = "0"
    SrvWeight   = "0"
    TextData    = ""
    TTL         = "3600"
  }

  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $data, $state)

      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Record Configuration"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Hostname:" -FieldName 'txtHostName' -State $state -Value "" -Width 50 -FieldX 25

      $y += 1
      $lblHelp = [Terminal.Gui.Label]::new("(Use @ for zone apex, or enter subdomain name)")
      $lblHelp.X = 25; $lblHelp.Y = $y
      $view.Add($lblHelp)
      $y += 2

      $lblType = [Terminal.Gui.Label]::new("Record Type:")
      $lblType.X = 2; $lblType.Y = $y
      $view.Add($lblType)

      $state.cmbRecordType = [Terminal.Gui.ComboBox]::new()
      $state.cmbRecordType.X = 25; $state.cmbRecordType.Y = $y; $state.cmbRecordType.Width = 20
      $state.cmbRecordType.SetSource(@("A", "AAAA", "CNAME", "MX", "NS", "PTR", "SRV", "TXT"))
      $state.cmbRecordType.SelectedItem = 0
      $view.Add($state.cmbRecordType)
      $y += 2

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Record Data" -SpaceBefore 0
      ## A Record
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "IPv4 Address:" -FieldName 'txtIPv4' -State $state -Value "" -Width 30 -FieldX 25
      ## AAAA Record
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "IPv6 Address:" -FieldName 'txtIPv6' -State $state -Value "" -Width 50 -FieldX 25
      ## CNAME Record
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Alias Target:" -FieldName 'txtAlias' -State $state -Value "" -Width 50 -FieldX 25
      ## MX Record
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Mail Server:" -FieldName 'txtMailServer' -State $state -Value "" -Width 50 -FieldX 25
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "MX Priority:" -FieldName 'txtMxPriority' -State $state -Value "10" -Width 10 -FieldX 25
      ## NS Record
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Name Server:" -FieldName 'txtNameServer' -State $state -Value "" -Width 50 -FieldX 25
      ## PTR Record
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "PTR Domain:" -FieldName 'txtPtrDomain' -State $state -Value "" -Width 50 -FieldX 25
      ## SRV Record
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "SRV Target:" -FieldName 'txtSrvTarget' -State $state -Value "" -Width 50 -FieldX 25
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "SRV Port:" -FieldName 'txtSrvPort' -State $state -Value "0" -Width 10 -FieldX 25
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "SRV Priority:" -FieldName 'txtSrvPriority' -State $state -Value "0" -Width 10 -FieldX 25
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "SRV Weight:" -FieldName 'txtSrvWeight' -State $state -Value "0" -Width 10 -FieldX 25
      ## TXT Record
      $y += 1
      $lblTxt = [Terminal.Gui.Label]::new("TXT Data:")
      $lblTxt.X = 2; $lblTxt.Y = $y
      $view.Add($lblTxt)
      $y += 1
      $state.txtTextData = [Terminal.Gui.TextView]::new()
      $state.txtTextData.X = 25; $state.txtTextData.Y = $y
      $state.txtTextData.Width = 50
      $state.txtTextData.Height = 4
      $view.Add($state.txtTextData)
    }
  }

  $advancedTab = @{
    Name = "Advanced"
    Builder = {
      param($view, $data, $state)

      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Time To Live (TTL)"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "TTL (seconds):" -FieldName 'txtTTL' -State $state -Value "3600" -Width 20 -FieldX 25

      $y += 1
      $lblHelp = [Terminal.Gui.Label]::new("Common values: 300 (5 min), 3600 (1 hour), 86400 (1 day)")
      $lblHelp.X = 25; $lblHelp.Y = $y
      $view.Add($lblHelp)

      $y += 2
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Options" -SpaceBefore 0
      $state.chkCreatePtr = [Terminal.Gui.CheckBox]::new("Create associated PTR record (for A/AAAA records)")
      $state.chkCreatePtr.X = 4; $state.chkCreatePtr.Y = $y
      $view.Add($state.chkCreatePtr)
    }
  }

  $onOK = {
    param($data, $state)

    $showModalFunc = ${function:Show-Modal}
    $debugLogFunc  = ${function:Debug-Log}
    $hostName      = $state.txtHostName.Text.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($hostName)) {
      & $showModalFunc "Error" "Please enter a hostname"
      return
    }

    $recordType = @("A", "AAAA", "CNAME", "MX", "NS", "PTR", "SRV", "TXT")[$state.cmbRecordType.SelectedItem]
    $ttlSeconds = $state.txtTTL.Text.ToString().Trim()

    if ($ttlSeconds -notmatch '^\d+$') {
      & $showModalFunc "Error" "TTL must be a number (seconds)"
      return
    }

    try {
      $params = @{
        ZoneName     = $ZoneName
        Name         = $hostName
        ComputerName = $ServerName
        TimeToLive   = [TimeSpan]::FromSeconds([int]$ttlSeconds)
        ErrorAction  = "Stop"
      }

      switch ($recordType) {
        'A' {
          $ipv4 = $state.txtIPv4.Text.ToString().Trim()
          if ($ipv4 -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
            & $showModalFunc "Error" "Invalid IPv4 address"
            return
          }
          $params.A = $true
          $params.IPv4Address = $ipv4
          $params.CreatePtr = $state.chkCreatePtr.Checked
        }
        'AAAA' {
          $ipv6 = $state.txtIPv6.Text.ToString().Trim()
          if ([string]::IsNullOrWhiteSpace($ipv6)) {
            & $showModalFunc "Error" "Please enter an IPv6 address"
            return
          }
          $params.AAAA = $true
          $params.IPv6Address = $ipv6
          $params.CreatePtr = $state.chkCreatePtr.Checked
        }
        'CNAME' {
          $alias = $state.txtAlias.Text.ToString().Trim()
          if ([string]::IsNullOrWhiteSpace($alias)) {
            & $showModalFunc "Error" "Please enter an alias target"
            return
          }
          $params.CName = $true
          $params.HostNameAlias = $alias
        }
        'MX' {
          $mailServer = $state.txtMailServer.Text.ToString().Trim()
          $priority = $state.txtMxPriority.Text.ToString().Trim()
          if ([string]::IsNullOrWhiteSpace($mailServer)) {
            & $showModalFunc "Error" "Please enter a mail server"
            return
          }
          $params.MX = $true
          $params.MailExchange = $mailServer
          $params.Preference = [uint16]$priority
        }
        'NS' {
          $ns = $state.txtNameServer.Text.ToString().Trim()
          if ([string]::IsNullOrWhiteSpace($ns)) {
            & $showModalFunc "Error" "Please enter a name server"
            return
          }
          $params.NS = $true
          $params.NameServer = $ns
        }
        'PTR' {
          $ptr = $state.txtPtrDomain.Text.ToString().Trim()
          if ([string]::IsNullOrWhiteSpace($ptr)) {
            & $showModalFunc "Error" "Please enter a PTR domain name"
            return
          }
          $params.Ptr = $true
          $params.PtrDomainName = $ptr
        }
        'SRV' {
          $target = $state.txtSrvTarget.Text.ToString().Trim()
          $port = $state.txtSrvPort.Text.ToString().Trim()
          $priority = $state.txtSrvPriority.Text.ToString().Trim()
          $weight = $state.txtSrvWeight.Text.ToString().Trim()

          if ([string]::IsNullOrWhiteSpace($target)) {
            & $showModalFunc "Error" "Please enter a target host"
            return
          }

          $params.Srv = $true
          $params.DomainName = $target
          $params.Port = [uint16]$port
          $params.Priority = [uint16]$priority
          $params.Weight = [uint16]$weight
        }
        'TXT' {
          $txtData = $state.txtTextData.Text.ToString().Trim()
          if ([string]::IsNullOrWhiteSpace($txtData)) {
            & $showModalFunc "Error" "Please enter text data"
            return
          }
          $params.Txt = $true
          $params.DescriptiveText = $txtData
        }
      }

      Add-DnsServerResourceRecord @params
      & $debugLogFunc "DNS record created: $recordType $hostName in zone $ZoneName" -Type "Success"
      & $showModalFunc "Success" "DNS record created successfully!"
    } catch {
      & $debugLogFunc "Failed to create DNS record: $($_.Exception.Message)" -Type "Error"
      & $showModalFunc "Error" "Failed to create DNS record:`n`n$($_.Exception.Message)"
    }
  }
  New-PropertiesDialog -Title "Create DNS Record - $ZoneName" -Width 100 -Height 40 -Tabs @($generalTab, $advancedTab) -Data $recordData -OnOK $onOK -IncludeSearchTab $false
}

## ----------{ Updated Show-DNSDialog with Better Layout }----------

function Show-DNSDialog {
  param([string]$ServerName)

  if (-not $ServerName) { $ServerName = $global:DetectedDnsServer }
  if (-not $global:DNSConnectionStatus.IsConnected) {
    Show-Modal "Error" "Not connected to DNS server"
    return
  }
  Debug-Log "Opening DNS viewer for server: $ServerName" -Type "Tracing"

  ## Capture functions for closures
  $showModalFunc = ${function:Show-Modal}
  $showCreateRecordFunc = ${function:Show-CreateDNSRecordDialog}
  $showRecordPropsFunc = ${function:Show-DNSRecordProperties}
  $showZonePropsFunc = ${function:Show-DNSZoneProperties}
  $showDNSDialogFunc = ${function:Show-DNSDialog}
  $showImportFunc = ${function:Show-ImportDialog}
  $debugLogFunc = ${function:Debug-Log}

  ## Get DNS data from server
  $dnsZones = @()

  try {
    $dnsZones = Get-SafeDnsServerZone -DnsServerName $ServerName
    Debug-Log "Found $($dnsZones.Count) zones" -Type "Insight"
  } catch {
    Debug-Log "Failed to get zones: $_" -Type "Error"
  }

  # Create dialog
  $dialog = [Terminal.Gui.Dialog]::new()
  $dialog.Title = "DNS Zone & Record Manager - $ServerName"
  $dialog.Width = 120
  $dialog.Height = 40

  $y = 1

  # Summary label (left side)
  $lblSummary = [Terminal.Gui.Label]::new("Zones: $($dnsZones.Count)  |  Select zone → Create/Edit/Delete")
  $lblSummary.X = 2
  $lblSummary.Y = $y
  $dialog.Add($lblSummary)

  # Search label (right side)
  $lblSearch = [Terminal.Gui.Label]::new("Search Records:")
  $lblSearch.X = 70
  $lblSearch.Y = $y
  $dialog.Add($lblSearch)

  # Search textbox
  $txtSearch = [Terminal.Gui.TextField]::new("")
  $txtSearch.X = 87
  $txtSearch.Y = $y
  $txtSearch.Width = 30
  $dialog.Add($txtSearch)

  $y += 2

  ## ========== LEFT: ZONE DROPDOWN ==========
  $lblZones = [Terminal.Gui.Label]::new("DNS Zones:")
  $lblZones.X = 2
  $lblZones.Y = $y
  $dialog.Add($lblZones)
  $y += 1

  $cmbZones = [Terminal.Gui.ComboBox]::new()
  $cmbZones.X = 2
  $cmbZones.Y = $y
  $cmbZones.Width = 40
  $cmbZones.Height = 5

  $zoneNames = if ($dnsZones.Count -gt 0) {
    @($dnsZones | ForEach-Object { "$($_.ZoneName) [$($_.ZoneType)]" })
  } else {
    @("(No zones)")
  }
  $cmbZones.SetSource($zoneNames)
  $dialog.Add($cmbZones)

  ## ========== RIGHT: DETAILS BOX (TOP) ==========
  $lblDetails = [Terminal.Gui.Label]::new("Record Details:")
  $lblDetails.X = 44
  $lblDetails.Y = 3
  $dialog.Add($lblDetails)

  $txtDetails = [Terminal.Gui.TextView]::new()
  $txtDetails.X = 44
  $txtDetails.Y = 4
  $txtDetails.Width = [Terminal.Gui.Dim]::Fill(2)
  $txtDetails.Height = 7
  $txtDetails.ReadOnly = $true
  $txtDetails.WordWrap = $false
  $txtDetails.Text = "(Select a record to view details)"
  $dialog.Add($txtDetails)

  ## ========== BOTTOM: RECORDS LIST (BIG IN FRAME) ==========
  $recordsFrame = [Terminal.Gui.FrameView]::new()
  $recordsFrame.Title = "DNS Records"
  $recordsFrame.X = 2
  $recordsFrame.Y = 12
  $recordsFrame.Width = [Terminal.Gui.Dim]::Fill(2)
  $recordsFrame.Height = 15

  $lstRecords = [Terminal.Gui.ListView]::new()
  $lstRecords.X = 0
  $lstRecords.Y = 0
  $lstRecords.Width = [Terminal.Gui.Dim]::Fill()
  $lstRecords.Height = [Terminal.Gui.Dim]::Fill()
  $lstRecords.SetSource(@("(Select a zone to load records)"))

  $recordsFrame.Add($lstRecords)
  $dialog.Add($recordsFrame)

  $y = 28  # Position buttons below frame

  ## SHARED STATE - DSA TUI Pattern
  $sharedState = @{
    CurrentRecords = @()           # Actual record objects
    AllRecordLines = @()           # Unfiltered display strings (like currentSearchOutputLines in DSA)
    FilteredRecordLines = @()      # Filtered display strings
    FilteredRecords = @()          # Filtered record objects
  }

  ## HELPER: Format record for display
  $formatRecord = {
    param($record)
    "$($record.HostName) [$($record.RecordType)] → $($record.RecordData.IPv4Address ?? $record.RecordData.NameServer ?? $record.RecordData.HostNameAlias ?? 'Data')"
  }

  ## HELPER: Show record details
  $showRecordDetailsFunc = {
    param($record)

    $details = "DNS Record Details`n"
    $details += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n"
    $details += "Hostname: $($record.HostName)`n"
    $details += "Record Type: $($record.RecordType)`n"
    $details += "TTL: $($record.TimeToLive)`n"

    switch ($record.RecordType) {
      'A' { $details += "`nIPv4: $($record.RecordData.IPv4Address)" }
      'AAAA' { $details += "`nIPv6: $($record.RecordData.IPv6Address)" }
      'CNAME' { $details += "`nAlias: $($record.RecordData.HostNameAlias)" }
      'MX' { $details += "`nMail Server: $($record.RecordData.MailExchange)`nPriority: $($record.RecordData.Preference)" }
      'NS' { $details += "`nName Server: $($record.RecordData.NameServer)" }
      'PTR' { $details += "`nPTR Domain: $($record.RecordData.PtrDomainName)" }
      'SRV' {
        $details += "`nTarget: $($record.RecordData.DomainName)`nPort: $($record.RecordData.Port)`nPriority: $($record.RecordData.Priority)`nWeight: $($record.RecordData.Weight)"
      }
      'TXT' { $details += "`nText: $($record.RecordData.DescriptiveText)" }
      default { $details += "`nData: $($record.RecordData)" }
    }

    $txtDetails.Text = $details
  }

  ## HELPER: Load records for zone
  $loadRecordsFunc = {
    param($zoneName)

    try {
      $records = Get-DnsServerResourceRecord -ZoneName $zoneName -ComputerName $ServerName -ErrorAction Stop
      $sharedState.CurrentRecords = @($records)

      # Create display strings (like currentSearchOutputLines in DSA)
      $sharedState.AllRecordLines = @($sharedState.CurrentRecords | ForEach-Object { & $formatRecord $_ })

      # Initially show all (no filter)
      $sharedState.FilteredRecords = $sharedState.CurrentRecords
      $sharedState.FilteredRecordLines = $sharedState.AllRecordLines

      $lstRecords.SetSource($sharedState.AllRecordLines)

      & $debugLogFunc "Loaded $($records.Count) records from $zoneName" -Type "Success"
    } catch {
      & $showModalFunc "Error" "Failed to load records from zone ${zoneName}:`n$($_.Exception.Message)"
      & $debugLogFunc "Failed to load records: $_" -Type "Error"
    }
  }

  ## AUTO-LOAD: Zone selection changed
  $cmbZones.add_SelectedItemChanged({
    $selectedIndex = $cmbZones.SelectedItem
    if ($selectedIndex -lt 0 -or $selectedIndex -ge $dnsZones.Count) { return }

    $zone = $dnsZones[$selectedIndex]
    & $loadRecordsFunc $zone.ZoneName
  }.GetNewClosure())

  ## REAL-TIME SEARCH FILTER - DSA TUI Pattern (EXACT)
  $txtSearch.add_TextChanged({
    if ($sharedState.AllRecordLines) {
      $search = $txtSearch.Text.ToString().Trim()
      if ($search) {
        # Filter the display lines
        $filteredLines = @($sharedState.AllRecordLines | Where-Object { $_ -match "(?i)$search" })

        # Also filter the actual record objects to match
        $filteredRecords = @()
        for ($i = 0; $i -lt $sharedState.AllRecordLines.Count; $i++) {
          if ($sharedState.AllRecordLines[$i] -match "(?i)$search") {
            $filteredRecords += $sharedState.CurrentRecords[$i]
          }
        }

        $sharedState.FilteredRecordLines = $filteredLines
        $sharedState.FilteredRecords = $filteredRecords

        if ($filteredLines.Count -gt 0) {
          $lstRecords.SetSource($filteredLines)
        } else {
          $lstRecords.SetSource(@("(No matches)"))
        }
      } else {
        # No filter - show all
        $sharedState.FilteredRecordLines = $sharedState.AllRecordLines
        $sharedState.FilteredRecords = $sharedState.CurrentRecords
        $lstRecords.SetSource($sharedState.AllRecordLines)
      }
    }
  }.GetNewClosure())

  ## REAL-TIME DETAILS: Record selection changed
  $lstRecords.add_SelectedItemChanged({
    $selectedIndex = $lstRecords.SelectedItem
    if ($selectedIndex -lt 0 -or $sharedState.FilteredRecords.Count -eq 0) { return }
    if ($selectedIndex -ge $sharedState.FilteredRecords.Count) { return }

    $record = $sharedState.FilteredRecords[$selectedIndex]
    & $showRecordDetailsFunc $record
  }.GetNewClosure())

  ## ========== BUTTONS ==========
  $btnX = 2

  # Create Record button
  $btnCreate = [Terminal.Gui.Button]::new("Create Record")
  $btnCreate.X = $btnX
  $btnCreate.Y = $y
  $btnCreate.add_Clicked({
    $selectedIndex = $cmbZones.SelectedItem
    if ($selectedIndex -lt 0 -or $selectedIndex -ge $dnsZones.Count) {
      & $showModalFunc "Info" "Please select a DNS zone first"
      return
    }

    $zone = $dnsZones[$selectedIndex]
    [Terminal.Gui.Application]::RequestStop()
    & $showCreateRecordFunc -ZoneName $zone.ZoneName -ServerName $ServerName
    & $showDNSDialogFunc -ServerName $ServerName
  }.GetNewClosure())
  $dialog.Add($btnCreate)
  $btnX += 17

  # Edit Record button
  $btnEdit = [Terminal.Gui.Button]::new("Edit Record")
  $btnEdit.X = $btnX
  $btnEdit.Y = $y
  $btnEdit.add_Clicked({
    $selectedZoneIndex = $cmbZones.SelectedItem
    $selectedRecordIndex = $lstRecords.SelectedItem

    if ($selectedZoneIndex -lt 0 -or $selectedZoneIndex -ge $dnsZones.Count) {
      & $showModalFunc "Info" "Please select a zone first"
      return
    }

    if ($sharedState.FilteredRecords.Count -eq 0) {
      & $showModalFunc "Info" "No records loaded or visible"
      return
    }

    if ($selectedRecordIndex -lt 0 -or $selectedRecordIndex -ge $sharedState.FilteredRecords.Count) {
      & $showModalFunc "Info" "Please select a DNS record to edit"
      return
    }

    $zone = $dnsZones[$selectedZoneIndex]
    $record = $sharedState.FilteredRecords[$selectedRecordIndex]

    [Terminal.Gui.Application]::RequestStop()
    & $showRecordPropsFunc -Record $record -ZoneName $zone.ZoneName -ServerName $ServerName
    & $showDNSDialogFunc -ServerName $ServerName
  }.GetNewClosure())
  $dialog.Add($btnEdit)
  $btnX += 15

  # Delete Record button
  $btnDelete = [Terminal.Gui.Button]::new("Delete Record")
  $btnDelete.X = $btnX
  $btnDelete.Y = $y
  $btnDelete.add_Clicked({
    $selectedZoneIndex = $cmbZones.SelectedItem
    $selectedRecordIndex = $lstRecords.SelectedItem

    if ($selectedZoneIndex -lt 0 -or $selectedZoneIndex -ge $dnsZones.Count) {
      & $showModalFunc "Info" "Please select a zone first"
      return
    }
    if ($sharedState.FilteredRecords.Count -eq 0) {
      & $showModalFunc "Info" "No records loaded"
      return
    }

    if ($selectedRecordIndex -lt 0 -or $selectedRecordIndex -ge $sharedState.FilteredRecords.Count) {
      & $showModalFunc "Info" "Please select a DNS record to delete"
      return
    }

    $zone = $dnsZones[$selectedZoneIndex]
    $record = $sharedState.FilteredRecords[$selectedRecordIndex]

    # Confirm
    $confirmDlg = [Terminal.Gui.Dialog]::new()
    $confirmDlg.Title = "Confirm Delete"
    $confirmDlg.Width = 60
    $confirmDlg.Height = 12

    $lblConfirm = [Terminal.Gui.Label]::new("Delete DNS record?`n`nHostname: $($record.HostName)`nType: $($record.RecordType)`nZone: $($zone.ZoneName)")
    $lblConfirm.X = 2
    $lblConfirm.Y = 1
    $confirmDlg.Add($lblConfirm)

    $confirmResult = 1
    $btnYes = [Terminal.Gui.Button]::new("Yes")
    $btnYes.add_Clicked({ $confirmResult = 0; [Terminal.Gui.Application]::RequestStop() })
    $confirmDlg.AddButton($btnYes)

    $btnNo = [Terminal.Gui.Button]::new("No")
    $btnNo.add_Clicked({ $confirmResult = 1; [Terminal.Gui.Application]::RequestStop() })
    $confirmDlg.AddButton($btnNo)
    [Terminal.Gui.Application]::Run($confirmDlg)

    if ($confirmResult -eq 0) {
      try {
        Remove-DnsServerResourceRecord -ZoneName $zone.ZoneName -InputObject $record -ComputerName $ServerName -Force -ErrorAction Stop
        & $showModalFunc "Success" "DNS record deleted!"
        & $debugLogFunc "Deleted: $($record.HostName) [$($record.RecordType)]" -Type "Success"
        ## Reload
        & $loadRecordsFunc $zone.ZoneName
      } catch {
        & $showModalFunc "Error" "Failed to delete:`n$($_.Exception.Message)"
      }
    }
  }.GetNewClosure())
  $dialog.Add($btnDelete)
  $btnX += 17
  ## Zone Properties button
  $btnZoneProps = [Terminal.Gui.Button]::new("Zone Properties")
  $btnZoneProps.X = $btnX
  $btnZoneProps.Y = $y
  $btnZoneProps.add_Clicked({
    $selectedIndex = $cmbZones.SelectedItem
    if ($selectedIndex -lt 0 -or $selectedIndex -ge $dnsZones.Count) {
      & $showModalFunc "Info" "Please select a DNS zone"
      return
    }

    $zone = $dnsZones[$selectedIndex]
    [Terminal.Gui.Application]::RequestStop()
    & $showZonePropsFunc -ZoneName $zone.ZoneName -ZoneType $(if ($zone.IsReverse) { 'Reverse' } else { 'Forward' })
    & $showDNSDialogFunc -ServerName $ServerName
  }.GetNewClosure())
  $dialog.Add($btnZoneProps)
  $btnX += 19

  # Import button
  $btnImport = [Terminal.Gui.Button]::new("Import")
  $btnImport.X = $btnX
  $btnImport.Y = $y
  $btnImport.add_Clicked({
    [Terminal.Gui.Application]::RequestStop()
    & $showImportFunc
    & $showDNSDialogFunc -ServerName $ServerName
  }.GetNewClosure())
  $dialog.Add($btnImport)
  $btnX += 10

  # Export button
  $btnExport = [Terminal.Gui.Button]::new("Export")
  $btnExport.X = $btnX
  $btnExport.Y = $y
  $btnExport.add_Clicked({
    if ($sharedState.CurrentRecords.Count -eq 0) {
      & $showModalFunc "Info" "No records loaded. Select a zone first."
      return
    }

    try {
      $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
      $selectedZoneIndex = $cmbZones.SelectedItem
      $zoneName = if ($selectedZoneIndex -ge 0) { $dnsZones[$selectedZoneIndex].ZoneName } else { "all" }
      $filename = "dns_records_${zoneName}_$timestamp.csv"

      $exportData = $sharedState.CurrentRecords | ForEach-Object {
        [PSCustomObject]@{
          HostName = $_.HostName
          RecordType = $_.RecordType
          TTL = $_.TimeToLive
          RecordData = $_.RecordData | Out-String
        }
      }

      $exportData | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
      & $showModalFunc "Success" "Exported $($sharedState.CurrentRecords.Count) records to:`n$filename"
      & $debugLogFunc "Exported to $filename" -Type "Success"
    } catch {
      & $showModalFunc "Error" "Export failed:`n$($_.Exception.Message)"
    }
  }.GetNewClosure())
  $dialog.Add($btnExport)
  $btnX += 10

  # Close button
  $btnClose = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = $btnX
  $btnClose.Y = $y
  $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $dialog.Add($btnClose)

  [Terminal.Gui.Application]::Run($dialog)
}

function Show-DNSToolsInfo {
  $termHeight = [Console]::WindowHeight
  $dlgHeight  = [Math]::Floor($termHeight * 0.30)
  if ($dlgHeight -lt 14) { $dlgHeight = 14 }
  $dialog       = [Terminal.Gui.Dialog]::new("DNS Tools Information", 60, $dlgHeight)
  $hasDNSServer = $null -ne (Get-Command Get-DnsServer -ErrorAction SilentlyContinue)
  $hasResolve   = $null -ne (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)
  $hasPing      = $null -ne (Get-Command Test-Connection -ErrorAction SilentlyContinue)
  $hasTrace     = $null -ne (Get-Command Test-NetConnection -ErrorAction SilentlyContinue)
  $y = 1
  $dialog.Add([Terminal.Gui.Label]::new(2, $y, "Available DNS Management Tools:"))
  $y += 2
  $toolStatus = @()
  if ($hasDNSServer) { $toolStatus += "DNS Server Module:     ✔  Available" } else { $toolStatus += "DNS Server Module:     ✖  Not Found" }
  if ($hasResolve)   { $toolStatus += "Resolve-DnsName:       ✔  Available" } else { $toolStatus += "Resolve-DnsName:       ✖  Not Found" }
  if ($hasPing)      { $toolStatus += "Test-Connection:       ✔  Available" } else { $toolStatus += "Test-Connection:       ✖  Not Found" }
  if ($hasTrace)     { $toolStatus += "Test-NetConnection:    ✔  Available" } else { $toolStatus += "Test-NetConnection:    ✖  Not Found" }
  $toolStatus | ForEach-Object {
    $dialog.Add([Terminal.Gui.Label]::new(4, $y, $_))
    $y++
  }
  $y += 1
  $dialog.Add([Terminal.Gui.Label]::new(2, $y, "Connected to: $($global:DetectedDnsServer)"))
  $y++
  $dialog.Add([Terminal.Gui.Label]::new(2, $y, "Status: $(if ($global:DNSConnectionStatus.IsConnected) { '✔ Connected' } else { '✖ Not Connected' })"))
  $btnOK = [Terminal.Gui.Button]::new("OK")
  $btnOK.X = [Terminal.Gui.Pos]::Center()
  $btnOK.Y = $dlgHeight - 3
  $btnOK.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $dialog.AddButton($btnOK)
  [Terminal.Gui.Application]::Run($dialog)
}

function Show-Dashboard {

    $global:AutoRefresh.CurrentView = { Show-Dashboard }

    ## Remove existing dashboard/tree if refreshing
    if ($script:dashboardPane) {
        $script:win.Remove($script:dashboardPane)
        $script:dashboardPane = $null
    }
    if ($script:treeFrame) {
        $script:win.Remove($script:treeFrame)
        $script:treeFrame = $null
        $script:tree = $null
    }

    ## -------------------- TREE VIEW (LEFT 70%) --------------------
    Debug-Log "Creating DNS Zone TreeView..." -Type "Tracing"
    $treeFrame = [Terminal.Gui.FrameView]::new("DNS Zones")
    $treeFrame.X = 0
    $treeFrame.Y = 0
    $treeFrame.Width  = [Terminal.Gui.Dim]::Percent(70)
    $treeFrame.Height = [Terminal.Gui.Dim]::Fill()

    $tree = [Terminal.Gui.TreeView]::new()
    $tree.X = 0
    $tree.Y = 0
    $tree.Width  = [Terminal.Gui.Dim]::Fill()
    $tree.Height = [Terminal.Gui.Dim]::Fill()

    $treeFrame.Add($tree)

    ## -------------------- DASHBOARD (RIGHT 30%) --------------------
    $dashboardPane = [Terminal.Gui.FrameView]::new()
    $dashboardPane.Title = "Dashboard"
    $dashboardPane.X = [Terminal.Gui.Pos]::Percent(70)
    $dashboardPane.Y = 0
    $dashboardPane.Width = [Terminal.Gui.Dim]::Percent(30)
    $dashboardPane.Height = [Terminal.Gui.Dim]::Fill()

    ## ADD BOTH TO WINDOW BEFORE SETTING SCRIPT VARIABLES
    $script:win.Add($treeFrame)
    $script:win.Add($dashboardPane)

    ## NOW set script variables
    $script:treeFrame = $treeFrame
    $script:tree = $tree
    $script:dashboardPane = $dashboardPane

    ## -------------------- Server Information --------------------
    $serverFrame = [Terminal.Gui.FrameView]::new()
    $serverFrame.Title = "Server Information"
    $serverFrame.X = 0
    $serverFrame.Y = 0
    $serverFrame.Width = [Terminal.Gui.Dim]::Fill()
    $serverFrame.Height = 9

    $infoText = @"
Operating System : $((Get-CimInstance Win32_OperatingSystem).Caption)
(DOMAIN)\User    : $env:USERDOMAIN\$env:USERNAME
DNS Server       : $($global:DetectedDnsServer)
Connection       : $(if ($global:DNSConnectionStatus.IsConnected) { 'Connected ✔' } else { 'Not Connected ✖' })
Project Name     : $($Script:ProjectName)
Project Version  : $($Script:BuildVersion)
Project Codename : $($Script:FruitName)
"@

    $infoLabel = [Terminal.Gui.Label]::new()
    $infoLabel.X = 1
    $infoLabel.Y = 0
    $infoLabel.Width = [Terminal.Gui.Dim]::Fill(1)
    $infoLabel.Height = [Terminal.Gui.Dim]::Fill()
    $infoLabel.Text = $infoText
    $serverFrame.Add($infoLabel)
    $dashboardPane.Add($serverFrame)

    ## -------------------- DNS Statistics --------------------
    $lastTopFrame = $serverFrame
    if ($global:DNSConnectionStatus.IsConnected) {
        $statsFrame = [Terminal.Gui.FrameView]::new()
        $statsFrame.Title = "DNS Statistics"
        $statsFrame.X = 0
        $statsFrame.Y = [Terminal.Gui.Pos]::Bottom($serverFrame)
        $statsFrame.Width = [Terminal.Gui.Dim]::Fill()
        $statsFrame.Height = 7

        try {
            $zones = Get-SafeDnsServerZone -DnsServerName $global:DetectedDnsServer
            if ($null -eq $zones) { $zones = @() }
            $forwardZones = ($zones | Where-Object { -not $_.IsReverse }).Count
            $reverseZones = ($zones | Where-Object { $_.IsReverse }).Count

            $statsText = @"
Total Zones: $($zones.Count)
Forward Zones: $forwardZones
Reverse Zones: $reverseZones
Operations: $($global:PerformanceCounters.OperationCount)
Errors: $($global:PerformanceCounters.ErrorCount)
Avg Response: $([math]::Round($global:PerformanceCounters.AverageResponseTime, 2))ms
"@

            $statsLabel = [Terminal.Gui.Label]::new()
            $statsLabel.X = 1
            $statsLabel.Y = 0
            $statsLabel.Width = [Terminal.Gui.Dim]::Fill(1)
            $statsLabel.Height = [Terminal.Gui.Dim]::Fill()
            $statsLabel.Text = $statsText

            $statsFrame.Add($statsLabel)
            $dashboardPane.Add($statsFrame)
            $lastTopFrame = $statsFrame
        }
        catch { }
    }

    ## -------------------- Zone Management --------------------
    $zoneActionsFrame = [Terminal.Gui.FrameView]::new()
    $zoneActionsFrame.Title = "Zone Management"
    $zoneActionsFrame.X = 0
    $zoneActionsFrame.Y = [Terminal.Gui.Pos]::Bottom($lastTopFrame)
    $zoneActionsFrame.Width = [Terminal.Gui.Dim]::Fill()
    $zoneActionsFrame.Height = 10

    # Capture functions for closures
    $showDNSDialogFunc = ${function:Show-DNSDialog}
    $showCreateForwardFunc = ${function:Show-CreateForwardZoneDialog}
    $showCreateReverseFunc = ${function:Show-CreateReverseZoneDialog}
    $showModalFunc = ${function:Show-Modal}
    $showDashboardFunc = ${function:Show-Dashboard}

    $btnViewZones = [Terminal.Gui.Button]::new("📋 DNS Zones & Records")
    $btnViewZones.X = 2
    $btnViewZones.Y = 1
    $btnViewZones.Width = [Terminal.Gui.Dim]::Fill(2)
    $btnViewZones.add_Clicked({
        if (-not $global:DNSConnectionStatus.IsConnected) {
            & $showModalFunc "Error" "Not connected to DNS server"
            return
        }
        & $showDNSDialogFunc -ServerName $global:DetectedDnsServer
    }.GetNewClosure())
    $zoneActionsFrame.Add($btnViewZones)

    $btnCreateForward = [Terminal.Gui.Button]::new("➕ Create Forward Zone")
    $btnCreateForward.X = 2
    $btnCreateForward.Y = 3
    $btnCreateForward.Width = [Terminal.Gui.Dim]::Fill(2)
    $btnCreateForward.add_Clicked({
        if (-not $global:DNSConnectionStatus.IsConnected) {
            & $showModalFunc "Error" "Not connected to DNS server"
            return
        }
        & $showCreateForwardFunc
    }.GetNewClosure())
    $zoneActionsFrame.Add($btnCreateForward)

    $btnCreateReverse = [Terminal.Gui.Button]::new("➕ Create Reverse Zone")
    $btnCreateReverse.X = 2
    $btnCreateReverse.Y = 5
    $btnCreateReverse.Width = [Terminal.Gui.Dim]::Fill(2)
    $btnCreateReverse.add_Clicked({
        if (-not $global:DNSConnectionStatus.IsConnected) {
            & $showModalFunc "Error" "Not connected to DNS server"
            return
        }
        & $showCreateReverseFunc
    }.GetNewClosure())
    $zoneActionsFrame.Add($btnCreateReverse)

    $btnRefresh = [Terminal.Gui.Button]::new("🔄 Refresh Dashboard")
    $btnRefresh.X = 2
    $btnRefresh.Y = 7
    $btnRefresh.Width = [Terminal.Gui.Dim]::Fill(2)
    $btnRefresh.add_Clicked({
        & $showDashboardFunc
    }.GetNewClosure())
    $zoneActionsFrame.Add($btnRefresh)

    $dashboardPane.Add($zoneActionsFrame)

    ## -------------------- Records & Tools --------------------
    $recordActionsFrame = [Terminal.Gui.FrameView]::new()
    $recordActionsFrame.Title = "Records & Tools"
    $recordActionsFrame.X = 0
    $recordActionsFrame.Y = [Terminal.Gui.Pos]::Bottom($zoneActionsFrame)
    $recordActionsFrame.Width = [Terminal.Gui.Dim]::Fill()
    $recordActionsFrame.Height = [Terminal.Gui.Dim]::Fill(3)

    # Capture more functions
    $showConnectFunc = ${function:Show-ConnectDialog}
    $showToolsFunc = ${function:Show-Tools}
    $showImportFunc = ${function:Show-ImportDialog}
    $showExportFunc = ${function:Show-ExportDialog}

    $btnConnect = [Terminal.Gui.Button]::new("🔌 Connect to Server")
    $btnConnect.X = 2
    $btnConnect.Y = 1
    $btnConnect.Width = [Terminal.Gui.Dim]::Fill(2)
    $btnConnect.add_Clicked({
        & $showConnectFunc
    }.GetNewClosure())
    $recordActionsFrame.Add($btnConnect)

    $btnTools = [Terminal.Gui.Button]::new("🔧 Diagnostic Tools")
    $btnTools.X = 2
    $btnTools.Y = 3
    $btnTools.Width = [Terminal.Gui.Dim]::Fill(2)
    $btnTools.add_Clicked({
        & $showToolsFunc
    }.GetNewClosure())
    $recordActionsFrame.Add($btnTools)

    $btnImport = [Terminal.Gui.Button]::new("📥 Import Zone File")
    $btnImport.X = 2
    $btnImport.Y = 5
    $btnImport.Width = [Terminal.Gui.Dim]::Fill(2)
    $btnImport.add_Clicked({
        if (-not $global:DNSConnectionStatus.IsConnected) {
            & $showModalFunc "Error" "Not connected to DNS server"
            return
        }
        & $showImportFunc
    }.GetNewClosure())
    $recordActionsFrame.Add($btnImport)

    $btnExport = [Terminal.Gui.Button]::new("📤 Export Zone File")
    $btnExport.X = 2
    $btnExport.Y = 7
    $btnExport.Width = [Terminal.Gui.Dim]::Fill(2)
    $btnExport.add_Clicked({
        if (-not $global:DNSConnectionStatus.IsConnected) {
            & $showModalFunc "Error" "Not connected to DNS server"
            return
        }
        & $showExportFunc
    }.GetNewClosure())
    $recordActionsFrame.Add($btnExport)

    $dashboardPane.Add($recordActionsFrame)

    ## -------------------- Status (BOTTOM) --------------------
    $statusFrame = [Terminal.Gui.FrameView]::new()
    $statusFrame.Title = "Status"
    $statusFrame.X = 0
    $statusFrame.Y = [Terminal.Gui.Pos]::AnchorEnd(3)
    $statusFrame.Width = [Terminal.Gui.Dim]::Fill()
    $statusFrame.Height = 3

    $statusText = if ($global:DNSConnectionStatus.IsConnected) {
        "✔ Connected to $($global:DetectedDnsServer) | Last Update: $(Get-Date -Format 'HH:mm:ss')"
    }
    else { "✖ Not connected | Use 'Connect to Server' to establish connection" }

    $statusLabel = [Terminal.Gui.Label]::new()
    $statusLabel.X = 1
    $statusLabel.Y = 0
    $statusLabel.Width = [Terminal.Gui.Dim]::Fill(1)
    $statusLabel.Text = $statusText

    $statusFrame.Add($statusLabel)
    $dashboardPane.Add($statusFrame)

    ## -------------------- REBUILD TREE --------------------
    if ($global:DNSConnectionStatus.IsConnected -and $null -ne $tree) {
      try {
        Debug-Log "Refreshing DNS zone tree..." -Type "Insight"
        Build-DNSZoneTree -dnsServer $global:DetectedDnsServer
      }
      catch { Debug-Log "Failed to rebuild tree: $_" -Type "Warning" }
    }

    ## -------------------- ADD F12 CONTEXT MENU HANDLER --------------------
    if ($null -ne $tree) {
      $showContextFunc = ${function:Show-DNSObjectContextMenu}
      $showModalFunc = ${function:Show-Modal}
      $tree.add_KeyPress({
        param($e)

        if ($e.KeyEvent.Key -eq [Terminal.Gui.Key]::F12) {
          $selectedNode = $script:tree.SelectedObject
          if ($null -eq $selectedNode -or $null -eq $selectedNode.Tag) {
            & $showModalFunc "Info" "No item selected"
            $e.Handled = $true
            return
          }

          $tag = $selectedNode.Tag
          if ($tag -is [hashtable]) {
            $objectType = $tag['Type']
            $object = $tag['Object']
            ## Normalize zone types
            if ($objectType -eq 'ForwardZone' -or $objectType -eq 'ReverseZone') { $objectType = 'Zone' }
            if ($null -ne $objectType -and $null -ne $object) {
              & $showContextFunc -Object $object -ObjectType $objectType
            } else {
              & $showModalFunc "Info" "No context menu available for this item"
            }
          }
          $e.Handled = $true
        }
    }.GetNewClosure())
    Debug-Log "F12 context menu handler added to tree" -Type "Tracing"
  }
}


function Show-CreateForwardZoneDialog {
  $zoneData = [PSCustomObject]@{
    ZoneName = ""
    ReplicationScope = "Domain"
    ZoneType = "Primary"
    DynamicUpdate = "Secure"
  }

  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $data, $state)

      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Zone Configuration"

      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Zone Name:" -FieldName 'txtZoneName' -State $state -Value "" -Width 50 -FieldX 25

      $y += 1
      $lblHelp = [Terminal.Gui.Label]::new("(e.g., example.com, subdomain.example.com)")
      $lblHelp.X = 25; $lblHelp.Y = $y
      $view.Add($lblHelp)
      $y += 2

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Zone Type" -SpaceBefore 0

      $lblType = [Terminal.Gui.Label]::new("Zone Type:")
      $lblType.X = 2; $lblType.Y = $y
      $view.Add($lblType)

      $state.rbPrimary = [Terminal.Gui.RadioGroup]::new()
      $state.rbPrimary.X = 25
      $state.rbPrimary.Y = $y
      $state.rbPrimary.RadioLabels = @("Primary Zone", "Secondary Zone", "Stub Zone")
      $state.rbPrimary.SelectedItem = 0
      $view.Add($state.rbPrimary)
      $y += 4

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Active Directory Integration" -SpaceBefore 0

      $lblRep = [Terminal.Gui.Label]::new("Replication Scope:")
      $lblRep.X = 2; $lblRep.Y = $y
      $view.Add($lblRep)

      $state.cmbReplication = [Terminal.Gui.ComboBox]::new()
      $state.cmbReplication.X = 25; $state.cmbReplication.Y = $y; $state.cmbReplication.Width = 30
      $state.cmbReplication.SetSource(@("Domain", "Forest", "Legacy"))
      $state.cmbReplication.SelectedItem = 0
      $view.Add($state.cmbReplication)
      $y += 2

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Dynamic Updates" -SpaceBefore 0

      $lblDyn = [Terminal.Gui.Label]::new("Dynamic Update:")
      $lblDyn.X = 2; $lblDyn.Y = $y
      $view.Add($lblDyn)

      $state.cmbDynamic = [Terminal.Gui.ComboBox]::new()
      $state.cmbDynamic.X = 25; $state.cmbDynamic.Y = $y; $state.cmbDynamic.Width = 30
      $state.cmbDynamic.SetSource(@("Secure", "Nonsecure and secure", "None"))
      $state.cmbDynamic.SelectedItem = 0
      $view.Add($state.cmbDynamic)
    }
  }

  $advancedTab = @{
    Name = "Advanced"
    Builder = {
      param($view, $data, $state)

      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Aging/Scavenging"

      $state.chkScavenging = [Terminal.Gui.CheckBox]::new("Enable scavenging of stale records")
      $state.chkScavenging.X = 4; $state.chkScavenging.Y = $y
      $view.Add($state.chkScavenging)
      $y += 2

      $lblRefresh = [Terminal.Gui.Label]::new("No-refresh interval (days):")
      $lblRefresh.X = 4; $lblRefresh.Y = $y
      $view.Add($lblRefresh)

      $state.txtNoRefresh = [Terminal.Gui.TextField]::new("7")
      $state.txtNoRefresh.X = 35; $state.txtNoRefresh.Y = $y; $state.txtNoRefresh.Width = 10
      $view.Add($state.txtNoRefresh)
      $y += 1

      $lblScavenge = [Terminal.Gui.Label]::new("Refresh interval (days):")
      $lblScavenge.X = 4; $lblScavenge.Y = $y
      $view.Add($lblScavenge)

      $state.txtRefresh = [Terminal.Gui.TextField]::new("7")
      $state.txtRefresh.X = 35; $state.txtRefresh.Y = $y; $state.txtRefresh.Width = 10
      $view.Add($state.txtRefresh)
      $y += 2

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Zone Transfer" -SpaceBefore 1

      $state.rbTransfer = [Terminal.Gui.RadioGroup]::new()
      $state.rbTransfer.X = 4
      $state.rbTransfer.Y = $y
      $state.rbTransfer.RadioLabels = @(
        "Allow zone transfers to any server",
        "Allow zone transfers only to name servers",
        "Allow zone transfers only to specific servers",
        "Do not allow zone transfers"
      )
      $state.rbTransfer.SelectedItem = 1
      $view.Add($state.rbTransfer)
    }
  }

  $onOK = {
    param($data, $state)

    $zoneName = $state.txtZoneName.Text.ToString().Trim()

    if ([string]::IsNullOrWhiteSpace($zoneName)) {
      Show-Modal "Error" "Please enter a zone name"
      return
    }

    try {
      $replicationScope = @("Domain", "Forest", "Legacy")[$state.cmbReplication.SelectedItem]
      $zoneType = @("Primary", "Secondary", "Stub")[$state.rbPrimary.SelectedItem]

      if ($zoneType -eq "Primary") {
        $params = @{
          Name = $zoneName
          ReplicationScope = $replicationScope
          ComputerName = $global:DetectedDnsServer
          ErrorAction = "Stop"
        }

        if ($state.chkScavenging.Checked) {
          $params.ScavengingInterval = "$($state.txtRefresh.Text.ToString()):00:00:00"
        }

        Add-DnsServerPrimaryZone @params
        Debug-Log "Forward zone created: $zoneName ($replicationScope)" -Type "Success"
        Show-DNSDialog
      }
      else {
        Show-Modal "Info" "Secondary and Stub zone creation require additional configuration"
      }
    }
    catch {
      Debug-Log "Failed to create zone: $_" -Type "Error"
      Show-Modal "Error" "Failed to create zone: $_"
    }
  }

  New-PropertiesDialog -Title "Create New Forward Lookup Zone" -Width 90 -Height 35 -Tabs @($generalTab, $advancedTab) -Data $zoneData -OnOK $onOK -IncludeSearchTab $false
}

function Show-CreateReverseZoneDialog {
  $zoneData = [PSCustomObject]@{
    NetworkID = "192.168.1"
    ReplicationScope = "Domain"
    ZoneType = "Primary"
  }

  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $data, $state)

      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Reverse Lookup Zone Configuration"

      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Network ID:" -FieldName 'txtNetworkID' -State $state -Value "192.168.1" -Width 40 -FieldX 25

      $y += 1
      $lblHelp = [Terminal.Gui.Label]::new("(e.g., 192.168.1 for Class C or 10.0 for Class B)")
      $lblHelp.X = 25; $lblHelp.Y = $y
      $view.Add($lblHelp)
      $y += 1

      $lblZone = [Terminal.Gui.Label]::new("Zone name will be:")
      $lblZone.X = 25; $lblZone.Y = $y
      $view.Add($lblZone)
      $y += 1

      $state.lblPreview = [Terminal.Gui.Label]::new("1.168.192.in-addr.arpa")
      $state.lblPreview.X = 25; $state.lblPreview.Y = $y
      $view.Add($state.lblPreview)
      $y += 2

      # Update preview when network ID changes
      $state.txtNetworkID.add_TextChanged({
        $netId = $state.txtNetworkID.Text.ToString().Trim()
        if (-not [string]::IsNullOrWhiteSpace($netId)) {
          $octets = $netId.Split('.')
          $reverseName = "$($octets[-1..-$octets.Length] -join '.').in-addr.arpa"
          $state.lblPreview.Text = [NStack.ustring]::Make($reverseName)
        }
      }.GetNewClosure())

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Zone Type" -SpaceBefore 0

      $lblType = [Terminal.Gui.Label]::new("Zone Type:")
      $lblType.X = 2; $lblType.Y = $y
      $view.Add($lblType)

      $state.rbPrimary = [Terminal.Gui.RadioGroup]::new()
      $state.rbPrimary.X = 25
      $state.rbPrimary.Y = $y
      $state.rbPrimary.RadioLabels = @("Primary Zone", "Secondary Zone", "Stub Zone")
      $state.rbPrimary.SelectedItem = 0
      $view.Add($state.rbPrimary)
      $y += 4

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Active Directory Integration" -SpaceBefore 0

      $lblRep = [Terminal.Gui.Label]::new("Replication Scope:")
      $lblRep.X = 2; $lblRep.Y = $y
      $view.Add($lblRep)

      $state.cmbReplication = [Terminal.Gui.ComboBox]::new()
      $state.cmbReplication.X = 25; $state.cmbReplication.Y = $y; $state.cmbReplication.Width = 30
      $state.cmbReplication.SetSource(@("Domain", "Forest", "Legacy"))
      $state.cmbReplication.SelectedItem = 0
      $view.Add($state.cmbReplication)
    }
  }

  $advancedTab = @{
    Name = "Advanced"
    Builder = {
      param($view, $data, $state)

      $y = 1
      Add-SectionHeader -View $view -Y ([ref]$y) -Text "PTR Record Options"

      $state.chkUpdatePtr = [Terminal.Gui.CheckBox]::new("Update associated PTR records")
      $state.chkUpdatePtr.X = 4; $state.chkUpdatePtr.Y = $y
      $state.chkUpdatePtr.Checked = $true
      $view.Add($state.chkUpdatePtr)
      $y += 2

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Aging/Scavenging" -SpaceBefore 1

      $state.chkScavenging = [Terminal.Gui.CheckBox]::new("Enable scavenging of stale records")
      $state.chkScavenging.X = 4; $state.chkScavenging.Y = $y
      $view.Add($state.chkScavenging)
      $y += 2

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Zone Transfer" -SpaceBefore 1

      $state.rbTransfer = [Terminal.Gui.RadioGroup]::new()
      $state.rbTransfer.X = 4
      $state.rbTransfer.Y = $y
      $state.rbTransfer.RadioLabels = @(
        "Allow zone transfers to any server",
        "Allow zone transfers only to name servers",
        "Do not allow zone transfers"
      )
      $state.rbTransfer.SelectedItem = 1
      $view.Add($state.rbTransfer)
    }
  }

  $onOK = {
    param($data, $state)

    $networkId = $state.txtNetworkID.Text.ToString().Trim()

    if ([string]::IsNullOrWhiteSpace($networkId)) {
      Show-Modal "Error" "Please enter a network ID"
      return
    }

    try {
      $replicationScope = @("Domain", "Forest", "Legacy")[$state.cmbReplication.SelectedItem]
      $zoneType = @("Primary", "Secondary", "Stub")[$state.rbPrimary.SelectedItem]

      if ($zoneType -eq "Primary") {
        $octets = $networkId.Split('.')
        $reverseName = "$($octets[-1..-$octets.Length] -join '.').in-addr.arpa"

        Add-DnsServerPrimaryZone `
          -NetworkId $networkId `
          -ReplicationScope $replicationScope `
          -ComputerName $global:DetectedDnsServer `
          -ErrorAction Stop

        Debug-Log "Reverse zone created: $reverseName ($replicationScope)" -Type "Success"
        Show-DNSDialog
      }
      else {
        Show-Modal "Info" "Secondary and Stub zone creation require additional configuration"
      }
    }
    catch {
      Debug-Log "Failed to create reverse zone: $_" -Type "Error"
      Show-Modal "Error" "Failed to create reverse zone: $_"
    }
  }
  New-PropertiesDialog -Title "Create New Reverse Lookup Zone" -Width 90 -Height 35 -Tabs @($generalTab, $advancedTab) -Data $zoneData -OnOK $onOK -IncludeSearchTab $false
}

function Load-DNSRecords {
  param(
    [string]$ZoneName,
    [Terminal.Gui.TableView]$TableView
  )

  try {
    $records = Get-DnsServerResourceRecord -ZoneName $ZoneName -ComputerName $global:DetectedDnsServer -ErrorAction Stop
    $dataTable = New-Object System.Data.DataTable
    $dataTable.Columns.Add("Name") | Out-Null
    $dataTable.Columns.Add("Type") | Out-Null
    $dataTable.Columns.Add("Data") | Out-Null
    $dataTable.Columns.Add("TTL") | Out-Null

    foreach ($record in $records) {
      if ($record.RecordType -in @("SOA", "NS") -and $record.HostName -eq "@") { continue }
      $row = $dataTable.NewRow()
      $row["Name"] = $record.HostName
      $row["Type"] = $record.RecordType
      $row["Data"] = Format-RecordData -record $record
      $row["TTL"] = $record.TimeToLive.TotalSeconds
      $dataTable.Rows.Add($row)
    }
    $TableView.Table = $dataTable
    Debug-Log "Loaded $($records.Count) records for zone $ZoneName" Type "Insight"
  } catch {
    Debug-Log "Error loading DNS records: $_" -Type "Problem"
    Show-Modal "Error" "Failed to load DNS records: $_"
  }
}

function Show-CreateRecordDialog {
  param([string]$ZoneName)

  ## Create dialog with explicit dimensions
  $dialog = [Terminal.Gui.Dialog]::new()
  $dialog.Title = "Create DNS Record - $ZoneName"
  $dialog.Width = 70
  $dialog.Height = 18

  ## Name
  $lblName = [Terminal.Gui.Label]::new()
  $lblName.X = 1
  $lblName.Y = 1
  $lblName.Text = "Name:"
  $dialog.Add($lblName)
  $txtName = [Terminal.Gui.TextField]::new()
  $txtName.X = 15
  $txtName.Y = 1
  $txtName.Width = 50
  $dialog.Add($txtName)

  ## Type
  $lblType = [Terminal.Gui.Label]::new()
  $lblType.X = 1
  $lblType.Y = 3
  $lblType.Text = "Type:"
  $dialog.Add($lblType)
  $typeList = [Terminal.Gui.ListView]::new()
  $typeList.X = 15
  $typeList.Y = 3
  $typeList.Width = 20
  $typeList.Height = 5
  $typeList.SetSource([string[]]@("A", "AAAA", "CNAME", "TXT"))
  $dialog.Add($typeList)

  ## Data
  $lblData = [Terminal.Gui.Label]::new()
  $lblData.X = 1
  $lblData.Y = 9
  $lblData.Text = "Data:"
  $dialog.Add($lblData)
  $txtData = [Terminal.Gui.TextField]::new()
  $txtData.X = 15
  $txtData.Y = 9
  $txtData.Width = 50
  $dialog.Add($txtData)

  ## TTL
  $lblTTL = [Terminal.Gui.Label]::new()
  $lblTTL.X = 1
  $lblTTL.Y = 11
  $lblTTL.Text = "TTL (seconds):"
  $dialog.Add($lblTTL)
  $txtTTL = [Terminal.Gui.TextField]::new()
  $txtTTL.X = 15
  $txtTTL.Y = 11
  $txtTTL.Width = 15
  $txtTTL.Text = "3600"
  $dialog.Add($txtTTL)

  ## Create button
  $btnCreate = [Terminal.Gui.Button]::new()
  $btnCreate.X = 20
  $btnCreate.Y = 14
  $btnCreate.Text = "Create"
  $btnCreate.IsDefault = $true
  $btnCreate.add_Clicked({
    $name = $txtName.Text.ToString()
    $type = @("A", "AAAA", "CNAME", "TXT")[$typeList.SelectedItem]
    $data = $txtData.Text.ToString()
    $ttl = $txtTTL.Text.ToString()
    if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($data)) {
      Show-Modal "Error" "Please fill in all fields"
      return
    }
    try {
      Create-DNSRecord -ZoneName $ZoneName -Name $name -Type $type -Data $data -TTL $ttl
      [Terminal.Gui.Application]::RequestStop()
    } catch {
      Show-Modal "Error" "Failed to create record: $_"
    }
  })
  $dialog.AddButton($btnCreate)

  ## Cancel button
  $btnCancel = [Terminal.Gui.Button]::new()
  $btnCancel.X = 35
  $btnCancel.Y = 14
  $btnCancel.Text = "Cancel"
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $dialog.AddButton($btnCancel)
  [Terminal.Gui.Application]::Run($dialog)
}

function Create-DNSRecord {
  param(
    [string]$ZoneName,
    [string]$Name,
    [string]$Type,
    [string]$Data,
    [string]$TTL
  )

  try {
    $ttlValue = [int]$TTL
    $timeSpan = [TimeSpan]::FromSeconds($ttlValue)

    switch ($Type) {
      "A"     { Add-DnsServerResourceRecordA -ZoneName $ZoneName -Name $Name -IPv4Address $Data -TimeToLive $timeSpan -ComputerName $global:DetectedDnsServer -ErrorAction Stop }
      "AAAA"  { Add-DnsServerResourceRecordAAAA -ZoneName $ZoneName -Name $Name -IPv6Address $Data -TimeToLive $timeSpan -ComputerName $global:DetectedDnsServer -ErrorAction Stop }
      "CNAME" { Add-DnsServerResourceRecordCName -ZoneName $ZoneName -Name $Name -HostNameAlias $Data -TimeToLive $timeSpan -ComputerName $global:DetectedDnsServer -ErrorAction Stop }
      "TXT"   { Add-DnsServerResourceRecordTxt -ZoneName $ZoneName -Name $Name -DescriptiveText $Data -TimeToLive $timeSpan -ComputerName $global:DetectedDnsServer -ErrorAction Stop }
    }
    Debug-Log "DNS record created: $Type $Name in zone $ZoneName" -Type "Success"
    Show-Modal "Success" "DNS record created successfully!"
  } catch {
    Debug-Log "Failed to create DNS record: $_" -Type "Error"
    Show-Modal "Error" "Failed to create DNS record: $_"
  }
}

function Delete-DNSRecord {
  param(
    [string]$ZoneName,
    [string]$RecordName,
    [string]$RecordType,
    [Terminal.Gui.TableView]$TableView
  )
  $result = Show-Modal "Confirm Delete" "Delete this DNS record?`n`nZone: $ZoneName`nName: $RecordName`nType: $RecordType" -YesNo
  if ($result -ne 0) { return }
    try {
      $record = Get-DnsServerResourceRecord -ZoneName $ZoneName -Name $RecordName -RRType $RecordType -ComputerName $global:DetectedDnsServer -ErrorAction Stop | Select-Object -First 1
      if ($record) {
        Remove-DnsServerResourceRecord -ZoneName $ZoneName -InputObject $record -ComputerName $global:DetectedDnsServer -Force -ErrorAction Stop
        Debug-Log "DNS record deleted: $RecordType $RecordName from zone $ZoneName" -Type "Success"
        Show-Modal "Success" "DNS record deleted successfully!"
        Load-DNSRecords -ZoneName $ZoneName -TableView $TableView
      }
    } catch {
      Debug-Log "Failed to delete DNS record: $_" -Type "Error"
      Show-Modal "Error" "Failed to delete DNS record: $_"
  }
}

function Delete-Zone {
  param([string]$ZoneName)

  $result = Show-Modal "Confirm Delete" "Delete zone '$ZoneName'?`n`nThis will delete all records in the zone!" -YesNo

  if ($result -ne 0) { return }

  try {
    Remove-DnsServerZone -Name $ZoneName -ComputerName $global:DetectedDnsServer -Force -ErrorAction Stop
    Debug-Log "Zone deleted: $ZoneName" -Type "Success"
    Show-Modal "Success" "Zone deleted successfully!"
    Show-DNSDialog
  } catch {
    Debug-Log "Failed to delete zone: $_" -Type "Error"
    Show-Modal "Error" "Failed to delete zone: $_"
  }
}

function Show-Tools {
  Debug-Log "Opening Diagnostic Tools dialog" -Type "Tracing"

  ## Capture functions for closures
  $showModalFunc = ${function:Show-Modal}
  $debugLogFunc = ${function:Debug-Log}
  $formatRecordFunc = ${function:Format-RecordData}

  ## Create dialog
  $dialog = [Terminal.Gui.Dialog]::new()
  $dialog.Title = "DNS Diagnostic Tools"
  $dialog.Width = 100
  $dialog.Height = 35

  ## ---------------- Target Field ----------------
  $y = 1

  $lblTarget = [Terminal.Gui.Label]::new("Target:")
  $lblTarget.X = 2
  $lblTarget.Y = $y
  $dialog.Add($lblTarget)

  $txtTarget = [Terminal.Gui.TextField]::new("")
  $txtTarget.X = 12
  $txtTarget.Y = $y
  $txtTarget.Width = 50
  $dialog.Add($txtTarget)

  $lblHelp = [Terminal.Gui.Label]::new("(hostname or IP address)")
  $lblHelp.X = 63
  $lblHelp.Y = $y
  $dialog.Add($lblHelp)

  $y += 2

  ## ---------------- Tool Buttons (Row 1) ----------------
  $btnPing = [Terminal.Gui.Button]::new("Ping")
  $btnPing.X = 2
  $btnPing.Y = $y
  $btnPing.Width = 12
  $dialog.Add($btnPing)

  $btnNslookup = [Terminal.Gui.Button]::new("Nslookup")
  $btnNslookup.X = 15
  $btnNslookup.Y = $y
  $btnNslookup.Width = 12
  $dialog.Add($btnNslookup)

  $btnResolve = [Terminal.Gui.Button]::new("Resolve")
  $btnResolve.X = 28
  $btnResolve.Y = $y
  $btnResolve.Width = 12
  $dialog.Add($btnResolve)

  $btnTraceroute = [Terminal.Gui.Button]::new("Traceroute")
  $btnTraceroute.X = 41
  $btnTraceroute.Y = $y
  $btnTraceroute.Width = 14
  $dialog.Add($btnTraceroute)

  $y += 2

  ## ---------------- Tool Buttons (Row 2) ----------------
  $btnDNSCache = [Terminal.Gui.Button]::new("DNS Cache")
  $btnDNSCache.X = 2
  $btnDNSCache.Y = $y
  $btnDNSCache.Width = 12
  $dialog.Add($btnDNSCache)

  $btnClearCache = [Terminal.Gui.Button]::new("Clear Cache")
  $btnClearCache.X = 15
  $btnClearCache.Y = $y
  $btnClearCache.Width = 14
  $dialog.Add($btnClearCache)

  $btnBenchmark = [Terminal.Gui.Button]::new("Benchmark")
  $btnBenchmark.X = 30
  $btnBenchmark.Y = $y
  $btnBenchmark.Width = 12
  $dialog.Add($btnBenchmark)

  $btnClear = [Terminal.Gui.Button]::new("Clear Output")
  $btnClear.X = 43
  $btnClear.Y = $y
  $btnClear.Width = 14
  $dialog.Add($btnClear)

  $y += 2

  ## ---------------- Output View ----------------
  $lblOutput = [Terminal.Gui.Label]::new("Results:")
  $lblOutput.X = 2
  $lblOutput.Y = $y
  $dialog.Add($lblOutput)
  $y += 1

  $outputView = [Terminal.Gui.TextView]::new()
  $outputView.X = 2
  $outputView.Y = $y
  $outputView.Width = [Terminal.Gui.Dim]::Fill(2)
  $outputView.Height = [Terminal.Gui.Dim]::Fill(3)
  $outputView.ReadOnly = $true
  $outputView.WordWrap = $false
  $outputView.Text = "Select a diagnostic tool above and enter a target to begin..."
  $dialog.Add($outputView)

  ## ---------------- Helper Function for Running Tools ----------------
  function Run-Tool {
    param(
      [string]$Tool,
      [string]$Target,
      [object]$OutputView
    )

    if ([string]::IsNullOrWhiteSpace($Target) -and $Tool -notin @("DNSCache", "ClearCache")) {
      & $showModalFunc "Error" "Please enter a target"
      return
    }

    $output = "=== $Tool $(if ($Target) { "- $Target" }) ===`n"
    $output += "Time: $(Get-Date -Format 'HH:mm:ss')`n"
    $output += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n"

    try {
      switch ($Tool) {
        "Ping" {
          $results = Test-Connection -ComputerName $Target -Count 4 -ErrorAction Stop
          foreach ($result in $results) {
            $output += "Reply from $($result.Address): Time=$($result.ResponseTime)ms TTL=$($result.TimeToLive)`n"
          }
        }
        "Nslookup" {
          $results = Resolve-DnsName -Name $Target -Server $global:DetectedDnsServer -ErrorAction Stop
          foreach ($result in $results) {
            $output += "Name: $($result.Name) Type: $($result.Type) Data: $(& $formatRecordFunc -record $result)`n"
          }
        }
        "Resolve" {
          $results = Resolve-DnsName -Name $Target -ErrorAction Stop
          foreach ($result in $results) {
            $output += "$($result.Type): $($result.Name) -> $(& $formatRecordFunc -record $result)`n"
          }
        }
        "Traceroute" {
          $output += "Tracing route to $Target...`n"
          $results = Test-NetConnection -ComputerName $Target -TraceRoute -ErrorAction Stop
          if ($results.TraceRoute) {
            $hop = 1
            foreach ($addr in $results.TraceRoute) {
              $output += "$hop  $addr`n"
              $hop++
            }
          }
          $output += "`nTrace complete.`n"
        }
        "DNSCache" {
          $cache = Get-DnsClientCache -ErrorAction Stop
          $output += "DNS Client Cache ($($cache.Count) entries):`n`n"
          $maxDisplay = 50
          foreach ($entry in $cache | Select-Object -First $maxDisplay) {
            $output += "Name: $($entry.Name) Type: $($entry.Type) Data: $($entry.Data) TTL: $($entry.TimeToLive)`n"
          }
          if ($cache.Count -gt $maxDisplay) {
            $output += "`n... and $($cache.Count - $maxDisplay) more entries`n"
          }
        }
        "ClearCache" {
          Clear-DnsClientCache -ErrorAction Stop
          $output += "DNS client cache cleared successfully.`n"
        }
        "Benchmark" {
          $output += "Running DNS benchmark for $Target...`n"
          $iterations = 10
          $times = @()
          for ($i = 1; $i -le $iterations; $i++) {
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            Resolve-DnsName -Name $Target -Server $global:DetectedDnsServer -ErrorAction SilentlyContinue | Out-Null
            $sw.Stop()
            $times += $sw.ElapsedMilliseconds
            $output += "Test $i : $($sw.ElapsedMilliseconds)ms`n"
          }
          $avg = ($times | Measure-Object -Average).Average
          $min = ($times | Measure-Object -Minimum).Minimum
          $max = ($times | Measure-Object -Maximum).Maximum

          $output += "`nResults:`n"
          $output += "Average: $([math]::Round($avg, 2))ms`n"
          $output += "Minimum: $min ms`n"
          $output += "Maximum: $max ms`n"
        }
      }
    } catch {
      $output += "ERROR: $_`n"
      & $debugLogFunc "Diagnostic tool $Tool failed: $_" -Type "Error"
    }

    $output += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n`n"

    # Append to output
    $currentText = $OutputView.Text.ToString()
    $OutputView.Text = $currentText + $output

    & $debugLogFunc "Diagnostic tool $Tool completed" -Type "Success"
  }

  ## ---------------- Button Click Handlers ----------------
  $btnPing.add_Clicked({
    Run-Tool -Tool "Ping" -Target $txtTarget.Text.ToString() -OutputView $outputView
  }.GetNewClosure())

  $btnNslookup.add_Clicked({
    Run-Tool -Tool "Nslookup" -Target $txtTarget.Text.ToString() -OutputView $outputView
  }.GetNewClosure())

  $btnResolve.add_Clicked({
    Run-Tool -Tool "Resolve" -Target $txtTarget.Text.ToString() -OutputView $outputView
  }.GetNewClosure())

  $btnTraceroute.add_Clicked({
    Run-Tool -Tool "Traceroute" -Target $txtTarget.Text.ToString() -OutputView $outputView
  }.GetNewClosure())

  $btnDNSCache.add_Clicked({
    Run-Tool -Tool "DNSCache" -Target "" -OutputView $outputView
  }.GetNewClosure())

  $btnClearCache.add_Clicked({
    Run-Tool -Tool "ClearCache" -Target "" -OutputView $outputView
  }.GetNewClosure())

  $btnBenchmark.add_Clicked({
    Run-Tool -Tool "Benchmark" -Target $txtTarget.Text.ToString() -OutputView $outputView
  }.GetNewClosure())

  $btnClear.add_Clicked({
    $outputView.Text = "Output cleared. Select a tool to begin..."
    & $debugLogFunc "Diagnostic output cleared" -Type "Tracing"
  }.GetNewClosure())

  ## ---------------- Close Button ----------------
  $btnClose = [Terminal.Gui.Button]::new("Close")
  $btnClose.add_Clicked({
    & $debugLogFunc "Diagnostic Tools dialog closed" -Type "Tracing"
    [Terminal.Gui.Application]::RequestStop()
  }.GetNewClosure())
  $dialog.AddButton($btnClose)

  ## Run dialog
  [Terminal.Gui.Application]::Run($dialog)
}

function Run-DiagnosticTool {
  param(
    [string]$Tool,
    [string]$Target
  )

  if ([string]::IsNullOrWhiteSpace($Target) -and $Tool -notin @("DNSCache", "ClearCache")) {
    Show-Modal "Error" "Please enter a target"
    return
  }

  $output = "=== $Tool $(if ($Target) { "- $Target" }) ===`n"

  try {
    switch ($Tool) {
      "Ping" {
        $results = Test-Connection -ComputerName $Target -Count 4 -ErrorAction Stop
        foreach ($result in $results) { $output += "Reply from $($result.Address): Time=$($result.ResponseTime)ms TTL=$($result.TimeToLive)`n"}
      }
      "Nslookup" {
        $results = Resolve-DnsName -Name $Target -Server $global:DetectedDnsServer -ErrorAction Stop
        foreach ($result in $results) { $output += "Name: $($result.Name) Type: $($result.Type) Data: $(Format-RecordData -record $result)`n"  }
      }
      "Resolve" {
        $results = Resolve-DnsName -Name $Target -ErrorAction Stop
        foreach ($result in $results) { $output += "$($result.Type): $($result.Name) -> $(Format-RecordData -record $result)`n" }
      }
      "Traceroute" {
        $output += "Tracing route to $Target...`n"
        $results = Test-NetConnection -ComputerName $Target -TraceRoute -ErrorAction Stop
        if ($results.TraceRoute) {
          $hop = 1
          foreach ($addr in $results.TraceRoute) {
            $output += "$hop  $addr`n"
            $hop++
          }
        }
        $output += "`nTrace complete.`n"
      }
      "DNSCache" {
        $cache = Get-DnsClientCache -ErrorAction Stop
        $output += "DNS Client Cache ($($cache.Count) entries):`n`n"
        foreach ($entry in $cache | Select-Object -First $global:AppConfig.MaxCacheDisplay) { $output += "Name: $($entry.Name) Type: $($entry.Type) Data: $($entry.Data) TTL: $($entry.TimeToLive)`n" }
          if ($cache.Count -gt $global:AppConfig.MaxCacheDisplay) { $output += "`n... and $($cache.Count - $global:AppConfig.MaxCacheDisplay) more entries`n" }
      }
      "ClearCache" {
        Clear-DnsClientCache -ErrorAction Stop
        $output += "DNS client cache cleared successfully.`n"
      }
      "Benchmark" {
        $output += "Running DNS benchmark for $Target...`n"
        $iterations = 10
        $times = @()
        for ($i = 1; $i -le $iterations; $i++) {
          $sw = [System.Diagnostics.Stopwatch]::StartNew()
          Resolve-DnsName -Name $Target -Server $global:DetectedDnsServer -ErrorAction SilentlyContinue | Out-Null
          $sw.Stop()
          $times += $sw.ElapsedMilliseconds
          $output += "Test $i : $($sw.ElapsedMilliseconds)ms`n"
        }
        $avg = ($times | Measure-Object -Average).Average
        $min = ($times | Measure-Object -Minimum).Minimum
        $max = ($times | Measure-Object -Maximum).Maximum

        $output += "`nResults:`n"
        $output += "Average: $([math]::Round($avg, 2))ms`n"
        $output += "Minimum: $min ms`n"
        $output += "Maximum: $max ms`n"
      }
    }
  } catch {
    $output += "Error: $_`n"
  }
  $output += "`n"
  if ($global:DiagnosticOutputView) {
    $currentText = $global:DiagnosticOutputView.Text.ToString()
    $global:DiagnosticOutputView.Text = $currentText + $output
  }
}

## ----------------------------{ Main Execution }----------------------------
## Apply theme globally
Apply-Theme -ThemeData $themeData -TopLevel $top -MainWindow $script:win -Menu $script:menu -Status $script:StatusBar
if ($themeData -and $themeData.Global) {
  $script:win.ColorScheme      = $themeData.MainWindow
  $script:menu.ColorScheme     = $themeData.Global
  $script:StatusBar.ColorScheme= $themeData.Global
}

## Initialize paths
$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }

## ----------{ Program Launch starts Here }----------
## Echo basic info for debugging
Debug-Log "DemoMode: $DemoMode" -Type "Insight"
Debug-Log "Logging: $Logging" -Type "Insight"
Debug-Log "LogFile: $LogFile" -Type "Insight"

## Initialise logging if requested
if ($Logging -or $LogFile) {
  Debug-Log "Logging condition TRUE" -Type "Insight"
  if (-not $LogFile) {
    $LogFile = Join-Path $PSScriptRoot "dnstui_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
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
    $Script:LogStream.WriteLine("========== DNS-TUI Log Started $(Get-Date) ==========")
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

## DNS-specific global variables
$global:PerformanceCounters = @{
  OperationCount      = 0
  ErrorCount          = 0
  LastOperationTime   = $null
  AverageResponseTime = 0
}

$global:DNSConnectionStatus = @{
  IsConnected       = $false
  LastChecked       = $null
  ServerName        = ""
  CacheValidSeconds = 30
}

$global:AutoRefresh = @{
  Enabled         = $false
  Timer           = $null
  IntervalSeconds = 300
  CurrentView     = $null
}
Debug-Log "Starting $($Script:ProjectName) - ${Global:BuildVersion} Codename: ($Script:FruitName) with theme: ($Script:ThemeMode)" -Type "Insight"

## -------------------------------{ App }---------------------------------
Debug-Log "Starting $($Script:ProjectName)..." -Type "Insight"

# STEP 1: Check requirements
$Script:hasConsoleTools = Test-Requirement -Type Module -Name 'Microsoft.PowerShell.ConsoleGuiTools' -InstallMsg 'Install-Module -Name Microsoft.PowerShell.ConsoleGuiTools -RequiredVersion 0.7.2'

## ----------{ Preflight Checks }----------
Debug-Log "Performing pre-flight module checks..." -Type "Tracing"

## Check all modules ONCE at startup
$Script:hasConsoleTools    = Test-Requirement -Type Module -Name 'Microsoft.PowerShell.ConsoleGuiTools' -InstallMsg 'Install-Module -Name Microsoft.PowerShell.ConsoleGuiTools -RequiredVersion 0.7.2'
$Script:HasPSWriteColor    = Test-Requirement -Type Module -Name "PSWriteColor" -InstallMsg 'Install-Module -Name PSWriteColor' -Optional
$Script:HasTerminalIcons   = Test-Requirement -Type Module -Name "Terminal-Icons" -InstallMsg 'Install-Module -Name Terminal-Icons' -Optional
$Script:hasNerdFonts       = Test-Requirement -Type Module -Name 'NerdFonts' -InstallMsg 'Install-Module -Name NerdFonts' -Optional
$Script:hasDNSClient       = Test-Requirement -Type Module -Name 'DNSClient' -InstallMsg 'Install-Module -Name DNSClient'
$Script:HasDNSServer       = Test-Requirement -Type WindowsCapability -Name "Rsat.Dns.Tools~~~~0.0.1.0" -InstallMsg 'Add-WindowsCapability -Online -Name Rsat.Dns.Tools~~~~0.0.1.0'

## Optional insight/log
Debug-Log "To install all RSAT features (may be overkill): 'Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability -Online'" -Type "Insight"

## Initialise icon set based on available modules
Initialise-Icons
Initialise-DirectoryEmoji

if ($Script:LaunchReady) {
  Debug-Log "Only checking of script prerequisites requested. Now exiting..." -Type "Insight"
  return
}

Debug-Log "Module availability check complete" -Type "Insight"
$Script:UseIcons = $false
if ($Script:HasTerminalIcons) { try { Write-Host '' -NoNewline; $Script:UseIcons = $true } catch {} }

# STEP 2: Initialize UI Framework
Debug-Log "Initializing Terminal.Gui..." -Type "Insight"
$uiComponents = Initialise-UIFramework -Theme $Script:ThemeMode -Title "$Script:Projectname - $($Script:BuildVersion) $Script:FruitName $Script:Codename $Script:FruitEmoji"

## STEP 3: Set global variables
$script:top = $uiComponents.Top
$script:win = $uiComponents.Window
$script:menu = $uiComponents.Menu
$script:StatusBar = $uiComponents.StatusBar
Debug-Log "UI Framework ready" -Type "Insight"

## STEP 4: Auto-detect and connect to DNS server
$global:DNSDetection = Get-DNSServerDetection
$global:DetectedDnsServer = $global:DNSDetection.DnsServerName
if ($global:DNSDetection.AutoConnect -and $global:DNSDetection.IsLocalDNS) {
  Debug-Log "Auto-connecting to local DNS server: $($global:DetectedDnsServer)" -Type "Insight"
  try {
    Connect-ToDNSServer -ServerName $global:DetectedDnsServer
  } catch {
    Debug-Log "Auto-connect failed: $_" -Type "Warning"
  }
}

## STEP 5: Load dashboard content (populates right panel and builds tree)
Set-StatusBar "Refreshing users" -Icon Working -Percent 25
Show-Dashboard
Set-StatusBar "Populating tree..." -Icon 'Working' -Percent 80
## Debug view tree dump
if ($DebugMode -or $Logging) {
  Debug-Log "========== Full View Tree Dump ==========" -Type "Tracing"
  Debug-DumpViewTree -View $top
  Debug-Log "========== End View Tree Dump ==========" -Type "Tracing"
}
Set-StatusBar "Tree built successfully" -Icon 'Success' -Percent 90

## Capture functions for closure
$debugLogFunc     = ${function:Debug-Log}
$buildContextFunc = ${function:Build-ContextMenuItems}
$showContextFunc  = ${function:Show-ContextMenu}

## STEP 6: Verify and run
Debug-Log "Starting Terminal.Gui main loop..." -Type "Success"
$finalTop = [Terminal.Gui.Application]::Top
if ($null -eq $finalTop) {
  Debug-Log "FATAL ERROR: Application.Top is null!" -Type "Error"
  exit 1
}

Set-StatusBar "Ready" -Icon 'Success'
Debug-Log "Starting application with $($finalTop.Subviews.Count) top-level views..." -Type "Insight"
[Terminal.Gui.Application]::Run()

## ----------{ Cleanup }----------
Debug-Log "Application stopped, cleaning up..." -Type "Tracing"
Set-StatusBar "Shutting down"
[Terminal.Gui.Application]::Shutdown()
Debug-Log "Application shut down cleanly" -Type "Success"
Debug-Log "End of line..." -Type "Insight"

if ($Script:LogStream) {
  $Script:LogStream.Close()
  $Script:LogStream.Dispose()
}
