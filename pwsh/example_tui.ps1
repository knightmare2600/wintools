<#

Header info about project 

===========================================================================================
 Name of project (Danish Fruit name) — Brief description
 Historical Build Notes and Change Log
===========================================================================================

 1.0.0  (Initial Experimental)
 - First internal test build. Basic TUI scaffolding only.
 1.0.1  (Bugfix)
 - Removed every zig for great justice

===========================================================================================
#>

param(
  [switch]$DemoMode,
  [switch]$Logging,
  [string]$Domain,
  [ValidateSet("light","dark","matrix","british")]
  [string]$Theme = "dark"
)

## Define the build version, project name and fruit codename once only
$BuildVersion = "1.0.0"
$ProjectName = "My First TUI Pwsh App"
$FruitName = "Dansk Frugt Namen"

Write-Host "Starting ${ProjectName} v ${BuildVersion} in $(if($DemoMode){'DEMO'}else{'PRODUCTION'}) mode with ${Theme} theme..."

## -------------------------{ Load Terminal.Gui }------------------------
Write-Host "Checking Terminal.Gui assembly..."
if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Terminal.Gui' })) {
    $mod = Get-Module -ListAvailable Microsoft.PowerShell.ConsoleGuiTools | Select-Object -First 1
    if ($mod) {
        $dll = Join-Path $mod.ModuleBase 'Terminal.Gui.dll'
        if (Test-Path $dll) { Add-Type -Path $dll -ErrorAction Stop; Write-Host "Loaded Terminal.Gui from $dll" } 
        else { Write-Error "Terminal.Gui.dll not found. Install Microsoft.PowerShell.ConsoleGuiTools."; return }
    } else { Write-Error "Microsoft.PowerShell.ConsoleGuiTools module not found."; return }
} else { Write-Host "Terminal.Gui assembly already loaded." }

## -------------------------{ Example Globals }------------------------
#if (-not $Domain) { $Domain = if ($DemoMode) { "example.com" } else { (Get-ADDomain).Forest } }
##$Global:Domain = $Domain
#$Global:CurrentDC = $null
#$Global:Users = @()
#$Global:DCs = @()
#$Global:ADObjects = @()  # New for production AD object detection
#$Global:SelectedObjects = @()
#$Global:SelectionMode = $false

## Set global demo mode flag immediately
$Global:DemoMode = $DemoMode

## Set global flags immediately after param block - themes here
$Global:DemoMode = $DemoMode
$script:ThemeMode = $Theme

Write-Host "Starting ${ProjectName} in $(if($DemoMode){'DEMO'}else{'PRODUCTION'}) mode..."

## -------------------------{ Theme Definitions }------------------------
function Get-Theme {
  param([string]$mode)

  ## Initialize color schemes and Ensure ColorSchemes are instantiated
  if (-not $globalCs)     { $globalCs     = [Terminal.Gui.ColorScheme]::new() }
  if (-not $mainWindowCs) { $mainWindowCs = [Terminal.Gui.ColorScheme]::new() }

  ## Normalize theme string: lowercase + ASCII
  $mode = $mode.Trim().ToLower()
  
 <#
 Adding Themes:

 Add an option above in the [ValidateSet() then define a theme below:

        "faxekondi" {
           $globalCs.Normal     <-- Foreground borders and background colour for all modals
           $globalCs.Focus      <-- Foreground and background for menus
           $mainWindowCs.Normal <-- Main opening dialog and foreground text colour
           $mainWindowCs.Focus  <-- Main opening window focus colours foreground nad background
        }
 #>

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
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::Blue)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::White)
      }

    default {
      ## fallback to dark
      $globalCs.Normal     = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Black)
      $globalCs.Focus      = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::Gray)
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Black)
      $mainWindowCs.Focus  = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Black,[Terminal.Gui.Color]::DarkGray)
      }
  }

  ## Ensure HotNormal/HotFocus
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

  ## --- Global / TopLevel ---
  if ($TopLevel -and $TopLevel.PSObject.Properties.Name -contains 'ColorScheme') {
    $TopLevel.ColorScheme = $ThemeData.Global
  }

  ## --- Main window ---
  if ($MainWindow -and $MainWindow.PSObject.Properties.Name -contains 'ColorScheme') {
    $MainWindow.ColorScheme = $ThemeData.MainWindow
  }

  ## --- Menu ---
  if ($Menu -and $Menu.PSObject.Properties.Name -contains 'ColorScheme') {
    $Menu.ColorScheme = $ThemeData.Global
  }

  # --- StatusBar ---
  if ($Status -and $Status.PSObject.Properties.Name -contains 'ColorScheme') {
    $Status.ColorScheme = $ThemeData.Global
  }

  ## --- Terminal.Gui base colors ---
  [Terminal.Gui.Colors]::Base     = $ThemeData.Global
  [Terminal.Gui.Colors]::Dialog   = $ThemeData.Global
  [Terminal.Gui.Colors]::Menu     = $ThemeData.Global
  [Terminal.Gui.Colors]::Error    = $ThemeData.Global
  [Terminal.Gui.Colors]::TopLevel = $ThemeData.Global
}

## Diagnostics helper to show what's inside a ColorScheme
function Dump-ColorScheme {
  param([Terminal.Gui.ColorScheme]$Scheme)
  if ($null -eq $Scheme) { Write-Output "ColorScheme is null"; return }
    Write-Output "Normal    : $($Scheme.Normal)"
    Write-Output "Focus     : $($Scheme.Focus)"
    Write-Output "HotNormal : $($Scheme.HotNormal)"
    Write-Output "HotFocus  : $($Scheme.HotFocus)"
    Write-Output "Disabled  : $($Scheme.Disabled)"
}

## Select theme before proceeding. Save the mode string
$script:ThemeMode = $Theme

## Get the selected colour scheme
$cs = Get-Theme -mode $Theme
Apply-Theme -ThemeData $themeData -TopLevel $TopLevel -MainWindow $MainWindow -Menu $Menu -Status $Status

## ---{ Helper: Show a simple loading/progress dialog with spinner }---
function Show-LoadingDialog {
  param(
    [string]$Message = "Loading, please wait..."
  )

  ## Create dialog and label
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

  ## Spinner frames and timer setup - For those old enough, this is from Vulnhub circa 2014
  $frames = @("|", "/", "-", "\")
  $i = 0
  $timer = [System.Threading.Timer]::new( {
    $global:spinnerFrameIndex = ($global:spinnerFrameIndex + 1) % 4
    [Terminal.Gui.Application]::MainLoop.Invoke({
      $spinner.Text = $frames[$global:spinnerFrameIndex]
      })
  },
  $null, 0, 150
  )

  ## Start non-blocking dialog
  [Terminal.Gui.Application]::Begin($dlg)

  ## Return both dialog and timer so caller can close/stop cleanly
  return [PSCustomObject]@{ Dialog = $dlg; Timer = $timer }
}

## ---- Helper: Close loading/progress dialog ----
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
     Write-Host "Switching to theme: $selectedTheme"
     $script:ThemeMode = $selectedTheme
        
     ## Get new theme
     $newTheme = Get-Theme -mode $selectedTheme
        
     ## Apply to all components
     Apply-Theme -ThemeData $newTheme -TopLevel $top -MainWindow $win -Menu $menu -Status $status
        
     ## Refresh display
     [Terminal.Gui.Application]::Refresh()
        
     [Terminal.Gui.MessageBox]::Query(50, 7, "Theme Changed", "Theme changed to: $selectedTheme", "OK") | Out-Null
     [Terminal.Gui.Application]::RequestStop()
  })
  $dlg.AddButton($btnApply)
    
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $dlg.AddButton($btnCancel)
    
  [Terminal.Gui.Application]::Run($dlg)
}

## ---------------------{ Initialize Terminal.Gui }--------------------
[Terminal.Gui.Application]::Init()
$top = [Terminal.Gui.Application]::Top

## Get the selected theme
Write-Host "Applying theme: $Theme"
$themeData = Get-Theme -mode $Theme

## Apply theme to top level first
if ($themeData -and $themeData.Global) {
  $top.ColorScheme = $themeData.Global
}

## ------------------------- Main Window ------------------------
$win = [Terminal.Gui.Window]::new("ProjectName — Some Project ${BuildVersion} ${FruitName}")
$win.X=0; $win.Y=0; $win.Width=[Terminal.Gui.Dim]::Fill(); $win.Height=[Terminal.Gui.Dim]::Fill()

## Apply theme to main window
if ($themeData -and $themeData.MainWindow) {
  $win.ColorScheme = $themeData.MainWindow
}

$top.Add($win)

Write-Host "Theme applied successfully: $Theme"

## <-- Adding things such as filter panes, and so on, to the main window goes here -->

## ---------------------------{ Status Bar }---------------------------
$status = [Terminal.Gui.StatusBar]::new(@(
  [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F1,"~F1~ Help",{
    Show-Modal "Shortcuts" "F1 - Help`nF9 -Do Something Function`nF10 - Delete`nF10 - Quit`nF12 Redraw" }),
  [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F9,"~F9~ New",{ Show-SomeThingFunction }),
  [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F10,"~F10~ Quit",{ [Terminal.Gui.Application]::RequestStop() }),
  [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F12,"~F12~ Redraw",{ [Terminal.Gui.Application]::Refresh() })
))
$top.Add($status)

## ------------------------- Menu ------------------------
# Existing menu items
$mFile = [Terminal.Gui.MenuItem]::new("_Exit","Exit application",[Action]{ [Terminal.Gui.Application]::RequestStop() })
$mNew = [Terminal.Gui.MenuItem]::new("New Thing","Create a new thing",[Action]{ Show-NewThingWizard })
$mProps = [Terminal.Gui.MenuItem]::new("_Properties","Edit thing properties",[Action]{ Show-ThingProperties })
$mUndo = [Terminal.Gui.MenuItem]::new("_Undo","Undo last action",[Action]{ Write-Host "DEBUG: Undo placeholder" })

## ---{ About Menu Items }---
$mShortcuts = [Terminal.Gui.MenuItem]::new(
  "_Shortcuts", "Keyboard shortcuts",
  [Action]{
    Show-Modal "Shortcuts" "F1 - Help`nF9 -New`nF10 - Quit`nF11 -Redraw"
  }
)

$mAboutProject = [Terminal.Gui.MenuItem]::new(
  "_About ${ProjectName}",
  "About this application",
  [Action]{ 
    [Terminal.Gui.MessageBox]::Query("About","${ProjectName} ${BuildVersion} ${FruitName}`n© 1995 Copyleft (GPL-3)`nDemo Mode: $DemoMode","OK") | Out-Null 
  }
)

## A small modal about the fruit based names
$mWhyFruitName = [Terminal.Gui.MenuItem]::new(
  "Why _FruitName?",
  "Why the Fruit Name codename?",
  [Action]{ Show-FruitNameInfo }
)

## --- NEW: About Top-Level Menu ---
$aboutMenu = [Terminal.Gui.MenuBarItem]::new("_About", @(
  $mShortcuts,
  $mAboutProject,
  $mWhyFruitName
))

## Mandatory (BEFORE creating $menu):
$mTheme = [Terminal.Gui.MenuItem]::new("_Theme","Change color theme",[Action]{ Show-ThemeSelector })

## -------------------------{ Fixed Menu Bar }-------------------------
$menu = [Terminal.Gui.MenuBar]::new(@(
  ## FILE menu
  [Terminal.Gui.MenuBarItem]::new("_File", @(
    $mRefresh,
    $mTheme,
    $mFile
    )),

  ## ACTION menu
  [Terminal.Gui.MenuBarItem]::new("_Action", @(
    $mNew,
    $mProps,
    $mUndo
    )),

  ## ABOUT menu
  $aboutMenu
  ))

## Apply full theme to all components <-- do this BEFORE the menus
Apply-Theme -ThemeData $themeData -TopLevel $top -MainWindow $win -Menu $menu -Status $status
Write-Host "Theme applied successfully: $Theme"

$top.Add($menu)

## -------------------------{ Render Tools Row }------------------------
## Render tools row - show whether or not ancilory tols such as openssl
## or wget are found - if required. Tools change based on requirements
function Render-ToolsRow {
  $toolsView.RemoveAll()
  $offset = 0
  $tools = $script:Tools
  $items = @(
    @{ Name="more"; Found = $tools.more },
    @{ Name="curl"; Found = $tools.curl },
    @{ Name="whois"; Found = $tools.whois },
    @{ Name="ping"; Found = $tools.ping },
    @{ Name="arp"; Found = $tools.arp }
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
          [Terminal.Gui.MessageBox]::Query(50,7,"Tool Found","$nameCopy detected - will be used where applicable.")
        } else {
          [Terminal.Gui.MessageBox]::Query(60,8,"Tool Missing","$nameCopy not found. Some detail modes will be disabled or fall back to PowerShell cmdlets.")
          }
        })
      $toolsView.Add($lbl)
      $offset += ($lbl.Text.Length + 1)
  }
}

## -----------------{ Generate Random Password modal }-----------------
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

  $script:actualPassword = ""

  ## --- Build UI ---
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

  ## ---{ Entropy Display }---
  $lblStrength = [Terminal.Gui.Label]::new(30,7,"Strength: Not generated")
  $lblStrength.Width = 30
  $dlg.Add($lblStrength)

  $progEntropy = [Terminal.Gui.ProgressBar]::new()
  $progEntropy.X = 30; $progEntropy.Y = 8; $progEntropy.Width = 25
  $progEntropy.Fraction = 0.0
  $dlg.Add($progEntropy)

  ## Show Password checkbox
  $chkShowPwd = [Terminal.Gui.CheckBox]::new(30,5,"Show Password",$false)
  $dlg.Add($chkShowPwd)

  ## Buttons
  $btnGenerate = [Terminal.Gui.Button]::new("Generate"); $btnGenerate.X=2; $btnGenerate.Y=10
  $btnCopy     = [Terminal.Gui.Button]::new("Copy");     $btnCopy.X=15; $btnCopy.Y=10
  $btnClose    = [Terminal.Gui.Button]::new("Close");    $btnClose.X=28; $btnClose.Y=10
  $dlg.Add($btnGenerate, $btnCopy, $btnClose)

  ## ---{ Generate Logic }---
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

      ## ---{ Entropy update }---
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

  ## --- Show Password toggle ---
  $chkShowPwd.add_Toggled({
    if ($chkShowPwd.Checked) {
      $txtPwd.Text = $script:actualPassword
    } else {
      $txtPwd.Text = ('*' * $script:actualPassword.Length)
    }
  })

  ## --- Copy to Clipboard ---
  $btnCopy.add_Clicked({
     if (-not $script:actualPassword) { return }
     if ($IsWindows) { Set-Clipboard -Value $script:actualPassword }
     elseif ($IsMacOS) { $script:actualPassword | pbcopy }
     else { $script:actualPassword | xsel --clipboard --input }
     [Terminal.Gui.MessageBox]::Query(50,7,"Copied","Password copied to clipboard.","OK") | Out-Null
  })

  ## Close
  $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })

  [Terminal.Gui.Application]::Run($dlg)
  return $script:actualPassword
}

## -------------------------{ Example Globals }------------------------

## This is a theme now. Danish Fruit soda based fun. Method to the
## madness
function Show-BlaabaerInfo {
  $dlg = [Terminal.Gui.Dialog]::new("Why Blaabaer? 🫐", 60, 12)
    
  $message = @"
${ProjectName} is codenamed "Fruitname" because:

- Reason 1
- 'Fruit Name' (Danish Speling) is Danish for Fruit Name
- Reason 2
- Every great project needs a forest-fruit mascot!
"@
    
  $label = [Terminal.Gui.Label]::new(1, 1, $message)
  $dlg.Add($label)
    
  [Terminal.Gui.MessageBox]::Query("Why Blaabaer? 🫐", $message, @("OK"))
}

## ------------------------{ Show message boxes }----------------------
## This cuts down significantly on code re-use
function Show-Modal { 
  param($title, $msg) 
  [Terminal.Gui.MessageBox]::Query($title, $msg, @("OK")) | Out-Null 
}

## Call the initial function to start things off
Show-Modal "Hello there" "Example PowerShell TUI App using ConsoleTools 1.16 Module"


## ------------------------{ Run Applicaiton }-----------------------
[Terminal.Gui.Application]::Run($top)
[Terminal.Gui.Application]::Shutdown()
