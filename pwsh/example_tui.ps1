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
1.1.0  (Add tools feature)
- Added tools modal
- Added Function keys to menus and status bar
1.2.4  (Add themes)
- Cleaned up tools modal
- Corrected function keys in menus and status bar
- Add command line options for DemoMode, Logging and Verbose
- Global Debug-Log and Show-Modal functions ot significantly cut down on code
1.2.8 (Globals)
 - Make use of global variables and use accordingly in dialogs and modals
 - Add codename Global and use accordingly
===========================================================================================
#>

param(
  [switch]$DemoMode,
  [switch]$Logging,
  [string]$Verbose,
  [ValidateSet("light","dark","matrix","british")]
  [string]$Theme = "matrix"
)

## -------------------------{ Load Terminal.Gui }-------------------------
Write-Host "Checking Terminal.Gui assembly..."
if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Terminal.Gui' })) {
  $mod = Get-Module -ListAvailable Microsoft.PowerShell.ConsoleGuiTools | Select-Object -First 1
  if ($mod) {
    $dll = Join-Path $mod.ModuleBase 'Terminal.Gui.dll'
    if (Test-Path $dll) { Add-Type -Path $dll -ErrorAction Stop; Write-Host "Loaded Terminal.Gui from $dll" } 
      else { Write-Error "Terminal.Gui.dll not found. Install module."; return }
  } else { Write-Error "Microsoft.PowerShell.ConsoleGuiTools module not found."; return }
} else { Write-Host "Terminal.Gui assembly already loaded." }

## -------------------------{ Global Variables }--------------------------
## Define the build version, project name and fruit codename once only
$Global:ProjectName = "Demo TUI Powershell App"
$Global:FruitName = "Dansk Frugt Namen"
$Global:BuildVersion = "1.2.8"

## Set global flags immediately after param block - themes here
$Global:DemoMode = $DemoMode
$Global:ThemeMode = $Theme

Write-Host "Starting $($Global:ProjectName) - ${Global:BuildVersion} Codename: ($Global:FruitName) with theme: ($Global:ThemeMode)"

## -------------------------------{ App }---------------------------------
[Terminal.Gui.Application]::Init()
[Terminal.Gui.Application]::Top

## Root window
$win = [Terminal.Gui.Window]::new("${ProjectName} — ${BuildVersion} ${FruitName}")
$win.X      = 0
$win.Y      = 1
$win.Width  = [Terminal.Gui.Dim]::Fill()
$win.Height = [Terminal.Gui.Dim]::Fill() - 1    # ← leave room for bottom tools bar

## --------------------------{ Core Functions }----------------------------
## - All core functions needed in menus, pop-ups, modals, etc. go in here -
## ------------------------------------------------------------------------

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

## --------------------------{ Debug Logging }-------------------------
function Debug-Log {
  param([string]$Message)
    if ($Verbose) {
      $ts = (Get-Date).ToString('HH:mm:ss')
      Write-Host "[$ts] LOG: $Message" -ForegroundColor Cyan
    }
}

## ---------------------{ Pretty Theme Selections }--------------------
function Show-ThemeSelector {
  $dlg = [Terminal.Gui.Dialog]::new("Select Theme", 50, 14)
  $lbl = [Terminal.Gui.Label]::new("Choose a color theme:"); $lbl.X=2; $lbl.Y=1; $dlg.Add($lbl)
  
  $themes = @("light", "dark", "matrix", "british")
  $currentIndex = $themes.IndexOf($Global:ThemeMode)
  if ($currentIndex -lt 0) { $currentIndex = 1 } # default to dark

  $rdoThemes = [Terminal.Gui.RadioGroup]::new($themes)
  $rdoThemes.X=2; $rdoThemes.Y=3; $rdoThemes.SelectedItem=$currentIndex
  $dlg.Add($rdoThemes)
    
  $btnApply = [Terminal.Gui.Button]::new("Apply")

  $btnApply.add_Clicked({
    $selectedTheme = $themes[$rdoThemes.SelectedItem]
    Debug-Log "Switching to theme: $selectedTheme"
    $Global:ThemeMode = $selectedTheme
        
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
      $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::White,[Terminal.Gui.Color]::Blue)
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

## -------------------------{ Theme Application }------------------------
function Apply-Theme {
  param(
    [hashtable]$ThemeData,
    [object]$TopLevel,
    [object]$MainWindow,
    [object]$Menu,
    [object]$Status
  )

  if ($null -eq $ThemeData) { return }

  ## Supported in 1.16
  [Terminal.Gui.Colors]::Base     = $ThemeData.Global
  [Terminal.Gui.Colors]::Dialog   = $ThemeData.Global
  [Terminal.Gui.Colors]::TopLevel = $ThemeData.Global
  [Terminal.Gui.Colors]::Menu     = $ThemeData.Global

  if ($Menu) {
    $Menu.ColorScheme = $ThemeData.Global
    $Menu.SetNeedsDisplay()
  }
  if ($Status) {
    $Status.ColorScheme = $ThemeData.Global
    $Status.SetNeedsDisplay()
  }
  if ($MainWindow) {
    $MainWindow.ColorScheme = $ThemeData.MainWindow
    $MainWindow.SetNeedsDisplay()
  }
  if ($TopLevel) {
    $TopLevel.SetNeedsDisplay()
  }

  [Terminal.Gui.Application]::Refresh()
}

## Diagnostics helper to show what's inside a ColorScheme
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
$Global:ThemeMode = $Theme

## Get the selected colour scheme
$cs = Get-Theme -mode $Theme

## Apply theme to all components
Apply-Theme -ThemeData $newTheme -TopLevel [Terminal.Gui.Application]::Top -MainWindow $win -Menu $menu -Status $StatusBar

## -------------------------{ Tools Modal }------------------------
function Show-ToolsDialog {

  ## --- Build dynamic 30% height dialog ---
  $termHeight = [Console]::WindowHeight
  $dlgHeight  = [Math]::Floor($termHeight * 0.20)
  if ($dlgHeight -lt 11) { $dlgHeight = 11 }   # minimum usable height

  $dialog = [Terminal.Gui.Dialog]::new("Available Tools", 50, $dlgHeight)

  ## --- Tool availability checks (stanza-style, same as your working code) ---
  $hasMore  = $null -ne (Get-Command more  -ErrorAction SilentlyContinue)
  $hasCurl  = $null -ne (Get-Command curl  -ErrorAction SilentlyContinue)
  $hasWhois = $null -ne (Get-Command whois -ErrorAction SilentlyContinue)
  $hasPing  = $null -ne (Get-Command ping  -ErrorAction SilentlyContinue)
  $hasArp   = $null -ne (Get-Command arp   -ErrorAction SilentlyContinue)

  ## --- Start printing at Y=1 (same as your style) ---
  $y = 0
  $dialog.Add([Terminal.Gui.Label]::new(2, $y, "Detected system tools:"))
  $y += 2

  ## --- Build list ---
  $toolStatus = @()

  if ($hasMore)  { $toolStatus += "more:     ✔  Available" } else { $toolStatus += "more:     ✖  Unavailable" }
  if ($hasCurl)  { $toolStatus += "curl:     ✔  Available" } else { $toolStatus += "curl:     ✖  Unavailable" }
  if ($hasWhois) { $toolStatus += "whois:    ✔  Available" } else { $toolStatus += "whois:    ✖  Unavailable" }
  if ($hasPing)  { $toolStatus += "ping:     ✔  Available" } else { $toolStatus += "ping:     ✖  Unavailable" }
  if ($hasArp)   { $toolStatus += "arp:      ✔  Available" } else { $toolStatus += "arp:      ✖  Unavailable" }

  ## --- Render each line vertically ---
  $toolStatus | ForEach-Object {
    $dialog.Add([Terminal.Gui.Label]::new(4, $y, $_))
    $y++
  }

  ## Add spacing before OK button
  $y += 2

  ## --- OK button ---
  $btnOK = [Terminal.Gui.Button]::new("OK")
  $btnOK.X = 28
  $btnOK.Y = $y
  $btnOK.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })
  $dialog.AddButton($btnOK)

  ## --- Show dialog ---
  [Terminal.Gui.Application]::Run($dialog)
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
    ## --- End Entropy update ---

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
  Show-Modal "Copied" "Password copied to clipboard."
  })

  ## Close
  $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() })

  [Terminal.Gui.Application]::Run($dlg)
  return $script:actualPassword
}

## This is a theme now. Danish Fruit soda based fun. Method to the
## madness
function Show-DanskFrugtInfo {
  $dlg = [Terminal.Gui.Dialog]::new("Why ($Global:DruitName)? 🫐", 60, 12)
    
  $message = @"
${ProjectName} is codenamed "Fruitname" because:

- Reason 1
- '($Global:FruitName)' (Danish Speling) is Danish for Fruit Namen
- Reason 2
- Every great project needs a forest-fruit mascot!
"@
    
  $label = [Terminal.Gui.Label]::new(1, 1, $message)
  $dlg.Add($label)
    
  [Terminal.Gui.MessageBox]::Query("Why ($Global:FruitName)? 🫐", $message, @("OK"))
}

## ------------------------{ Show message boxes }----------------------
## This cuts down significantly on code re-use
function Show-Modal { 
  param($title, $msg) 
  [Terminal.Gui.MessageBox]::Query($title, $msg, @("OK")) | Out-Null 
}

# ----------------------{ Password Generator View }----------------------
$PasswordMenuAction = {
  $dialog = [Terminal.Gui.Dialog]::new("Password Generator", 60, 12)

  $lbl = [Terminal.Gui.Label]::new("Password length:")
  $lbl.X = 1; $lbl.Y = 1
  $dialog.Add($lbl)

  $input = [Terminal.Gui.TextField]::new("16")
  $input.X = 18; $input.Y = 1; $input.Width = 10
  $dialog.Add($input)

  $output = [Terminal.Gui.TextView]::new()
  $output.X = 1; $output.Y = 4; $output.Width = 58; $output.Height = 4
  $dialog.Add($output)

  $btnGen = [Terminal.Gui.Button]::new("Generate")
  $btnGen.X = 1; $btnGen.Y = 9
  $btnGen.add_Click({
    $len = [int]$input.Text.ToString()
    if ($len -lt 1) { $len = 1 }

      # Full printable ASCII range 33–126
      $chars = 33..126 | ForEach-Object { [char]$_ }
      $rand = -join (1..$len | ForEach-Object { $chars[(Get-Random -Min 0 -Max $chars.Count)] })

      $output.Text = $rand
    })
  $dialog.Add($btnGen)

  $btnClose = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = 12; $btnClose.Y = 9
  $btnClose.add_Click({ $dialog.Running = $false })
  $dialog.Add($btnClose)

  [Terminal.Gui.Application]::Run($dialog)
}

## -----------------------------------{ Menus }-----------------------------------

## ---------------------------{ MenuBar (File / Help) }---------------------------
## Build MenuBarItems whose children are the MenuItem objects (same handlers as your first block)

$Menu = [Terminal.Gui.MenuBar]::new(@(
  # File menu
  [Terminal.Gui.MenuBarItem]::new("_File", @(
    [Terminal.Gui.MenuItem]::new("_Password Generator", "Random Password Generator (F2)", [Action]{ Generate-RandomPassword }),
    [Terminal.Gui.MenuItem]::new("_Show Tools", "Show Tools  (F3)", [Action]{ Show-ToolsDialog }),
    [Terminal.Gui.MenuItem]::new("_Undo", "Undo last action", [Action]{ Show-Modal "Not Implemented Yet" "Not yet`nCheck Future Builds" }),
    [Terminal.Gui.MenuItem]::new("_Themes", "Theme Selector (F12)", [Action]{ Show-ThemeSelector }),
    [Terminal.Gui.MenuItem]::new("_Exit", "Exit (F10)", [Action]{ [Terminal.Gui.Application]::RequestStop() })
  )),

  ## Help/About menu
  [Terminal.Gui.MenuBarItem]::new("_Help", @(
    [Terminal.Gui.MenuItem]::new("_Shortcuts", "Keyboard shortcuts (F1)", [Action]{ Show-Modal "Shortcuts" "F1 - Help`nF2 -Password Generator`nF3 Tools`n F10 - Quit`nF12 -Themes" }),
    [Terminal.Gui.MenuItem]::new("_Why Frught Namen?", "Danish Fruit Explainer", [Action]{ Show-DanskFrugtInfo }),
    [Terminal.Gui.MenuItem]::new("_About", "About ($Global:FruitName)", [Action]{
      Show-Modal "About" "$($Global:ProjectName)`n`nCodename: $($Global:FruitName)`nv$($Global:PSMC_Version) STABLE`nGPL-3 Copyleft`nBy Knightmare2600 (https://github.com/knightmare2600"
    })
  ))
))

## ---------------------{ Apply theme BEFORE adding to Top }---------------------
$themeData = Get-Theme -mode $Theme
$menu.ColorScheme = $themeData.Global

## ------------------------------{ StatusBar }-----------------------------
$StatusBar = [Terminal.Gui.StatusBar]::new(
  @(
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F1, "F1 - Help", { Show-Modal "Shortcuts" "F1 - Help`nF2 -Password Generator`nF3 Tools`n F10 - Quit`nF12 -Themes" }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F2, "F2 - Password Generator", { Show-PasswordGenerator }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F3, "F3 - Tools", { Show-ToolsDialog }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F10, "F10 - Quit", { [Terminal.Gui.Application]::RequestStop() }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F12, "F12 - Themes", { Show-ThemeSelector })
  )
)

## -------------------------{ Bottom Tools View }-------------------------
# Build bottom bar as a view inside remaining window space
$toolsView = [Terminal.Gui.View]::new()
$toolsView.Width  = [Terminal.Gui.Dim]::Fill()
$toolsView.Height = 1

## This is the ONLY correct, safe Terminal.Gui 1.16 bottom alignment:
## 1. We reduced $win.Height by 1 earlier
## 2. Now we place toolsView on the LAST line of the window using Frame.Height
$toolsView.Y = [Terminal.Gui.Pos]::At( $win.Frame.Height - 1 )
$toolsView.X = 0

## Add content to bottom tools bar
$toolsLabel = [Terminal.Gui.Label]::new(" F1 Help | F10 Quit ")
$toolsLabel.X = 1
$toolsLabel.Y = 0
$toolsView.Add($toolsLabel)

$win.Add($toolsView)

## ------------------------------{ Layout }------------------------------

## --- Init Application and Top ---
[Terminal.Gui.Application]::Init()
$top = [Terminal.Gui.Application]::Top

## Add UI components first
$top.Add($menu)
$top.Add($StatusBar)
$top.Add($win)

## Get the selected theme (use this variable)
$themeData = Get-Theme -mode $Theme
$Global:ThemeMode = $Theme
Debug-Log "Applying theme: $Theme"

## Apply theme globally and per-component
Apply-Theme -ThemeData $themeData -TopLevel $top -MainWindow $win -Menu $menu -Status $StatusBar

## Also ensure explicit per-component schemes (defensive)
if ($themeData -and $themeData.Global) {
  $win.ColorScheme      = $themeData.MainWindow
  $menu.ColorScheme     = $themeData.Global
  $StatusBar.ColorScheme= $themeData.Global
}

## Force redraws (Menu requires Draw(); Window/Menu/Top need SetNeedsDisplay)
$menu.SetNeedsDisplay()
$win.SetNeedsDisplay()
$StatusBar.SetNeedsDisplay()
$top.SetNeedsDisplay()

## Final refresh to push everything to screen
[Terminal.Gui.Application]::Refresh()

## Run app
[Terminal.Gui.Application]::Run()
[Terminal.Gui.Application]::Shutdown()
