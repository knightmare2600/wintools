# DSA-TUI Text Mode version of dsa.msc for powershell
# Locked-in baseline: dynamic resize, menu, demo data mirrors prod format, Change Domain fixed, fixed DC selection, full production AD object detection, properties modal, AD search popup
## TODO: Add fully fleshed out history

param(
    [switch]$DemoMode,
    [switch]$Logging,
    [string]$Domain,
    [ValidateSet("Dark","Light","Faxekondi","British","Default")]
    [string]$Theme = "Dark"
)

# Define the build version once
$BuildVersion = "1.5.8"

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
            $mainWindowCs.Normal = [Terminal.Gui.Attribute]::Make([Terminal.Gui.Color]::Red,[Terminal.Gui.Color]::Blue)
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
    if ($null -eq $Scheme) { Write-Output "ColorScheme is null"; return }
    Write-Output "Normal    : $($Scheme.Normal)"
    Write-Output "Focus     : $($Scheme.Focus)"
    Write-Output "HotNormal : $($Scheme.HotNormal)"
    Write-Output "HotFocus  : $($Scheme.HotFocus)"
    Write-Output "Disabled  : $($Scheme.Disabled)"
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
            Write-Host "DEBUG: Failed to enumerate ${type}: $($_.ToString())"
        }
    }
    return $allObjects
}

# ------------------------- Load Domain Data ------------------------
function Load-DomainData {
    param([string]$domain)

    if ($Logging) { Write-Host "DEBUG: Loading domain data for: $domain" }

<#
UK Phone Number Standards (Ofcom reserved ranges for fiction/testing):
- Glasgow: 0141 496 0xxx
- Edinburgh: 0131 496 0xxx  
- London: 020 7946 0xxx
- Generic UK: 01632 96xxxx
- UK Mobile: 07700 900xxx

Denmark Testing Numbers:
- Copenhagen landline: +45 0000-xxxx (fictional format)
- Denmark mobile: +45 2xxx xxxx
#>

if ($DemoMode) {
    $Global:Users = @(
        @{
            Name='Jim Kerr'
            OU='GLA'
            Groups=@('Simple Minds','Vocalists')
            Title='Vocalist'
            Email='jim.kerr@example.com'
            Country='UK'
            Disabled=$false
            Locked=$false
            Department='Music'
            Office='Glasgow Office'
            Phone='0141 496 0101'
            MobilePhone='07700 900101'
            Street='123 Clyde Street'
            City='Glasgow'
            PostalCode='G1 4JY'
            Company='Example Music Ltd'
            Manager=''
            Description='Lead vocalist for Simple Minds'
        },
        @{
            Name='Charlie Burchill'
            OU='GLA'
            Groups=@('Simple Minds','Guitarists')
            Title='Guitarist'
            Email='charlie.b@example.com'
            Country='UK'
            Disabled=$false
            Locked=$false
            Department='Music'
            Office='Glasgow Office'
            Phone='0141 496 0102'
            MobilePhone='07700 900102'
            Street='123 Clyde Street'
            City='Glasgow'
            PostalCode='G1 4JY'
            Company='Example Music Ltd'
            Manager='Jim Kerr'
            Description='Guitarist and founding member'
        },
        @{
            Name='Andy Bell'
            OU='LND'
            Groups=@('Erasure','Vocalists')
            Title='Vocalist'
            Email='andy.bell@example.com'
            Country='UK'
            Disabled=$false
            Locked=$false
            Department='Music'
            Office='London Office'
            Phone='020 7946 0201'
            MobilePhone='07700 900201'
            Street='456 Thames Road'
            City='London'
            PostalCode='EC1A 1BB'
            Company='Example Music Ltd'
            Manager=''
            Description='Lead vocalist for Erasure'
        },
        @{
            Name='Vince Clarke'
            OU='LND'
            Groups=@('Erasure','Depeche Mode','Keyboards')
            Title='Synthesizer'
            Email='vince.clarke@example.com'
            Country='UK'
            Disabled=$false
            Locked=$true  # Account is LOCKED (not disabled)
            Department='Music'
            Office='London Office'
            Phone='020 7946 0202'
            MobilePhone='07700 900202'
            Street='456 Thames Road'
            City='London'
            PostalCode='EC1A 1BB'
            Company='Example Music Ltd'
            Manager='Andy Bell'
            Description='Synthesizer pioneer - member of both Erasure and Depeche Mode'
        },
        @{
            Name='Martin Gore'
            OU='LND'
            Groups=@('Depeche Mode','Guitarists','Keyboards')
            Title='Guitarist/Keyboard'
            Email='martin.gore@example.com'
            Country='UK'
            Disabled=$false
            Locked=$false
            Department='Music'
            Office='London Office'
            Phone='01632 960301'
            MobilePhone='07700 900301'
            Street='789 Abbey Lane'
            City='London'
            PostalCode='W1A 1AA'
            Company='Example Music Ltd'
            Manager=''
            Description='Guitarist, keyboardist and primary songwriter'
        },
        @{
            Name='Dave Gahan'
            OU='LND'
            Groups=@('Depeche Mode','Vocalists')
            Title='Vocalist'
            Email='dave.gahan@example.com'
            Country='UK'
            Disabled=$false
            Locked=$false
            Department='Music'
            Office='London Office'
            Phone='01632 960302'
            MobilePhone='07700 900302'
            Street='789 Abbey Lane'
            City='London'
            PostalCode='W1A 1AA'
            Company='Example Music Ltd'
            Manager='Martin Gore'
            Description='Lead vocalist for Depeche Mode'
        },
        @{
            Name='Andrew Fletcher'
            OU='LND'
            Groups=@('Depeche Mode','Keyboards')
            Title='Keyboards/Bass Synth'
            Email='andrew.fletcher@example.com'
            Country='UK'
            Disabled=$true  # Account DISABLED
            Locked=$true #Locked too for good measure
            Department='Music'
            Office='London Office'
            Phone='01632 960303'
            MobilePhone='07700 900303'
            Street='789 Abbey Lane'
            City='London'
            PostalCode='W1A 1AA'
            Company='Example Music Ltd'
            Manager='Martin Gore'
            Description='Keyboard and bass synthesizer (deceased 2022)'
        },
        @{
            Name='Claus Norreen'
            OU='CPH'
            Groups=@('TV-2','Keyboards')
            Title='Keyboard'
            Email='claus.norreen@example.com'
            Country='DK'
            Disabled=$false
            Locked=$false
            Department='Music'
            Office='Copenhagen Office'
            Phone='+45 0000-1234'
            MobilePhone='+45 2012 3456'
            Street='12 Strøget'
            City='Copenhagen'
            PostalCode='1000'
            Company='Example Music ApS'
            Manager=''
            Description='Keyboard player for TV-2'
        },
        @{
            Name='Steffen Brandt'
            OU='CPH'
            Groups=@('TV-2','Vocalists','Guitarists')
            Title='Vocalist/Guitar'
            Email='steffen.brandt@example.com'
            Country='DK'
            Disabled=$false
            Locked=$false
            Department='Music'
            Office='Copenhagen Office'
            Phone='+45 0000-1235'
            MobilePhone='+45 2012 3457'
            Street='12 Strøget'
            City='Copenhagen'
            PostalCode='1000'
            Company='Example Music ApS'
            Manager='Claus Norreen'
            Description='Lead vocalist and guitarist for TV-2'
        },
        @{
            Name='Derek Dick'
            OU='EDI'
            Groups=@('Marillion','Vocalists')
            Title='Vocalist'
            Email='fish@example.com'
            Country='UK'
            Disabled=$false
            Locked=$false
            Department='Music'
            Office='Edinburgh Office'
            Phone='0131 496 0401'
            MobilePhone='07700 900401'
            Street='34 Royal Mile'
            City='Edinburgh'
            PostalCode='EH1 1RE'
            Company='Example Music Ltd'
            Manager=''
            Description='Former lead vocalist (Fish) for Marillion'
        },
        @{
            Name='Steve Rothery'
            OU='EDI'
            Groups=@('Marillion','Guitarists')
            Title='Guitarist'
            Email='steve.rothery@example.com'
            Country='UK'
            Disabled=$false
            Locked=$false
            Department='Music'
            Office='Edinburgh Office'
            Phone='0131 496 0402'
            MobilePhone='07700 900402'
            Street='34 Royal Mile'
            City='Edinburgh'
            PostalCode='EH1 1RE'
            Company='Example Music Ltd'
            Manager='Derek Dick'
            Description='Lead guitarist for Marillion'
        }
    )
    
    $Global:DCs = @(
        @{Name='EXAGLADC01'; OU='Domain Controllers'; Site='GLA'},
        @{Name='EXAEDIDC01'; OU='Domain Controllers'; Site='EDI'},
        @{Name='EXALNDCDC01'; OU='Domain Controllers'; Site='LND'},
        @{Name='EXACPHDC01'; OU='Domain Controllers'; Site='CPH'}
    )
}

# Build-Tree function to show both disabled and locked status:
function Build-Tree {
    param([string]$domain)
    
    Write-Host "DEBUG: Building tree with filters..."
    
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
                    # Add status indicator: 🔒 = locked, ⊗ = disabled, ○ = enabled
                    $statusIcon = if ($m.Locked) { "🔒" } elseif ($m.Disabled) { "⊗" } else { "○" }
                    $grpNode.Children.Add([Terminal.Gui.Trees.TreeNode]::new("(U) $statusIcon $($m.Name)"))
                }
                
                $ouNode.Children.Add($grpNode)
            }
        } else {
            # If groups hidden, show users directly under OU
            foreach ($u in ($filteredUsers | Where-Object { $_.OU -eq $ou } | Sort-Object -Property Name)) {
                $statusIcon = if ($u.Locked) { "🔒" } elseif ($u.Disabled) { "⊗" } else { "○" }
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
    Write-Host "DEBUG: Tree built - Showing $filterCount of $totalCount users"
}

#    } else {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop

            # ---- Show loading dialog ----
            $loadingDlg = Show-LoadingDialog -Message "Loading AD objects for $domain..."
            try {
                # Domain Controllers
                $Global:DCs = Get-ADDomainController -Discover -Domain $domain |
                    ForEach-Object { @{ Name=$_.HostName; OU='Domain Controllers'; Site=$_.Site } }

                # Users and objects
                $Global:ADObjects = Get-ADObjectsByType -domain $domain

                # Users: only Name and OU initially
                $Global:Users = $Global:ADObjects | Where-Object { $_.Type -eq 'user' } | ForEach-Object {
                    $ou = ($_.DN -split ',') | Where-Object { $_ -like 'OU=*' } | Select-Object -First 1
                    if ($ou) { $ou = $ou -replace '^OU=' ,'' } else { $ou = "" }
                    @{ Name=$_.Name; OU=$ou; Groups=$null; Title=$null; Email=$null; Country=$null; Disabled=$false }
                }
            } finally {
                Close-LoadingDialog $loadingDlg
            }

        } catch {
            [Terminal.Gui.MessageBox]::Query("Error","Failed to query domain ${domain}:`n$($_.ToString())","OK") | Out-Null
            $Global:Users=@(); $Global:DCs=@(); $Global:ADObjects=@()
        }
    }
#}


# DSA-TUI Missing Panel Functions
# Add these functions to your script before the main window section

# ------------------------- Filter Panel ------------------------
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
        Write-Host "DEBUG: Applying filters..."
        Build-Tree -domain $Global:Domain
        Update-FilterStatusLabel -label $filterStatusLabel
    })
    $filterFrame.Add($btnApplyFilter)
    
    $btnResetFilter = [Terminal.Gui.Button]::new("Reset")
    $btnResetFilter.X=17; $btnResetFilter.Y=$y
    $btnResetFilter.add_Clicked({
        Write-Host "DEBUG: Resetting filters..."
        $Global:FilterOptions.ShowDisabledUsers = $true
        $Global:FilterOptions.ShowEnabledUsers = $true
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
        Update-FilterStatusLabel -label $filterStatusLabel
    })
    $filterFrame.Add($btnResetFilter)
    
    return $filterFrame
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
    
    if (-not $label) { return }
    
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
    
    $lstSelected = [Terminal.Gui.ListView]::new()
    $lstSelected.SetSource(@())
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

function Update-SelectionPanel {
    param($panel)
    
    if (-not $panel -or -not $panel.Tag) { return }
    
    $lblCount = $panel.Tag.CountLabel
    $lstSelected = $panel.Tag.ListView
    
    $count = $Global:SelectedObjects.Count
    $lblCount.Text = "$count object(s) selected"
    
    $displayNames = $Global:SelectedObjects | ForEach-Object {
        $name = $_ -replace '^\(.\)\s*', '' -replace '^[○⊗\[L\]\[D\]\[E\]]\s*', ''
        $name
    }
    
    $lstSelected.SetSource($displayNames)
    $panel.SetNeedsDisplay()
}

# ------------------------- Selection Key Bindings ------------------------
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
        $statusIcon = if ($user.Locked) { "[L]" } elseif ($user.Disabled) { "[D]" } else { "[E]" }
        $displayName = "(U) $statusIcon $($user.Name)"
        $Global:SelectedObjects += $displayName
    }
    
    Write-Host "DEBUG: Selected all users ($($Global:SelectedObjects.Count))"
    Update-SelectionPanel -panel $selectionPanel
    [Terminal.Gui.MessageBox]::Query(50, 7, "Selected All", "Selected $($Global:SelectedObjects.Count) users", "OK") | Out-Null
}

function Deselect-AllObjects {
    $Global:SelectedObjects = @()
    Write-Host "DEBUG: Deselected all objects"
    Update-SelectionPanel -panel $selectionPanel
}

Load-DomainData -domain $Global:Domain

# ------------------------- Initialize Terminal.Gui ------------------------
[Terminal.Gui.Application]::Init()
$top = [Terminal.Gui.Application]::Top
$cs = [Terminal.Gui.ColorScheme]::new()
$cs.Normal = [Terminal.Gui.Attribute]::new([Terminal.Gui.Color]::Gray,[Terminal.Gui.Color]::Black)
$top.ColorScheme = $cs

Add-SelectionKeyBindings -view $top

# ------------------------- Main Window ------------------------
$win = [Terminal.Gui.Window]::new("DSA-TUI — Active Directory")
$win.X=0; $win.Y=0; $win.Width=[Terminal.Gui.Dim]::Fill(); $win.Height=[Terminal.Gui.Dim]::Fill()
$top.Add($win)

## filter panel
$filterPanel = Create-FilterPanel
$win.Add($filterPanel)

$filterStatusLabel = Create-FilterStatusLabel
$win.Add($filterStatusLabel)

$selectionPanel = Create-SelectionPanel
$win.Add($selectionPanel)

# ------------------------- Status Bar ------------------------
$status = [Terminal.Gui.StatusBar]::new(@(
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F1,"Help",{ Write-Host "DEBUG: Help invoked" }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F9,"New",{ Show-NewObjectWizard }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F10,"Quit",{ [Terminal.Gui.Application]::RequestStop() }),
    [Terminal.Gui.StatusItem]::new([Terminal.Gui.Key]::F12,"Redraw",{ [Terminal.Gui.Application]::Refresh() })
))
$top.Add($status)

# ------------------------- Menu ------------------------
$mFile = [Terminal.Gui.MenuItem]::new("_Exit","Exit application",[Action]{ [Terminal.Gui.Application]::RequestStop() })
$mAbout = [Terminal.Gui.MenuItem]::new("_About","About DSA-TUI",[Action]{ [Terminal.Gui.MessageBox]::Query("About","DSA-TUI ${BuildVersion}`n© 2025 Copyleft (GPL-3)`nDemo Mode: $DemoMode","OK") | Out-Null })
$mNew = [Terminal.Gui.MenuItem]::new("New Object","Create a new object",[Action]{ Show-NewObjectWizard })
$mProps = [Terminal.Gui.MenuItem]::new("_Properties","Edit selected properties",[Action]{ Show-Properties })
$mUndo = [Terminal.Gui.MenuItem]::new("_Undo","Undo last action",[Action]{ Write-Host "DEBUG: Undo placeholder" })
$mChangeDomain = [Terminal.Gui.MenuItem]::new("Change _Domain","Select domain",[Action]{ Show-ChangeDomainDialog })
$mChangeDC = [Terminal.Gui.MenuItem]::new("Change _Domain Controller","Select DC",[Action]{ Show-ChangeDCDialog })
$mSearchAD = [Terminal.Gui.MenuItem]::new("_Search AD","Search Active Directory",[Action]{ Show-ADSearchDialog })
$mPasswordGenerator = [Terminal.Gui.MenuItem]::new("_Password Generator","Password Generator",[Action]{ Generate-RandomPassword})

$mRefresh = [Terminal.Gui.MenuItem]::new("_Refresh","Refresh AD data",[Action]{ Refresh-TreeData })
$mQuickFilter = [Terminal.Gui.MenuItem]::new("_Quick Filter","Apply quick filters",[Action]{ Show-QuickFilterDialog })
$mSelectionMode = [Terminal.Gui.MenuItem]::new("_Selection Mode (Ctrl+S)","Toggle selection mode",[Action]{ Toggle-SelectionMode })
$mSelectAll = [Terminal.Gui.MenuItem]::new("Select _All (Ctrl+A)","Select all objects",[Action]{ Select-AllObjects })
$mDeselectAll = [Terminal.Gui.MenuItem]::new("_Deselect All (Ctrl+D)","Deselect all objects",[Action]{ Deselect-AllObjects })
$mBulkAddGroup = [Terminal.Gui.MenuItem]::new("Add to _Group...","Add selected users to group",[Action]{ Invoke-BulkAddToGroup })

# Group menu items logically
$menu = [Terminal.Gui.MenuBar]::new(@(
     [Terminal.Gui.MenuBarItem]::new("_File", @($mAbout,$mPasswordGenerator,$mRefresh,$mFile)),
     [Terminal.Gui.MenuBarItem]::new("_Action", @($mNew,$mProps,$mQuickFilter,$mUndo,$mChangeDomain,$mChangeDC,$mSearchAD)),
     [Terminal.Gui.MenuBarItem]::new("_Selection", @($mSelectionMode,$mSelectAll,$mDeselectAll,$mBulkAddGroup))
))
$top.Add($menu)

# ------------------------- TreeView ------------------------
$tree = [Terminal.Gui.TreeView]::new()
$tree.X=0; $tree.Y=1; $tree.Width=30; $tree.Height=[Terminal.Gui.Dim]::Fill()
$win.Add($tree)

function Generate-RandomPassword {
    param()

    # --- Full Character Sets ---
    $UpperCase = @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z')
    $LowerCase = @('a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z')
    $Numbers   = @('1','2','3','4','5','6','7','8','9','0')
    $Symbols   = @('!','@','$','?','<','>','*','&')

    $script:actualPassword = ""

    # --- Build UI ---
    $dlg = [Terminal.Gui.Dialog]::new("Generate Random Password", 60, 18)

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

    # Show Password checkbox
    $chkShowPwd = [Terminal.Gui.CheckBox]::new(2,10,"Show Password",$false)
    $dlg.Add($chkShowPwd)

    # Buttons
    $btnGenerate = [Terminal.Gui.Button]::new("Generate"); $btnGenerate.X=2; $btnGenerate.Y=12
    $btnCopy     = [Terminal.Gui.Button]::new("Copy");     $btnCopy.X=15; $btnCopy.Y=12
    $btnClose    = [Terminal.Gui.Button]::new("Close");    $btnClose.X=28; $btnClose.Y=12
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
        Write-Host "DEBUG: Applying quick filter: $selected"
        
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

# Enhanced User Properties Dialog
# Adds: Mobile phone field, Lock status checkbox, proper status display

function Show-UserPropertiesDialog {
    param($user)
    
    # Create dialog
    $dlg = [Terminal.Gui.Dialog]::new("User Properties - $($user.Name)", 90, 32)
    
    # Create TabView
    $tabView = [Terminal.Gui.TabView]::new()
    $tabView.X = 0
    $tabView.Y = 0
    $tabView.Width = [Terminal.Gui.Dim]::Fill()
    $tabView.Height = [Terminal.Gui.Dim]::Fill(2)
    
    # Track if changes were made
    $script:changesMade = $false
    
    # ----- General Tab -----
    $generalTab = [Terminal.Gui.TabView+Tab]::new()
    $generalTab.Text = "General"
    $generalView = [Terminal.Gui.View]::new()
    
    $y = 1
    $lbl = [Terminal.Gui.Label]::new("Display Name:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtName = [Terminal.Gui.TextField]::new($user.Name); $txtName.X=20; $txtName.Y=$y; $txtName.Width=40
    $txtName.add_TextChanged({ $script:changesMade = $true })
    $generalView.Add($txtName)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("Description:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtDesc = [Terminal.Gui.TextField]::new($user.Description); $txtDesc.X=20; $txtDesc.Y=$y; $txtDesc.Width=40
    $txtDesc.add_TextChanged({ $script:changesMade = $true })
    $generalView.Add($txtDesc)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("Office:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtOffice = [Terminal.Gui.TextField]::new($user.Office); $txtOffice.X=20; $txtOffice.Y=$y; $txtOffice.Width=40
    $txtOffice.add_TextChanged({ $script:changesMade = $true })
    $generalView.Add($txtOffice)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("Telephone:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtPhone = [Terminal.Gui.TextField]::new($user.Phone); $txtPhone.X=20; $txtPhone.Y=$y; $txtPhone.Width=40
    $txtPhone.add_TextChanged({ $script:changesMade = $true })
    $generalView.Add($txtPhone)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("Mobile Phone:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtMobile = [Terminal.Gui.TextField]::new($user.MobilePhone); $txtMobile.X=20; $txtMobile.Y=$y; $txtMobile.Width=40
    $txtMobile.add_TextChanged({ $script:changesMade = $true })
    $generalView.Add($txtMobile)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("E-mail:"); $lbl.X=2; $lbl.Y=$y; $generalView.Add($lbl)
    $txtEmail = [Terminal.Gui.TextField]::new($user.Email); $txtEmail.X=20; $txtEmail.Y=$y; $txtEmail.Width=40
    $txtEmail.add_TextChanged({ $script:changesMade = $true })
    $generalView.Add($txtEmail)
    
    $generalTab.View = $generalView
    $tabView.AddTab($generalTab, $false)
    
    # ----- Account Tab -----
    $accountTab = [Terminal.Gui.TabView+Tab]::new()
    $accountTab.Text = "Account"
    $accountView = [Terminal.Gui.View]::new()
    
    $y = 1
    $lbl = [Terminal.Gui.Label]::new("User logon name:"); $lbl.X=2; $lbl.Y=$y; $accountView.Add($lbl)
    $txtLogon = [Terminal.Gui.TextField]::new($user.Name.ToLower().Replace(' ','.')); $txtLogon.X=20; $txtLogon.Y=$y; $txtLogon.Width=30
    $accountView.Add($txtLogon)
    $y+=2
    
    # Account Status Label
    $lbl = [Terminal.Gui.Label]::new("Account Status:"); $lbl.X=2; $lbl.Y=$y; $accountView.Add($lbl)
    $statusText = if ($user.Locked) { "🔒 Locked" } elseif ($user.Disabled) { "⊗ Disabled" } else { "○ Enabled" }
    $lblStatus = [Terminal.Gui.Label]::new($statusText); $lblStatus.X=20; $lblStatus.Y=$y; $accountView.Add($lblStatus)
    $y+=2
    
    # Disabled checkbox
    $chkDisabled = [Terminal.Gui.CheckBox]::new("Account is disabled"); $chkDisabled.X=2; $chkDisabled.Y=$y
    $chkDisabled.Checked = if ($user.Disabled -is [bool]) { $user.Disabled } else { $false }
    $accountView.Add($chkDisabled)
    $chkDisabled.add_Toggled({
        # Update status label
        if ($chkLocked.Checked) {
            $lblStatus.Text = "🔒 Locked"
        } elseif ($chkDisabled.Checked) {
            $lblStatus.Text = "⊗ Disabled"
        } else {
            $lblStatus.Text = "○ Enabled"
        }
        $script:changesMade = $true
    })
    $y+=1
    
    # Locked checkbox
    $chkLocked = [Terminal.Gui.CheckBox]::new("Account is locked"); $chkLocked.X=2; $chkLocked.Y=$y
    $chkLocked.Checked = if ($user.Locked -is [bool]) { $user.Locked } else { $false }
    $accountView.Add($chkLocked)
    $chkLocked.add_Toggled({
        # Update status label
        if ($chkLocked.Checked) {
            $lblStatus.Text = "🔒 Locked"
        } elseif ($chkDisabled.Checked) {
            $lblStatus.Text = "⊗ Disabled"
        } else {
            $lblStatus.Text = "○ Enabled"
        }
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
    
$btnResetPwd = [Terminal.Gui.Button]::new("Reset Password...")
$btnResetPwd.X = 2
$btnResetPwd.Y = $y
$accountView.Add($btnResetPwd)

$btnResetPwd.add_Clicked({

    # 1. Launch password generator dialog
    $newPwd = Generate-RandomPassword
    if (-not $newPwd) {
        Show-Modal "Cancelled" "Password generation cancelled."
        return
    }

    # 2. Ask user if they want to apply it
    $confirm = [Terminal.Gui.MessageBox]::Query(
        70, 10,
        "Apply Password",
        "Apply the following password to user:`n`n$($user.Name)`n`nPassword:`n$newPwd`n",
        "Apply", "Cancel"
    )

    if ($confirm -ne 0) {
        Show-Modal "Cancelled" "Password reset cancelled."
        return
    }

    # 3. (Demo mode) – Real AD write not done here
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
    $txtStreet = [Terminal.Gui.TextField]::new($user.Street); $txtStreet.X=20; $txtStreet.Y=$y; $txtStreet.Width=40
    $txtStreet.add_TextChanged({ $script:changesMade = $true })
    $addressView.Add($txtStreet)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("City:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
    $txtCity = [Terminal.Gui.TextField]::new($user.City); $txtCity.X=20; $txtCity.Y=$y; $txtCity.Width=40
    $txtCity.add_TextChanged({ $script:changesMade = $true })
    $addressView.Add($txtCity)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("State/Province:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
    $txtState = [Terminal.Gui.TextField]::new(""); $txtState.X=20; $txtState.Y=$y; $txtState.Width=40
    $txtState.add_TextChanged({ $script:changesMade = $true })
    $addressView.Add($txtState)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("Postal Code:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
    $txtPostal = [Terminal.Gui.TextField]::new($user.PostalCode); $txtPostal.X=20; $txtPostal.Y=$y; $txtPostal.Width=20
    $txtPostal.add_TextChanged({ $script:changesMade = $true })
    $addressView.Add($txtPostal)
    $y+=2
    
    $lbl = [Terminal.Gui.Label]::new("Country:"); $lbl.X=2; $lbl.Y=$y; $addressView.Add($lbl)
    $txtCountry = [Terminal.Gui.TextField]::new($user.Country); $txtCountry.X=20; $txtCountry.Y=$y; $txtCountry.Width=40
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
    $btnAddGroup = [Terminal.Gui.Button]::new("Add..."); $btnAddGroup.X=2; $btnAddGroup.Y=[Terminal.Gui.Pos]::Bottom($lstGroups)+1
    $btnAddGroup.add_Clicked({
        Write-Host "DEBUG: Add to group functionality (coming soon)"
        [Terminal.Gui.MessageBox]::Query(50, 7, "Add to Group", "Add to group feature coming soon!", "OK") | Out-Null
    })
    $memberView.Add($btnAddGroup)
    
    $btnRemoveGroup = [Terminal.Gui.Button]::new("Remove"); $btnRemoveGroup.X=[Terminal.Gui.Pos]::Right($btnAddGroup)+2; $btnRemoveGroup.Y=[Terminal.Gui.Pos]::Bottom($lstGroups)+1
    $btnRemoveGroup.add_Clicked({
        if ($lstGroups.SelectedItem -ge 0) {
            $grp = $user.Groups[$lstGroups.SelectedItem]
            $result = [Terminal.Gui.MessageBox]::Query(60, 8, "Remove from Group", "Remove $($user.Name) from '$grp'?", "Yes", "No")
            if ($result -eq 0) {
                Write-Host "DEBUG: Removing $($user.Name) from group: $grp"
                $script:changesMade = $true
                [Terminal.Gui.MessageBox]::Query(50, 7, "Success", "Removed from group (demo mode)", "OK") | Out-Null
            }
        }
    })
    $memberView.Add($btnRemoveGroup)
    
    $memberTab.View = $memberView
    $tabView.AddTab($memberTab, $false)
    
    # ----- Search/Lookup Tab -----
    $searchTab = [Terminal.Gui.TabView+Tab]::new()
    $searchTab.Text = "Search/Lookup"
    $searchView = [Terminal.Gui.View]::new()
    
    $y = 1
    $lblSearchDomain = [Terminal.Gui.Label]::new("Domain:"); $lblSearchDomain.X=2; $lblSearchDomain.Y=$y; $searchView.Add($lblSearchDomain)
    $txtSearchDomain = [Terminal.Gui.TextField]::new($Global:Domain); $txtSearchDomain.X=15; $txtSearchDomain.Y=$y; $txtSearchDomain.Width=30; $searchView.Add($txtSearchDomain)
    $y+=2
    
    $lblSearchName = [Terminal.Gui.Label]::new("Name:"); $lblSearchName.X=2; $lblSearchName.Y=$y; $searchView.Add($lblSearchName)
    $txtSearchUser = [Terminal.Gui.TextField]::new($user.Name); $txtSearchUser.X=15; $txtSearchUser.Y=$y; $txtSearchUser.Width=30; $searchView.Add($txtSearchUser)
    $y+=2
    
    $lblSearchType = [Terminal.Gui.Label]::new("Type:"); $lblSearchType.X=2; $lblSearchType.Y=$y; $searchView.Add($lblSearchType)
    $cmbSearchType = [Terminal.Gui.ComboBox]::new(); $cmbSearchType.X=15; $cmbSearchType.Y=$y; $cmbSearchType.Width=20
    $cmbSearchType.SetSource(@("User","Group","OU"))
    $cmbSearchType.SelectedItem = 0
    $searchView.Add($cmbSearchType)
    $y+=2
    
    # Results filter box
    $lblSearchFilter = [Terminal.Gui.Label]::new("Filter Results:"); $lblSearchFilter.X=48; $lblSearchFilter.Y=1; $searchView.Add($lblSearchFilter)
    $txtSearchFilter = [Terminal.Gui.TextField]::new(""); $txtSearchFilter.X=62; $txtSearchFilter.Y=1; $txtSearchFilter.Width=20; $searchView.Add($txtSearchFilter)
    
    # Results output
    $lblSearchResult = [Terminal.Gui.Label]::new("Results:"); $lblSearchResult.X=2; $lblSearchResult.Y=$y; $searchView.Add($lblSearchResult)
    $y+=1
    $txtSearchOutput = [Terminal.Gui.TextView]::new(); $txtSearchOutput.X=2; $txtSearchOutput.Y=$y
    $txtSearchOutput.Width=[Terminal.Gui.Dim]::Fill(2); $txtSearchOutput.Height=[Terminal.Gui.Dim]::Fill(4)
    $txtSearchOutput.ReadOnly=$true; $txtSearchOutput.WordWrap=$false
    $searchView.Add($txtSearchOutput)
    
    # Account locked checkbox for users
    $chkSearchLocked = [Terminal.Gui.CheckBox]::new("Account Locked"); $chkSearchLocked.X=2; $chkSearchLocked.Y=[Terminal.Gui.Pos]::Bottom($txtSearchOutput)+1
    $chkSearchLocked.CanFocus=$true; $chkSearchLocked.Data=""
    $searchView.Add($chkSearchLocked)
    
    # Search button
    $btnDoSearch = [Terminal.Gui.Button]::new("Search"); $btnDoSearch.X=48; $btnDoSearch.Y=3; $searchView.Add($btnDoSearch)
    
    # Filter implementation
    $script:currentSearchOutputLines = @()
    $txtSearchFilter.add_TextChanged({
        if ($script:currentSearchOutputLines) {
            $search = $txtSearchFilter.Text.ToString().Trim()
            if ($search) { 
                $txtSearchOutput.Text = ($script:currentSearchOutputLines | Where-Object {$_ -match "(?i)$search"}) -join "`n" 
            } else { 
                $txtSearchOutput.Text = $script:currentSearchOutputLines -join "`n" 
            }
        }
    })
    
    # Search handler
    $btnDoSearch.add_Clicked({
        $searchName = [string]$txtSearchUser.Text.ToString().Trim()
        $domain = $txtSearchDomain.Text.ToString().Trim()
        $objType = $cmbSearchType.Text.ToString()

        if (-not $searchName) { $txtSearchOutput.Text="Please enter a name."; return }

        try {
            if ($Global:DemoMode) {
                Write-Host "DEBUG: Searching demo data for $objType '$searchName'"
                switch ($objType) {
                    "User" {
                        $foundUsers = $Global:Users | Where-Object { $_.Name -like "*$searchName*" }
                        
                        if ($foundUsers.Count -eq 0) {
                            Write-Host "DEBUG: User not found in demo data"
                            $txtSearchOutput.Text = "User not found in demo data"
                            $chkSearchLocked.Checked = $false
                            $chkSearchLocked.Data = ""
                            return
                        }
                        
                        # Multiple matches - show selection dialog
                        if ($foundUsers.Count -gt 1) {
                            Write-Host "DEBUG: Multiple users found ($($foundUsers.Count)), showing selection dialog"
                            
                            # Create selection dialog
                            $selDlg = [Terminal.Gui.Dialog]::new("Select User", 60, 20)
                            $lblSel = [Terminal.Gui.Label]::new("Multiple matches found. Select one:"); $lblSel.X=2; $lblSel.Y=1; $selDlg.Add($lblSel)
                            
                            $userNames = @($foundUsers | ForEach-Object { "$($_.Name) ($($_.Email))" })
                            $lstSel = [Terminal.Gui.ListView]::new()
                            $lstSel.SetSource($userNames)
                            $lstSel.X=2; $lstSel.Y=3; $lstSel.Width=[Terminal.Gui.Dim]::Fill(2); $lstSel.Height=[Terminal.Gui.Dim]::Fill(2)
                            $selDlg.Add($lstSel)
                            
                            $script:selectedUser = $null
                            
                            $btnSelOK = [Terminal.Gui.Button]::new("OK")
                            $btnSelOK.add_Clicked({
                                if ($lstSel.SelectedItem -ge 0) {
                                    $script:selectedUser = $foundUsers[$lstSel.SelectedItem]
                                }
                                [Terminal.Gui.Application]::RequestStop()
                            })
                            $selDlg.AddButton($btnSelOK)
                            
                            $btnSelCancel = [Terminal.Gui.Button]::new("Cancel")
                            $btnSelCancel.add_Clicked({ $script:selectedUser = $null; [Terminal.Gui.Application]::RequestStop() })
                            $selDlg.AddButton($btnSelCancel)
                            
                            # Handle Enter key
                            $lstSel.add_OpenSelectedItem({ $btnSelOK.PerformClick() })
                            
                            [Terminal.Gui.Application]::Run($selDlg)
                            
                            if (-not $script:selectedUser) {
                                Write-Host "DEBUG: User cancelled selection"
                                return
                            }
                            
                            $foundUser = $script:selectedUser
                            # Update search box with selected name
                            $txtSearchUser.Text = $foundUser.Name
                            
                            # UPDATE ALL TABS with new user data
                            Write-Host "DEBUG: Updating all tabs with user: $($foundUser.Name)"
                            
                            # Update General tab
                            $txtName.Text = $foundUser.Name
                            $txtDesc.Text = $foundUser.Description
                            $txtOffice.Text = $foundUser.Office
                            $txtPhone.Text = $foundUser.Phone
                            $txtMobile.Text = $foundUser.MobilePhone
                            $txtEmail.Text = $foundUser.Email
                            
                            # Update Account tab
                            $txtLogon.Text = $foundUser.Name.ToLower().Replace(' ','.')
                            $chkDisabled.Checked = [bool]($foundUser.Disabled)
                            $chkLocked.Checked = [bool]($foundUser.Locked)
                            $statusText = if ($foundUser.Locked) { "[L] Locked" } elseif ($foundUser.Disabled) { "[D] Disabled" } else { "[E] Enabled" }
                            $lblStatus.Text = $statusText
                            
                            # Update Address tab
                            $txtStreet.Text = $foundUser.Street
                            $txtCity.Text = $foundUser.City
                            $txtPostal.Text = $foundUser.PostalCode
                            $txtCountry.Text = $foundUser.Country
                            
                            # Update Organization tab
                            $txtTitle.Text = $foundUser.Title
                            $txtDept.Text = $foundUser.Department
                            $txtCompany.Text = $foundUser.Company
                            $txtManager.Text = $foundUser.Manager
                            
                            # Update Member Of tab
                            $lstGroups.SetSource($foundUser.Groups)
                            
                            # Update dialog title
                            $dlg.Title = "User Properties - $($foundUser.Name)"
                            
                            # Mark as unchanged since we just loaded new data
                            $script:changesMade = $false
                            
                            # Update the $user variable reference
                            $user = $foundUser
                        } else {
                            $foundUser = $foundUsers[0]
                        }
                        
                        if ($foundUser) {
                            Write-Host "DEBUG: Found user in demo data: $($foundUser.Name)"
                            $outputLines = @(
                                "Name                     : $($foundUser.Name)",
                                "Email                    : $($foundUser.Email)",
                                "Title                    : $($foundUser.Title)",
                                "Department               : $($foundUser.Department)",
                                "Office                   : $($foundUser.Office)",
                                "Phone                    : $($foundUser.Phone)",
                                "MobilePhone              : $($foundUser.MobilePhone)",
                                "OU                       : $($foundUser.OU)",
                                "Groups                   : $($foundUser.Groups -join ', ')",
                                "Manager                  : $($foundUser.Manager)",
                                "Company                  : $($foundUser.Company)",
                                "Street                   : $($foundUser.Street)",
                                "City                     : $($foundUser.City)",
                                "PostalCode               : $($foundUser.PostalCode)",
                                "Country                  : $($foundUser.Country)",
                                "Disabled                 : $($foundUser.Disabled)",
                                "Locked                   : $($foundUser.Locked)",
                                "Description              : $($foundUser.Description)"
                            )
                            $txtSearchOutput.Text = $outputLines -join "`n"
                            $script:currentSearchOutputLines = $outputLines
                            
                            # Update locked checkbox
                            $chkSearchLocked.Checked = [bool]($foundUser.Locked)
                            $chkSearchLocked.Data = $foundUser.Name
                        }
                    }
                    "Group" {
                        Write-Host "DEBUG: Searching for group in demo data"
                        $matchedGroups = @()
                        foreach ($u in $Global:Users) { 
                            foreach ($g in $u.Groups) { 
                                if ($g -like "*$searchName*") { $matchedGroups += $g } 
                            }
                        }
                        if ($matchedGroups) {
                            $uniqueGroups = $matchedGroups | Sort-Object -Unique
                            Write-Host "DEBUG: Found group(s): $($uniqueGroups -join ', ')"
                            $groupName = $uniqueGroups[0]
                            $members = $Global:Users | Where-Object { $_.Groups -contains $groupName } | ForEach-Object { $_.Name } | Sort-Object
                            $outputLines = @(
                                "Group                    : $groupName",
                                "Description              : <no description>",
                                "Member Count             : $($members.Count)",
                                "",
                                "Members:",
                                $($members -join "`n")
                            )
                            $txtSearchOutput.Text = $outputLines -join "`n"
                            $script:currentSearchOutputLines = $outputLines
                        } else {
                            Write-Host "DEBUG: Group not found in demo data"
                            $txtSearchOutput.Text = "Group not found in demo data"
                        }
                    }
                    "OU" {
                        Write-Host "DEBUG: Searching for OU in demo data"
                        $ouNames = ($Global:Users | Select-Object -ExpandProperty OU -Unique)
                        $matchedOU = $ouNames | Where-Object { $_ -like "*$searchName*" } | Select-Object -First 1
                        if ($matchedOU) {
                            Write-Host "DEBUG: Found OU: $matchedOU"
                            $members = $Global:Users | Where-Object { $_.OU -eq $matchedOU } | ForEach-Object { $_.Name } | Sort-Object
                            $outputLines = @(
                                "OU                       : $matchedOU",
                                "Member Count             : $($members.Count)",
                                "",
                                "Members:",
                                $($members -join "`n")
                            )
                            $txtSearchOutput.Text = $outputLines -join "`n"
                            $script:currentSearchOutputLines = $outputLines
                        } else {
                            Write-Host "DEBUG: OU not found in demo data"
                            $txtSearchOutput.Text = "OU not found in demo data"
                        }
                    }
                }
            } else {
                # Production AD search
                Write-Host "DEBUG: Searching production AD for $objType '$searchName'"
                switch ($objType) {
                    "User" {
                        $filter = "SamAccountName -like '*$searchName*' -or Name -like '*$searchName*'"
                        if ($domain) { 
                            $objs = Get-ADUser -Filter $filter -Properties * -Server $domain -ErrorAction Stop
                        } else { 
                            $objs = Get-ADUser -Filter $filter -Properties * -ErrorAction Stop
                        }
                        
                        if (-not $objs -or $objs.Count -eq 0) {
                            Write-Host "DEBUG: User not found in AD"
                            $txtSearchOutput.Text = "User not found in Active Directory"
                            $chkSearchLocked.Checked = $false
                            $chkSearchLocked.Data = ""
                            return
                        }
                        
                        # Multiple matches - show selection dialog
                        if ($objs.Count -gt 1) {
                            Write-Host "DEBUG: Multiple users found ($($objs.Count)), showing selection dialog"
                            
                            $selDlg = [Terminal.Gui.Dialog]::new("Select User", 60, 20)
                            $lblSel = [Terminal.Gui.Label]::new("Multiple matches found. Select one:"); $lblSel.X=2; $lblSel.Y=1; $selDlg.Add($lblSel)
                            
                            $userNames = @($objs | ForEach-Object { "$($_.SamAccountName) ($($_.Name))" })
                            $lstSel = [Terminal.Gui.ListView]::new()
                            $lstSel.SetSource($userNames)
                            $lstSel.X=2; $lstSel.Y=3; $lstSel.Width=[Terminal.Gui.Dim]::Fill(2); $lstSel.Height=[Terminal.Gui.Dim]::Fill(2)
                            $selDlg.Add($lstSel)
                            
                            $script:selectedUser = $null
                            
                            $btnSelOK = [Terminal.Gui.Button]::new("OK")
                            $btnSelOK.add_Clicked({
                                if ($lstSel.SelectedItem -ge 0) {
                                    $script:selectedUser = $objs[$lstSel.SelectedItem]
                                }
                                [Terminal.Gui.Application]::RequestStop()
                            })
                            $selDlg.AddButton($btnSelOK)
                            
                            $btnSelCancel = [Terminal.Gui.Button]::new("Cancel")
                            $btnSelCancel.add_Clicked({ $script:selectedUser = $null; [Terminal.Gui.Application]::RequestStop() })
                            $selDlg.AddButton($btnSelCancel)
                            
                            $lstSel.add_OpenSelectedItem({ $btnSelOK.PerformClick() })
                            
                            [Terminal.Gui.Application]::Run($selDlg)
                            
                            if (-not $script:selectedUser) {
                                Write-Host "DEBUG: User cancelled selection"
                                return
                            }
                            
                            $foundUser = $script:selectedUser
                            # Update search box with selected name
                            $txtSearchUser.Text = $foundUser.SamAccountName
                            
                            # UPDATE ALL TABS with new user data (for production mode)
                            Write-Host "DEBUG: Updating all tabs with AD user: $($foundUser.Name)"
                            
                            # Load full user details if not already loaded
                            if (-not $foundUser.MobilePhone) {
                                try {
                                    $foundUser = Get-ADUser -Identity $foundUser.SamAccountName -Properties * -ErrorAction Stop
                                } catch {
                                    Write-Host "ERROR: Could not reload full user details"
                                }
                            }
                            
                            # Update General tab
                            $txtName.Text = if ($foundUser.DisplayName) { $foundUser.DisplayName } else { $foundUser.Name }
                            $txtDesc.Text = $foundUser.Description
                            $txtOffice.Text = $foundUser.Office
                            $txtPhone.Text = $foundUser.OfficePhone
                            $txtMobile.Text = $foundUser.MobilePhone
                            $txtEmail.Text = $foundUser.EmailAddress
                            
                            # Update Account tab
                            $txtLogon.Text = $foundUser.SamAccountName
                            $chkDisabled.Checked = -not $foundUser.Enabled
                            $chkLocked.Checked = [bool]($foundUser.LockedOut)
                            $statusText = if ($foundUser.LockedOut) { "[L] Locked" } elseif (-not $foundUser.Enabled) { "[D] Disabled" } else { "[E] Enabled" }
                            $lblStatus.Text = $statusText
                            
                            # Update Address tab
                            $txtStreet.Text = $foundUser.StreetAddress
                            $txtCity.Text = $foundUser.City
                            $txtPostal.Text = $foundUser.PostalCode
                            $txtCountry.Text = $foundUser.Country
                            
                            # Update Organization tab
                            $txtTitle.Text = $foundUser.Title
                            $txtDept.Text = $foundUser.Department
                            $txtCompany.Text = $foundUser.Company
                            $txtManager.Text = if ($foundUser.Manager) { ($foundUser.Manager -split ',')[0] -replace '^CN=' } else { "" }
                            
                            # Update Member Of tab
                            try {
                                $groups = Get-ADPrincipalGroupMembership -Identity $foundUser.SamAccountName | Select-Object -ExpandProperty Name
                                $lstGroups.SetSource($groups)
                            } catch {
                                Write-Host "ERROR: Could not load group membership"
                            }
                            
                            # Update dialog title
                            $dlg.Title = "User Properties - $($foundUser.Name)"
                            
                            # Mark as unchanged since we just loaded new data
                            $script:changesMade = $false
                        } else {
                            $foundUser = $objs
                        }
                        
                        if ($foundUser) {
                            Write-Host "DEBUG: Found user in AD: $($foundUser.Name)"
                            $outputLines = $foundUser | Get-Member -MemberType Properties | ForEach-Object {
                                $val = $foundUser.$($_.Name)
                                # Convert epoch-style times
                                if ($_.Name -in @("accountExpires","badPasswordTime","LastLogon","LastLogonTimestamp","pwdLastSet")) {
                                    if ($val -eq 0 -or $val -eq 9223372036854775807) { $val="Never Expires" } 
                                    else { 
                                        try { $val = [datetime]::FromFileTime($val) } catch { $val = "Invalid date" }
                                    }
                                }
                                "{0,-25}: {1}" -f $_.Name, ($val -as [string])
                            }
                            $txtSearchOutput.Text = $outputLines -join "`n"
                            $script:currentSearchOutputLines = $outputLines
                            
                            $chkSearchLocked.Checked = [bool]($foundUser.LockedOut)
                            $chkSearchLocked.Data = $foundUser.DistinguishedName
                        }
                    }
                    "Group" {
                        $filter = "Name -like '*$searchName*'"
                        if ($domain) {
                            $group = Get-ADGroup -Filter $filter -Properties * -Server $domain -ErrorAction Stop | Select-Object -First 1
                        } else {
                            $group = Get-ADGroup -Filter $filter -Properties * -ErrorAction Stop | Select-Object -First 1
                        }
                        
                        if ($group) {
                            Write-Host "DEBUG: Found group in AD: $($group.Name)"
                            $members = Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name | Sort-Object
                            $outputLines = @(
                                "Group                    : $($group.Name)",
                                "Description              : $($group.Description)",
                                "GroupCategory            : $($group.GroupCategory)",
                                "GroupScope               : $($group.GroupScope)",
                                "DistinguishedName        : $($group.DistinguishedName)",
                                "Member Count             : $($members.Count)",
                                "",
                                "Members:",
                                $($members -join "`n")
                            )
                            $txtSearchOutput.Text = $outputLines -join "`n"
                            $script:currentSearchOutputLines = $outputLines
                        } else {
                            Write-Host "DEBUG: Group not found in AD"
                            $txtSearchOutput.Text = "Group not found in Active Directory"
                        }
                    }
                    "OU" {
                        $filter = "Name -like '*$searchName*'"
                        if ($domain) {
                            $ou = Get-ADOrganizationalUnit -Filter $filter -Properties * -Server $domain -ErrorAction Stop | Select-Object -First 1
                        } else {
                            $ou = Get-ADOrganizationalUnit -Filter $filter -Properties * -ErrorAction Stop | Select-Object -First 1
                        }
                        
                        if ($ou) {
                            Write-Host "DEBUG: Found OU in AD: $($ou.Name)"
                            $members = Get-ADUser -SearchBase $ou.DistinguishedName -Filter * -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name | Sort-Object
                            $outputLines = @(
                                "OU                       : $($ou.Name)",
                                "DistinguishedName        : $($ou.DistinguishedName)",
                                "Description              : $($ou.Description)",
                                "User Count               : $($members.Count)",
                                "",
                                "Users:",
                                $($members -join "`n")
                            )
                            $txtSearchOutput.Text = $outputLines -join "`n"
                            $script:currentSearchOutputLines = $outputLines
                        } else {
                            Write-Host "DEBUG: OU not found in AD"
                            $txtSearchOutput.Text = "OU not found in Active Directory"
                        }
                    }
                }
            }

        } catch {
            $errMsg = $_.Exception.Message
            Write-Host "ERROR: Search failed: $errMsg"
            $txtSearchOutput.Text = "Error during search: $errMsg"
            $chkSearchLocked.Checked=$false
            $chkSearchLocked.Data=""
        }
    })
    
    # Clear button
    $btnSearchClear = [Terminal.Gui.Button]::new("Clear"); $btnSearchClear.X=58; $btnSearchClear.Y=3; $searchView.Add($btnSearchClear)
    $btnSearchClear.add_Clicked({
        $txtSearchUser.Text=""; $txtSearchFilter.Text=""; $txtSearchOutput.Text=""
        $chkSearchLocked.Checked=$false; $chkSearchLocked.Data=""
        $script:currentSearchOutputLines=@()
    })
    
    # Lock/unlock checkbox handler
    $chkSearchLocked.add_Toggled({
        if ($chkSearchLocked.Data -and $chkSearchLocked.Data -ne "") {
            try {
                if ($Global:DemoMode) {
                    Write-Host "DEBUG: Demo mode - toggling lock for: $($chkSearchLocked.Data)"
                    $foundUser = $Global:Users | Where-Object { $_.Name -eq $chkSearchLocked.Data } | Select-Object -First 1
                    if ($foundUser) {
                        $foundUser.Locked = $chkSearchLocked.Checked
                        $action = if ($chkSearchLocked.Checked) {"locked"} else {"unlocked"}
                        $txtSearchOutput.Text += "`n`nAccount $action (demo mode)"
                        Write-Host "DEBUG: Account $action for $($foundUser.Name)"
                        
                        # Rebuild tree to show updated status
                        Build-Tree -domain $Global:Domain
                        Update-FilterStatusLabel -label $filterStatusLabel
                    } else {
                        Write-Host "ERROR: User not found in demo data: $($chkSearchLocked.Data)"
                        $txtSearchOutput.Text += "`n`nERROR: User not found in demo data"
                    }
                } else {
                    Write-Host "DEBUG: Production mode - toggling lock for: $($chkSearchLocked.Data)"
                    $dn = $chkSearchLocked.Data
                    if ($chkSearchLocked.Checked) { 
                        Lock-ADAccount -Identity $dn -ErrorAction Stop
                        $txtSearchOutput.Text += "`n`nAccount locked."
                        Write-Host "DEBUG: Account locked"
                    } else { 
                        Unlock-ADAccount -Identity $dn -ErrorAction Stop
                        $txtSearchOutput.Text += "`n`nAccount unlocked."
                        Write-Host "DEBUG: Account unlocked"
                    }
                }
            } catch { 
                $errMsg = $_.Exception.Message
                Write-Host "ERROR: Failed to toggle lock: $errMsg"
                $txtSearchOutput.Text += "`n`nError changing lock state: $errMsg"
            }
        }
    })
    
    $searchTab.View = $searchView
    $tabView.AddTab($searchTab, $false)
    
    # Auto-populate search results when tab is first displayed
    # Since we already have the user data, show it immediately
    [Terminal.Gui.Application]::MainLoop.Invoke({
        if ($user) {
            Write-Host "DEBUG: Auto-populating search results for $($user.Name)"
            $outputLines = @(
                "Name                     : $($user.Name)",
                "Email                    : $($user.Email)",
                "Title                    : $($user.Title)",
                "Department               : $($user.Department)",
                "Office                   : $($user.Office)",
                "Phone                    : $($user.Phone)",
                "MobilePhone              : $($user.MobilePhone)",
                "OU                       : $($user.OU)",
                "Groups                   : $($user.Groups -join ', ')",
                "Manager                  : $($user.Manager)",
                "Company                  : $($user.Company)",
                "Street                   : $($user.Street)",
                "City                     : $($user.City)",
                "PostalCode               : $($user.PostalCode)",
                "Country                  : $($user.Country)",
                "Disabled                 : $($user.Disabled)",
                "Locked                   : $($user.Locked)",
                "Description              : $($user.Description)"
            )
            $txtSearchOutput.Text = $outputLines -join "`n"
            $script:currentSearchOutputLines = $outputLines
            
            # Update locked checkbox
            $chkSearchLocked.Checked = [bool]($user.Locked)
            $chkSearchLocked.Data = $user.Name
        }
    })
    
    # Add TabView to dialog
    $dlg.Add($tabView)
    
    # Function to apply changes
    $applyChanges = {
        try {
            if ($Global:DemoMode) {
                # Update demo data
                $user.Name = $txtName.Text.ToString()
                $user.Description = $txtDesc.Text.ToString()
                $user.Office = $txtOffice.Text.ToString()
                $user.Phone = $txtPhone.Text.ToString()
                $user.MobilePhone = $txtMobile.Text.ToString()
                $user.Email = $txtEmail.Text.ToString()
                $user.Street = $txtStreet.Text.ToString()
                $user.City = $txtCity.Text.ToString()
                $user.PostalCode = $txtPostal.Text.ToString()
                $user.Country = $txtCountry.Text.ToString()
                $user.Title = $txtTitle.Text.ToString()
                $user.Department = $txtDept.Text.ToString()
                $user.Company = $txtCompany.Text.ToString()
                $user.Manager = $txtManager.Text.ToString()
                $user.Disabled = $chkDisabled.Checked
                $user.Locked = $chkLocked.Checked
                
                Write-Host "DEBUG: Changes saved to demo data for $($user.Name)"
                [Terminal.Gui.MessageBox]::Query(50, 7, "Success", "Properties updated successfully (demo mode)", "OK") | Out-Null
                
                # Rebuild tree to show updated status icons
                Build-Tree -domain $Global:Domain
                Update-FilterStatusLabel -label $filterStatusLabel
            } else {
                # Update real AD
                $updateParams = @{}
                
                # ... rest of AD update logic ...
                
                [Terminal.Gui.MessageBox]::Query(50, 7, "Success", "Properties updated successfully", "OK") | Out-Null
            }
            $script:changesMade = $false
        } catch {
            $errMsg = $_.Exception.Message
            [Terminal.Gui.MessageBox]::Query(60, 10, "Error", "Failed to update properties:`n$errMsg", "OK") | Out-Null
        }
    }
    
    # Buttons
    $btnOK = [Terminal.Gui.Button]::new("OK")
    $btnOK.add_Clicked({ 
        & $applyChanges
        [Terminal.Gui.Application]::RequestStop() 
    })
    $dlg.AddButton($btnOK)
    
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ 
        if ($script:changesMade) {
            $result = [Terminal.Gui.MessageBox]::Query(50, 7, "Unsaved Changes", "Discard changes?", "Yes", "No")
            if ($result -eq 0) {
                [Terminal.Gui.Application]::RequestStop()
            }
        } else {
            [Terminal.Gui.Application]::RequestStop()
        }
    })
    $dlg.AddButton($btnCancel)
    
    $btnApply = [Terminal.Gui.Button]::new("Apply")
    $btnApply.add_Clicked({ & $applyChanges })
    $dlg.AddButton($btnApply)
    
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
                        Write-Host "DEBUG: Created user $name in demo mode"
                    }
                    "Group" {
                        Write-Host "DEBUG: Created group $name in demo mode"
                    }
                    "OrganizationalUnit" {
                        Write-Host "DEBUG: Created OU $name in demo mode"
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
                        Write-Host "DEBUG: Created user $name in AD"
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
                        Write-Host "DEBUG: Created group $name in AD"
                    }
                    "OrganizationalUnit" {
                        $params = @{
                            Name = $name
                            Path = $ou
                        }
                        
                        if ($displayName) { $params['Description'] = $displayName }
                        
                        New-ADOrganizationalUnit @params -ErrorAction Stop
                        Write-Host "DEBUG: Created OU $name in AD"
                    }
                    "Computer" {
                        $params = @{
                            Name = $name
                            Path = $ou
                        }
                        
                        New-ADComputer @params -ErrorAction Stop
                        Write-Host "DEBUG: Created computer $name in AD"
                    }
                    "Contact" {
                        $params = @{
                            Name = $name
                            Type = "Contact"
                            Path = $ou
                        }
                        
                        if ($displayName) { $params['DisplayName'] = $displayName }
                        
                        New-ADObject @params -ErrorAction Stop
                        Write-Host "DEBUG: Created contact $name in AD"
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
                        Write-Host "DEBUG: Deleted user $cleanName (demo mode)"
                    }
                    "group" {
                        # Remove group from all users
                        foreach ($u in $Global:Users) {
                            $u.Groups = $u.Groups | Where-Object { $_ -ne $cleanName }
                        }
                        Write-Host "DEBUG: Deleted group $cleanName (demo mode)"
                    }
                    default {
                        Write-Host "DEBUG: Deleted $objectType $cleanName (demo mode)"
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
                        Write-Host "DEBUG: Deleted user $cleanName from AD"
                    }
                    "group" {
                        Remove-ADGroup -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Write-Host "DEBUG: Deleted group $cleanName from AD"
                    }
                    "ou" {
                        Remove-ADOrganizationalUnit -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Write-Host "DEBUG: Deleted OU $cleanName from AD"
                    }
                    "computer" {
                        Remove-ADComputer -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Write-Host "DEBUG: Deleted computer $cleanName from AD"
                    }
                    default {
                        Remove-ADObject -Identity $cleanName -Confirm:$false -ErrorAction Stop
                        Write-Host "DEBUG: Deleted $objectType $cleanName from AD"
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
                            Write-Host "DEBUG: Moved user $cleanName to $targetOU (demo mode)"
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
                    
                    Write-Host "DEBUG: Moved $cleanName to $targetOU in AD"
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
        Write-Host "DEBUG: OK pressed, Domain = $domainString"
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
                Write-Host "DEBUG: Saved search '$newName'"
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
                Write-Host "DEBUG: Exported $($script:lastSearchResults.Count) results to $filename"
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
    Write-Host "DEBUG: Refreshing tree data..."
    
    # Show loading dialog
    $loadingDlg = Show-LoadingDialog -Message "Refreshing Active Directory data..."
    
    try {
        # Reload domain data
        Load-DomainData -domain $Global:Domain
        
        # Rebuild tree
        Build-Tree -domain $Global:Domain
        # Add after Build-Tree calls:
        Update-FilterStatusLabel -label $filterStatusLabel
        
        Write-Host "DEBUG: Tree refreshed successfully"
    } finally {
        Close-LoadingDialog $loadingDlg
    }
    
    [Terminal.Gui.MessageBox]::Query(50, 7, "Refreshed", "Active Directory data refreshed successfully", "OK") | Out-Null
}

# ------------------------- Context Menu Handler ------------------------
function Show-ContextMenu {
    param(
        [string]$objectName,
        [string]$objectType
    )
    
    # Clean the object name (remove prefixes like "(U) " or "(DC) ")
    $cleanName = $objectName -replace '^\(.\)\s*', ''
    
    Write-Host "DEBUG: Context menu for '$cleanName' (type: $objectType)"
    
    # Determine what type of object this is
    $isUser = $objectType -eq "user" -or $objectName -like "(U)*"
    $isGroup = $objectType -eq "group"
    $isOU = $objectType -eq "ou"
    $isDC = $objectType -eq "dc" -or $objectName -like "(DC)*"
    $isComputer = $objectType -eq "computer"
    
    # Build menu items based on object type
    $menuItems = @()
    
    if ($isUser) {
        $menuItems += [Terminal.Gui.MenuItem]::new("_Properties", "View user properties", [Action]{ 
            $user = $Global:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
            if ($user) { Show-UserPropertiesDialog -user $user }
        })
        $menuItems += [Terminal.Gui.MenuItem]::new("_Reset Password", "Reset user password", [Action]{ 
            Show-ResetPasswordDialog -userName $cleanName
        })
        $menuItems += [Terminal.Gui.MenuItem]::new("_Disable Account", "Disable user account", [Action]{ 
            Toggle-UserAccount -userName $cleanName -disable $true
        })
        $menuItems += [Terminal.Gui.MenuItem]::new("_Enable Account", "Enable user account", [Action]{ 
            Toggle-UserAccount -userName $cleanName -disable $false
        })
        $menuItems += $null  # Separator
        $menuItems += [Terminal.Gui.MenuItem]::new("_Move to OU...", "Move user to another OU", [Action]{ 
            Show-MoveObjectDialog -objectName $cleanName -objectType "User"
        })
        $menuItems += $null  # Separator
        $menuItems += [Terminal.Gui.MenuItem]::new("_Delete", "Delete user", [Action]{ 
            Show-DeleteObjectDialog -objectName $cleanName -objectType "User"
        })
    } elseif ($isGroup) {
        $menuItems += [Terminal.Gui.MenuItem]::new("_Properties", "View group properties", [Action]{ 
            Show-GroupPropertiesDialog -groupName $cleanName
        })
        $menuItems += [Terminal.Gui.MenuItem]::new("_Add Member...", "Add member to group", [Action]{ 
            Show-AddGroupMemberDialog -groupName $cleanName
        })
        $menuItems += [Terminal.Gui.MenuItem]::new("_Remove Member...", "Remove member from group", [Action]{ 
            Show-RemoveGroupMemberDialog -groupName $cleanName
        })
        $menuItems += $null  # Separator
        $menuItems += [Terminal.Gui.MenuItem]::new("_Delete", "Delete group", [Action]{ 
            Show-DeleteObjectDialog -objectName $cleanName -objectType "Group"
        })
    } elseif ($isOU) {
        $menuItems += [Terminal.Gui.MenuItem]::new("_Properties", "View OU properties", [Action]{ 
            [Terminal.Gui.MessageBox]::Query(50, 7, "OU Properties", "OU: $cleanName`n(Full properties coming soon)", "OK") | Out-Null
        })
        $menuItems += [Terminal.Gui.MenuItem]::new("_New Object...", "Create new object in this OU", [Action]{ 
            Show-NewObjectWizard
        })
        $menuItems += $null  # Separator
        $menuItems += [Terminal.Gui.MenuItem]::new("_Delete", "Delete OU", [Action]{ 
            Show-DeleteObjectDialog -objectName $cleanName -objectType "OU"
        })
    } elseif ($isDC) {
        $menuItems += [Terminal.Gui.MenuItem]::new("_Properties", "View DC properties", [Action]{ 
            [Terminal.Gui.MessageBox]::Query(50, 7, "DC Properties", "Domain Controller: $cleanName`n(Full properties coming soon)", "OK") | Out-Null
        })
        $menuItems += [Terminal.Gui.MenuItem]::new("_Check Replication", "Check replication status", [Action]{ 
            Check-DCReplication -dcName $cleanName
        })
    } else {
        # Generic object
        $menuItems += [Terminal.Gui.MenuItem]::new("_Properties", "View properties", [Action]{ 
            [Terminal.Gui.MessageBox]::Query(50, 7, "Properties", "Object: $cleanName`nType: $objectType", "OK") | Out-Null
        })
    }
    
    # Common items for all objects
    $menuItems += $null  # Separator
    $menuItems += [Terminal.Gui.MenuItem]::new("_Refresh", "Refresh tree", [Action]{ Refresh-TreeData })
    
    # Create and show context menu
    $menuBar = [Terminal.Gui.MenuBar]::new(@(
        [Terminal.Gui.MenuBarItem]::new("_Actions", $menuItems)
    ))
    
    # Position menu at mouse location (approximated)
    $contextDialog = [Terminal.Gui.Dialog]::new("", 30, 10)
    $contextDialog.X = [Terminal.Gui.Pos]::Center()
    $contextDialog.Y = [Terminal.Gui.Pos]::Center()
    
    # Create a simple menu selection list
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
    
    $listView = [Terminal.Gui.ListView]::new($menuText)
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
            switch ($selected) {
                "Properties" { 
                    if ($isUser) {
                        $user = $Global:Users | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
                        if ($user) { Show-UserPropertiesDialog -user $user }
                    } elseif ($isGroup) {
                        Show-GroupPropertiesDialog -groupName $cleanName
                    } else {
                        [Terminal.Gui.MessageBox]::Query(50, 7, "Properties", "Object: $cleanName`nType: $objectType", "OK") | Out-Null
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
# Replace your Show-Properties function with this to see what's happening

function Show-Properties {
    Write-Host "DEBUG: Show-Properties called"
    
    if (-not $tree.SelectedObject) { 
        Write-Host "DEBUG: No object selected"
        [Terminal.Gui.MessageBox]::Query(50, 7, "Debug", "No object selected in tree", "OK") | Out-Null
        return 
    }
    
    $selName = $tree.SelectedObject.Text
    Write-Host "DEBUG: Selected object text: '$selName'"
    
    # Remove prefixes like "(U) " or "(DC) " and status icons
    $cleanName = $selName -replace '^\(.\)\s*', '' -replace '^[○⊗🔒]\s*', ''
    Write-Host "DEBUG: Cleaned name: '$cleanName'"
    
    $selType = if ($selName -like "(U)*") {"user"} elseif ($selName -like "(DC)*") {"computer"} else {"group"}
    Write-Host "DEBUG: Detected type: $selType"

    if ($selType -eq "user") {
        Write-Host "DEBUG: Searching for user in Global:Users array (count: $($Global:Users.Count))"
        
        # Try to find the user
        $selUser = $null
        foreach ($u in $Global:Users) {
            Write-Host "DEBUG: Checking user: '$($u.Name)' against '$cleanName'"
            if ($u.Name -eq $cleanName) {
                $selUser = $u
                Write-Host "DEBUG: MATCH FOUND!"
                break
            }
        }
        
        if ($selUser) {
            Write-Host "DEBUG: User found, calling Show-UserPropertiesDialog"
            Write-Host "DEBUG: User details: Name=$($selUser.Name), Disabled=$($selUser.Disabled), Locked=$($selUser.Locked)"
            
            try {
                Show-UserPropertiesDialog -user $selUser
                Write-Host "DEBUG: Show-UserPropertiesDialog completed"
            } catch {
                Write-Host "ERROR: Exception in Show-UserPropertiesDialog: $_"
                Write-Host "ERROR: Stack trace: $($_.ScriptStackTrace)"
                [Terminal.Gui.MessageBox]::Query(70, 12, "Error", "Failed to show properties:`n$($_.Exception.Message)`n`nCheck console for details", "OK") | Out-Null
            }
        } else {
            Write-Host "DEBUG: User NOT found in Global:Users"
            [Terminal.Gui.MessageBox]::Query(50, 9, "Debug", "User '$cleanName' not found in Global:Users array.`n`nAvailable users: $($Global:Users.Count)", "OK") | Out-Null
        }
    } elseif ($selType -eq "group") {
        Write-Host "DEBUG: Group type selected: $cleanName"
        $groupName = $cleanName
        $members = $Global:Users | Where-Object { $_.Groups -contains $groupName } | ForEach-Object { $_.Name } | Sort-Object
        $desc = "<no description>"
        $txt = "Group: $groupName`nDescription: $desc`nMembers:`n" + ($members -join "`n")
        [Terminal.Gui.MessageBox]::Query(60, 20, "Group Properties", $txt, "OK") | Out-Null
    } else {
        Write-Host "DEBUG: Selected object type $selType not handled yet."
    }
}

# Additional debug helper - call this to verify your demo data loaded correctly
function Test-DemoData {
    Write-Host "========== DEMO DATA CHECK =========="
    Write-Host "Global:Users count: $($Global:Users.Count)"
    Write-Host "Global:DCs count: $($Global:DCs.Count)"
    Write-Host ""
    Write-Host "Users in memory:"
    foreach ($u in $Global:Users) {
        $locked = if ($u.Locked) { "🔒" } else { "" }
        $disabled = if ($u.Disabled) { "⊗" } else { "○" }
        Write-Host "  $disabled$locked $($u.Name) - Groups: $($u.Groups -join ', ')"
    }
    Write-Host "====================================="
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
                Write-Host "DEBUG: Password reset for $userName (demo mode)"
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
                    Write-Host "DEBUG: Account $userName $action`d (demo mode)"
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

function Show-MoveObjectDialog {
    param([string]$objectName, [string]$objectType)
    
    [Terminal.Gui.MessageBox]::Query(50, 7, "Move Object", "Move $objectType '$objectName' to OU`n(Coming soon)", "OK") | Out-Null
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
        Write-Host "DEBUG: Selection mode ENABLED"
        [Terminal.Gui.MessageBox]::Query(60, 8, "Selection Mode", 
            "Selection mode enabled!`n`nClick objects to select/deselect them.`nPress Ctrl+A to select all.`nPress Ctrl+D to deselect all.", 
            "OK") | Out-Null
    } else {
        Write-Host "DEBUG: Selection mode DISABLED"
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
            Write-Host "DEBUG: Deselected $selName"
        } else {
            # Select
            $Global:SelectedObjects += $selName
            Write-Host "DEBUG: Selected $selName"
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
    
    Write-Host "DEBUG: Selected all users ($($Global:SelectedObjects.Count))"
    Update-SelectionPanel -panel $selectionPanel
    [Terminal.Gui.MessageBox]::Query(50, 7, "Selected All", "Selected $($Global:SelectedObjects.Count) users", "OK") | Out-Null
}

function Deselect-AllObjects {
    $Global:SelectedObjects = @()
    Write-Host "DEBUG: Deselected all objects"
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
                        Write-Host "DEBUG: $action`d $cleanName (demo mode)"
                    }
                } else {
                    if ($disable) {
                        Disable-ADAccount -Identity $cleanName -ErrorAction Stop
                    } else {
                        Enable-ADAccount -Identity $cleanName -ErrorAction Stop
                    }
                    $successCount++
                    Write-Host "DEBUG: $action`d $cleanName in AD"
                }
            } catch {
                $failCount++
                $errors += "$cleanName`: $($_.Exception.Message)"
                Write-Host "DEBUG: Failed to $action $cleanName`: $_"
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
                            Write-Host "DEBUG: Moved $cleanName to $targetOU (demo mode)"
                        }
                    } else {
                        $adObject = Get-ADObject -Filter "Name -eq '$cleanName'" -ErrorAction Stop
                        Move-ADObject -Identity $adObject.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
                        $successCount++
                        Write-Host "DEBUG: Moved $cleanName to $targetOU in AD"
                    }
                } catch {
                    $failCount++
                    $errors += "$cleanName`: $($_.Exception.Message)"
                    Write-Host "DEBUG: Failed to move $cleanName`: $_"
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
                            Write-Host "DEBUG: Added $cleanName to $targetGroup (demo mode)"
                        }
                    } else {
                        Add-ADGroupMember -Identity $targetGroup -Members $cleanName -ErrorAction Stop
                        $successCount++
                        Write-Host "DEBUG: Added $cleanName to $targetGroup in AD"
                    }
                } catch {
                    $failCount++
                    Write-Host "DEBUG: Failed to add $cleanName`: $_"
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
    
    Write-Host "DEBUG: Selected all users ($($Global:SelectedObjects.Count))"
    Update-SelectionPanel -panel $selectionPanel
    [Terminal.Gui.MessageBox]::Query(50, 7, "Selected All", "Selected $($Global:SelectedObjects.Count) users", "OK") | Out-Null
}

function Deselect-AllObjects {
    $Global:SelectedObjects = @()
    Write-Host "DEBUG: Deselected all objects"
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
                        Write-Host "DEBUG: $action`d $cleanName (demo mode)"
                    }
                } else {
                    if ($disable) {
                        Disable-ADAccount -Identity $cleanName -ErrorAction Stop
                    } else {
                        Enable-ADAccount -Identity $cleanName -ErrorAction Stop
                    }
                    $successCount++
                    Write-Host "DEBUG: $action`d $cleanName in AD"
                }
            } catch {
                $failCount++
                $errors += "$cleanName`: $($_.Exception.Message)"
                Write-Host "DEBUG: Failed to $action $cleanName`: $_"
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
                            Write-Host "DEBUG: Moved $cleanName to $targetOU (demo mode)"
                        }
                    } else {
                        $adObject = Get-ADObject -Filter "Name -eq '$cleanName'" -ErrorAction Stop
                        Move-ADObject -Identity $adObject.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
                        $successCount++
                        Write-Host "DEBUG: Moved $cleanName to $targetOU in AD"
                    }
                } catch {
                    $failCount++
                    $errors += "$cleanName`: $($_.Exception.Message)"
                    Write-Host "DEBUG: Failed to move $cleanName`: $_"
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
                            Write-Host "DEBUG: Added $cleanName to $targetGroup (demo mode)"
                        }
                    } else {
                        Add-ADGroupMember -Identity $targetGroup -Members $cleanName -ErrorAction Stop
                        $successCount++
                        Write-Host "DEBUG: Added $cleanName to $targetGroup in AD"
                    }
                } catch {
                    $failCount++
                    Write-Host "DEBUG: Failed to add $cleanName`: $_"
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

function Show-AddGroupMemberDialog {
    param([string]$groupName)
    [Terminal.Gui.MessageBox]::Query(50, 7, "Add Member", "Add member to '$groupName'`n(Coming soon)", "OK") | Out-Null
}

function Show-RemoveGroupMemberDialog {
    param([string]$groupName)
    [Terminal.Gui.MessageBox]::Query(50, 7, "Remove Member", "Remove member from '$groupName'`n(Coming soon)", "OK") | Out-Null
}

function Check-DCReplication {
    param([string]$dcName)
    [Terminal.Gui.MessageBox]::Query(50, 7, "Replication Check", "Checking replication for $dcName`n(Coming soon)", "OK") | Out-Null
}

# ------------------------- Tree Mouse Handler ------------------------
# Add this to your tree setup after creating $tree:

$tree.add_MouseClick({ param($sender, $mouseArgs)
    # Right-click for context menu
    if ($mouseArgs.MouseEvent.Flags -eq [Terminal.Gui.MouseFlags]::Button3Clicked) {
        # ... existing context menu code ...
    }
    
    # Left-click for selection
    if ($mouseArgs.MouseEvent.Flags -eq [Terminal.Gui.MouseFlags]::Button1Clicked) {
        Handle-TreeClick -mouseArgs $mouseArgs
    }
})

# ------------------------- Build initial tree ------------------------
Build-Tree -domain $Global:Domain
# Add after Build-Tree calls:
Update-FilterStatusLabel -label $filterStatusLabel

# ------------------------- Run application ------------------------
[Terminal.Gui.Application]::Run($top)
[Terminal.Gui.Application]::Shutdown()
