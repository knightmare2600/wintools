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

===================================== How this code works =====================================
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

Recent changelog

3.0.0.84 (Bug fix and refactoring)
  - Fixed CSV importing and reduced unnecessary functions, replacing them with `$Script:` checks
    for Demo mode vs CSV vs Production AD.
  - Began melding repetitive code in the Show-*Properties dialogs to reduce duplication.
  - The Show-{Computer|User|Group|OU|Etc}-Properties dialogs previously contained significant duplicated
    code; this has now been cleaned up and refined.
  - Search tabs in each properties dialog have been finessed and merged with the corresponding
    *-Dialog properties.
  - LAPS dialog now correctly handles and displays both legacy and modern LAPS types.
  - DNS queries now run automatically; also fixed an `nslookup` regression.
  - CSV import now imports data, redraws the tree, and updates the information panel.
  - Handle-CSVAction fixed so refreshing CSV data also refreshes the tree automatically.
  - Reworked Refresh-Tree to squash a startup bug.
  - Fixed regression in Change DC so the correctly formatted DC list is used.
  - Retired the now-unnecessary Invoke-AD function in favour of direct AD commands.

-------------------------------------------------------------------------------
TODO / COME BACK TO
-------------------------------------------------------------------------------

REMAINING FEATURES TO IMPLEMENT:

  - Account Expiration Management
  - Audit Log Viewer (demo mode fake logs)
  - Misisng AD module is non fatal BUT if it's not installed, a global needs to not let users do stupid stuff
  - show locked users actually shows computers under maintenence, which is... nice but not what you're asking for - add both
  - Menu entries for Set-BulkAttribute and Find-StaleAccounts
    (Need to use existing helper functions for selection conversion)
  - AD Health tabs like Group Policy and domain controllers the tab pane could be smaller with a search box in them to help out
  - Start using dynamic resizes in panels and so on - slowly!

  BUGS:
  - Did the right click popup go away or is it broken...? Yes. Fix it later
  - Group Membership Report/Comparison - Report works, comparison code is broken <-- TODO NEXT

===========================================================================================
#>

param(
  [switch]$DemoMode,
  [switch]$Logging,
  [string]$LogFile,
  [string]$Domain,              ## User can specify domain
  [switch]$ImportDemoData,      ## Trigger demo data import
  [string]$DemoDataCsv,         ## Path to user's chosen CSV demo data file
  [ValidateSet("british", "class91", "dark", "db-1980s", "dsb", "gemstones", "intercity-swallow", "irn-bru", "light", "matrix", "ns", "network-southeast", "panam", "procomm", "scotrail", "twa", "viarail", "viarail-soft")]
  [string]$Theme
)

## Define the build version, project and code names once only - up here to ease patching. The rest in main
## execution loop, where they belong
$Script:ProjectName  = "DSA-TUI pwsh dsa.msc TUI"
$Script:FruitName    = "Blåbær"
$Script:BuildVersion = "3.0.0.84"

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

##################################################################################################################
## Any functions added in here, make sure to keep chronology when calling them from inside other functions...   ##
##################################################################################################################

## --------------------{ Test For Required Modules }--------------------
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
    Debug-Log "Module: '$Name' is installed. Importing..." -Type "Success"

    ## Try to import it
    try {
      Import-Module $Name -ErrorAction SilentlyContinue
    } catch {
      Debug-Log "Failed to import module: '$Name': $_" -Type "Warn"
    }

    return $true
  }
  else {
    if ($Optional) {
      Debug-Log "Optional module: '$Name' is NOT installed." -Type "Warn"
    } else {
      Debug-Log "Module '$Name' is NOT installed. Please run: Install-Module -Name $Name" -Type "Error"
    }
    Debug-Log "If you would like terminal icons, be sure to install this: https://github.com/jpawlowski/nerd-fonts-installer-PS" -Type "Info"
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

## -------------------------{ Debug Logging }-------------------------
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
    'Debug'   { $emoji = '🔧'; $color = 'Cyan'}    # Spanner (Wrench for left pondians)
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

function Build-MainMenu {
  [CmdletBinding()]
  param()

  Debug-Log ": Building main menu..." -Type "Info"

  ## -------------------------{ Menu Items }-------------------------
  $mFile         = [Terminal.Gui.MenuItem]::new("_Exit","Exit application (F10)",[Action]{ [Terminal.Gui.Application]::RequestStop() })
  $mNew          = [Terminal.Gui.MenuItem]::new("New Object","Create a new object (F3)",[Action]{ Show-NewObjectWizard })
  $mProps        = [Terminal.Gui.MenuItem]::new("_Properties","Edit selected properties",[Action]{ Show-Properties })

  ## In Build-MainMenu function:
  $mDemoExport   = [Terminal.Gui.MenuItem]::new("_Export Demo Data", "Export demo data to CSV", [Action]{ Handle-CSVAction  -Action 'Export' })
  $mDemoImport   = [Terminal.Gui.MenuItem]::new("_Import Demo Data", "Import demo data from CSV", [Action]{ $result = Handle-CSVAction -Action 'Import'
    if ($result) {
      ## Rebuild tree with imported data
      $rootNode = Build-Tree -domain $Script:CurrentDomain
      if ($rootNode) {
        $Script:tree.ClearObjects()
        $Script:tree.AddObject($rootNode)
        $Script:tree.SetNeedsDisplay()
        Debug-Log "Tree rebuilt after CSV import" -Type "Success"
      }
      Show-InfoPanel -UpdateOnly
    }
  })

  $mUndo         = [Terminal.Gui.MenuItem]::new("_Undo","Undo last action",[Action]{ Debug-Log (": Undo placeholder") -Type "Info" })
  $mChangeDomain = [Terminal.Gui.MenuItem]::new("Change _Domain","Select domain",[Action]{ Show-ChangeDomainDialog })
  $mChangeDC     = [Terminal.Gui.MenuItem]::new("Change _Domain Controller","Select DC",[Action]{ Show-ChangeDCDialog })
  $mSearchAD     = [Terminal.Gui.MenuItem]::new("_Search AD","Search Active Directory (F7)",[Action]{ Show-ADSearchDialog })
  $mRefresh      = [Terminal.Gui.MenuItem]::new("_Refresh","Refresh AD data (F5)",[Action]{
    Debug-Log (": Refresh menu clicked - scheduling refresh...") -Type "Info"
    [Terminal.Gui.Application]::MainLoop.AddTimeout([TimeSpan]::FromMilliseconds(100), {
    Debug-Log (": Timeout callback - starting refresh...") -Type "Info"
    try {
      Set-StatusBar "Refreshing..." -spinner
      $result = Refresh-Data -domain $Script:CurrentDomain -RebuildTree
      if ($result) { Set-StatusBar "Refresh complete" -final } else { Set-StatusBar "Refresh failed" -final }
        Debug-Log (": Refresh completed with result: $result") -Type "Info"
      } catch {
        Debug-Log (": Refresh crashed: $($_.Exception.Message)") -Type "Info"
        Set-StatusBar "Refresh error" -final
      }
      return $false
    })
    Debug-Log (": Refresh scheduled") -Type "Info"
  })

  $mQuickFilter       = [Terminal.Gui.MenuItem]::new("_Quick Filter","Apply quick filters",[Action]{Show-QuickFilterDialog})
  $mSelectionMode     = [Terminal.Gui.MenuItem]::new("_Selection Mode (Ctrl+S)","Toggle selection mode",[Action]{Toggle-SelectionMode})
  $mSelectAll         = [Terminal.Gui.MenuItem]::new("Select _All (Ctrl+A)","Select all objects",[Action]{Manage-Selection -Action 'SelectAll'})
  $mDeselectAll       = [Terminal.Gui.MenuItem]::new("_Deselect All (Ctrl+D)","Deselect all objects",[Action]{Manage-Selection -Action 'DeselectAll'})
  $mBulkDisable       = [Terminal.Gui.MenuItem]::new("_Disable Selected", "Disable selected accounts", [Action]{$objects = Get-SelectedObjectsAsObjects ; Manage-AccountStatus -Objects $objects -Action 'Disable'})
  $mBulkEnable        = [Terminal.Gui.MenuItem]::new("_Enable Selected", "Enable selected accounts", [Action]{$objects = Get-SelectedObjectsAsObjects ; Manage-AccountStatus -Objects $objects -Action 'Enable'})
  $mBulkAddGroup      = [Terminal.Gui.MenuItem]::new("Add to _Group...","Add selected users to group",[Action]{Invoke-BulkAddToGroup})
  $mBulkEdit          = [Terminal.Gui.MenuItem]::new("_Edit Attribute (Bulk)", "Change one attribute across selected objects", [Action]{$objects = Get-SelectedObjectsAsObjects ; Set-BulkAttribute -Objects $objects -ShowDialog})
  $mPasswordGenerator = [Terminal.Gui.MenuItem]::new("_Password Generator","Password Generator (F2)",[Action]{Generate-RandomPassword})
  $mLAPSPasswords     = [Terminal.Gui.MenuItem]::new("_LAPS Passwords","Lookup LAPS Creds (Fx)",[Action]{Show-LAPSSearchModal})
  $mADHealth          = [Terminal.Gui.MenuItem]::new("_AD Health Status", "AD Health And Replication Status", [Action]{Show-ADHealthDialog })
  $mShortcuts         = [Terminal.Gui.MenuItem]::new("_Shortcuts","Keyboard shortcuts (F1)",[Action]{Show-Modal "Shortcuts" "F1 - Help`nF2 - Password Generator`nF3 - New`nF5 - Refresh`nF6 - Themes`nF7 - Search`nF10 - Quit" })
  $mAboutDSATUI       = [Terminal.Gui.MenuItem]::new("_About","About $($Script:ProjectName)",[Action]{Show-Modal "About" "$($Script:ProjectName)`n`nCodename: $($Script:FruitName)`nv$($Script:BuildVersion) STABLE`nGPL-3 Copyleft`nBy Knightmare2600 (https://github.com/knightmare2600)" })
  $mWhyBlaabaer       = [Terminal.Gui.MenuItem]::new("Why _Blaabaer?","Why the $($Script:FruitName) codename?",[Action]{Show-BlaabaerInfo })
  $mTheme             = [Terminal.Gui.MenuItem]::new("_Theme","Change color theme (F6)",[Action]{Show-ThemeSelector })
  $menuItemGPOs       = [Terminal.Gui.MenuItem]::new("Group _Policy Objects", "Show Group Policies", {Show-GPOListDialog -Domain $Script:CurrentDomain})
  $menuItemTrusts     = [Terminal.Gui.MenuItem]::new("Trust _Relationships", "Show Trust Relationships", {Show-TrustsDialog -Domain $Script:CurrentDomain})
  $menuItemIPSec      = [Terminal.Gui.MenuItem]::new("_IPSec Policies", "Show IPSec Policies", {Show-IPSecPoliciesDialog -Domain $Script:CurrentDomain})
  $menuItemIPSecHelp  = [Terminal.Gui.MenuItem]::new("_IPSec Help", "IPSec Policies Help", {Show-IPSecHelpDialog})
  $menuItemPrinters   = [Terminal.Gui.MenuItem]::new("_Print Queues", "Show Printer Queues", {Show-PrintQueuesDialog -Domain $Script:CurrentDomain})
  $menuItemDNSDialog  = [Terminal.Gui.MenuItem]::new("_DNS Manager", "AD DNS Manager", { Show-DNSDialog })
  ## Submenu: Actions > AD Properties
  $mCopyTemplate      = [Terminal.Gui.MenuItem]::new("Copy as _Template", "Create new object from selected object", [Action]{
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

  ## -------------------------{ Menu Bar }-------------------------
  ## Be mindful of nested menus here!
  $menu = [Terminal.Gui.MenuBar]::new(@(
    $mExtraFunctions = [Terminal.Gui.MenuBarItem]::new("_AD Properties",@($mPasswordGenerator, $mLAPSPasswords, $mCopyTemplate, $menuItemGPOs, $menuItemTrusts, $menuItemPrinters, $menuItemDNSDialog, $menuItemIPSec))
    [Terminal.Gui.MenuBarItem]::new("_File", @($mRefresh, $mDemoExport, $mDemoImport, $mTheme, $mFile)),
    [Terminal.Gui.MenuBarItem]::new("_Action", @($mNew, $mProps, $mExtraFunctions, $mQuickFilter, $mUndo, $mChangeDomain, $mChangeDC, $mSearchAD, $mADHealth)),
    [Terminal.Gui.MenuBarItem]::new("_Selection", @($mSelectionMode, $mSelectAll, $mDeselectAll, $mBulkEdit, $mBulkAddGroup, $mBulkEnable, $mBulkDisable)),
    [Terminal.Gui.MenuBarItem]::new("_Help", @($mShortcuts, $menuItemIPSecHelp, $mAboutDSATUI, $mWhyBlaabaer))
  ))
  Debug-Log ": Main menu created successfully" -Type "Info"
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
        $State.chkSearchLocked = [Terminal.Gui.CheckBox]::new("Account Locked")
        $State.chkSearchLocked.X = 2
        $State.chkSearchLocked.Y = [Terminal.Gui.Pos]::Bottom($State.txtSearchOutput) + 1
        $State.chkSearchLocked.CanFocus = $true
        $View.Add($State.chkSearchLocked)

        $State.chkSearchDisabled = [Terminal.Gui.CheckBox]::new("Account Disabled")
        $State.chkSearchDisabled.X = 2
        $State.chkSearchDisabled.Y = [Terminal.Gui.Pos]::Bottom($State.chkSearchLocked) + 1
        $State.chkSearchDisabled.CanFocus = $true
        $View.Add($State.chkSearchDisabled)
      } else {
        if ($State.chkSearchLocked) {
          $State.chkSearchLocked.Checked = [bool]($Data.Locked ?? $Data.LockedOut)
        }
        if ($State.chkSearchDisabled) {
          $State.chkSearchDisabled.Checked = [bool]($Data.Disabled ?? -not $Data.Enabled)
        }
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
        if ($State.chkSearchGC) {
          $State.chkSearchGC.Checked = [bool]$Data.IsGlobalCatalog
        }
        if ($State.chkSearchRODC) {
          $State.chkSearchRODC.Checked = [bool]$Data.IsReadOnly
        }
      }
    }
  }
}

function Show-Properties {
  Debug-Log ": Show-Properties called" -Type "Info"
  $node = $Script:tree.SelectedObject
  if (-not $node) {
    Debug-Log ": No object selected" -Type "Info"
    Show-Modal "Debug" "No object selected in tree"
    return
  }
  Debug-Log ": Selected node text: '$($node.Text)'" -Type "Info"
  $tag = $node.Tag
  Debug-Log ": Tag is null: $($null -eq $tag)" -Type "Info"
  if ($tag) {
    Debug-Log ": Tag.Type: '$($tag.Type)'" -Type "Info"
    Debug-Log ": Tag.Object is null: $($null -eq $tag.Object)" -Type "Info"
    if ($tag.Object) {
      Debug-Log ": Tag.Object type: $($tag.Object.GetType().Name)" -Type "Info"
    }
  }

  ## Get the actual AD object from the Tag
  $obj = $tag.Object

  ## ---- Handle containers without objects ----
  if (-not $obj) {
    if ($tag.Type -eq 'container') {
      Debug-Log ": Container selected (no properties to show)" -Type "Info"
      Show-Modal "Container" "This is a container node.`n`nSelect an individual object to view its properties."
      return
    }
    Debug-Log ": No object attached to this node" -Type "Warn"
    return
  }

  ## ---- Use the Type property first (more reliable) ----
  switch ($tag.Type) {
    'user' {
      Debug-Log ": USER object selected: $($obj.Name)" -Type "Info"
      Show-UserPropertiesDialog -user $obj
      return
    }
    'group' {
      Debug-Log ": GROUP object selected: $($obj.Name)" -Type "Info"
      Show-GroupPropertiesDialog -group $obj
      return
    }
    'dc' {
      Debug-Log ": DC object selected: $($obj.Name)" -Type "Info"
      Show-DCPropertiesDialog -dc $obj
      return
    }
    'dc-container' {
      Debug-Log ": DC container selected (no properties to show)" -Type "Info"
      Show-Modal "Domain Controllers" "This is a container for Domain Controllers in this domain`n`nSelect an individual DC to view its properties."
      return
    }
    'computer' {
      Debug-Log ": COMPUTER object selected: $($obj.Name)" -Type "Info"
      Show-ComputerPropertiesDialog -computerName $obj.Name
      return
    }
    'ou' {
      Debug-Log ": Showing OU properties for $($obj.Name)" -Type "Info"
      Show-OUPropertiesDialog -ou $obj
      return
    }
    'container' {
      Debug-Log ": Container selected (no properties to show)" -Type "Info"
      Show-Modal "Container" "This is a container node.`n`nSelect an individual object to view its properties."
      return
    }
  }

  ## ---- Fallback: Try to detect type from properties (for backward compatibility) ----
  if ($obj.PSObject.Properties.Match('SamAccountName').Count -gt 0) {
    Debug-Log ": Detected USER object (fallback): $($obj.Name)" -Type "Info"
    Show-UserPropertiesDialog -user $obj
    return
  }
  if ($obj.PSObject.Properties.Match('GroupScope').Count -gt 0 -or
    $obj.PSObject.Properties.Match('Members').Count -gt 0) {
    Debug-Log ": Detected GROUP object (fallback): $($obj.Name)" -Type "Info"
    Show-GroupPropertiesDialog -group $obj
    return
  }
  if ($obj.PSObject.Properties.Match('Site').Count -gt 0) {
    Debug-Log ": Detected DC object (fallback): $($obj.Name)" -Type "Info"
    Show-DCPropertiesDialog -dc $obj
    return
  }
  if ($obj.PSObject.Properties.Match('OperatingSystem').Count -gt 0) {
    Debug-Log ": Detected COMPUTER object (fallback): $($obj.Name)" -Type "Info"
    Show-ComputerPropertiesDialog -computer $obj
    return
  }

  Debug-Log ": Selected object type unknown, cannot show properties" -Type "Warn"
}

function Show-UserPropertiesDialog {
  param($user)

  if (-not $user) {
    Debug-Log ": User object is null" -Type "Warn"
    return
  }
  Debug-Log ": Show-UserPropertiesDialog starting for: $($user.Name)" -Type "Info"

  ## ==================== General Tab ====================
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
      $hasExpiration = $false
      $expirationDate = $null

      try {
        if ($user.PSObject.Properties.Match('AccountExpirationDate') -and $user.AccountExpirationDate) {
          $hasExpiration = $true
          $expirationDate = $user.AccountExpirationDate
        }
      } catch {}

      $state.chkNeverExpires = [Terminal.Gui.CheckBox]::new("Never Expires")
      $state.chkNeverExpires.X = 4; $state.chkNeverExpires.Y = $y
      $state.chkNeverExpires.Checked = -not $hasExpiration
      $view.Add($state.chkNeverExpires)
      $y += 1

      $lbl = [Terminal.Gui.Label]::new("Expires On:")
      $lbl.X = 4; $lbl.Y = $y
      $view.Add($lbl)

      $expirationDateStr = if ($expirationDate) { $expirationDate.ToString("yyyy-MM-dd") } else { "" }
      $state.txtExpirationDate = [Terminal.Gui.TextField]::new($expirationDateStr)
      $state.txtExpirationDate.X = 20; $state.txtExpirationDate.Y = $y; $state.txtExpirationDate.Width = 20
      $state.txtExpirationDate.ReadOnly = $state.chkNeverExpires.Checked
      $view.Add($state.txtExpirationDate)

      $lblFmt = [Terminal.Gui.Label]::new("(yyyy-MM-dd)")
      $lblFmt.X = 42; $lblFmt.Y = $y
      $view.Add($lblFmt)
      $y += 1

      $lbl = [Terminal.Gui.Label]::new("Days Until Expiry:")
      $lbl.X = 4; $lbl.Y = $y
      $view.Add($lbl)

      $state.lblDaysUntilExpiry = [Terminal.Gui.Label]::new("N/A")
      $state.lblDaysUntilExpiry.X = 20; $state.lblDaysUntilExpiry.Y = $y
      $view.Add($state.lblDaysUntilExpiry)

      ## Event handlers
      $null = $state.chkNeverExpires.add_Toggled({
        $state.txtExpirationDate.ReadOnly = $state.chkNeverExpires.Checked
        if ($state.chkNeverExpires.Checked) {
          $state.txtExpirationDate.Text = [NStack.ustring]::Make("")
          $state.lblDaysUntilExpiry.Text = [NStack.ustring]::Make("N/A")
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

  ## ==================== Account Tab ====================
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
    }
  }

  ## ==================== Address Tab ====================
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

  ## ==================== Profile Tab ====================
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

  ## ==================== Organization Tab ====================
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

  ## ==================== Member Of Tab ====================
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

      ## FORCE ARRAY: Use @() to ensure it's always an array
      $state.groupList = @()
      if ($user.Groups) {
        $state.groupList = @($user.Groups)
      } elseif ($user.MemberOf) {
        $state.groupList = @($user.MemberOf | ForEach-Object { if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ } })
      }

      if ($state.groupList.Count -gt 0) {
        $state.lstGroups.SetSource($state.groupList)
      } else {
        $state.lstGroups.SetSource(@("(No group memberships)"))
      }
      $view.Add($state.lstGroups)

      $btnAdd = [Terminal.Gui.Button]::new("Add to Group...")
      $btnAdd.X = 2; $btnAdd.Y = 28
      $btnAdd.add_Clicked({
        Show-EditGroupMembershipDialog -User $user -OnUpdate {
          $refreshedGroups = @()
          if ($user.Groups) {
            $refreshedGroups = @($user.Groups)
          } elseif ($user.MemberOf) {
            $refreshedGroups = @($user.MemberOf | ForEach-Object { if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ } })
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
      $btnRemove.X = 22; $btnRemove.Y = 28
      $btnRemove.add_Clicked({
        $selectedIndex = $state.lstGroups.SelectedItem

        ## FORCE ARRAY
        $currentGroups = @()
        if ($user.Groups) {
          $currentGroups = @($user.Groups)
        } elseif ($user.MemberOf) {
          $currentGroups = @($user.MemberOf | ForEach-Object {
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
              if ($Script:DemoMode) {
                $user.Groups = @($user.Groups | Where-Object { $_ -ne $selectedGroup })
              } else {
                Remove-ADGroupMember -Identity $selectedGroup -Members $user.SamAccountName -Confirm:$false
              }

              ## FORCE ARRAY
              $updatedGroups = @()
              if ($user.Groups) {
                $updatedGroups = @($user.Groups)
              } elseif ($user.MemberOf) {
                $updatedGroups = @($user.MemberOf | Where-Object { $_ -ne $selectedGroup } | ForEach-Object {
                  if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
                })
              }

              if ($updatedGroups.Count -gt 0) {
                $state.lstGroups.SetSource($updatedGroups)
              } else {
                $state.lstGroups.SetSource(@("(No group memberships)"))
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

  ## ==================== Apply Logic ====================
  $applyLogic = {
    param($user, $state)
    try {
      $changesMade = $false

      if ($state.txtSamAccountName) {
        $newSamAccountName = $state.txtSamAccountName.Text.ToString().Trim()
        if ($newSamAccountName -ne $user.SamAccountName -and -not [string]::IsNullOrWhiteSpace($newSamAccountName)) {
          if ($Script:DemoMode) {
            $user.SamAccountName = $newSamAccountName
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
          if ($Script:DemoMode) {
            $user.UserPrincipalName = $newUPN
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
          if ($Script:DemoMode) {
            $user.DisplayName = $newDisplayName
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

      if ($changesMade) {
        Show-Modal "Success" "Changes applied successfully"
      } else {
        Show-Modal "Info" "No changes to apply"
      }
    } catch {
      Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)"
    }
  }

  ## ==================== Create Dialog ====================
  ## Note: Search tab is auto-added by New-PropertiesDialog!
  $tabs = @($generalTab, $accountTab, $addressTab, $profileTab, $organizationTab, $memberOfTab)
  New-PropertiesDialog -Title "User Properties - $($user.Name)" -Width 100 -Height 40 -Tabs $tabs -Data $user -OnApply $applyLogic -IncludeSearchTab $true -SearchTabConfig @{ObjectType='User'}
}

##  -------------------------{ Helper functions for property dialogs  }-------------------------
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
  param(
    [Parameter(Mandatory)]$View,
    [Parameter(Mandatory)][ref]$Y,
    [Parameter(Mandatory)][string]$Text,
    [int]$X = 2,
    [int]$SpaceBefore = 1,
    [int]$SpaceAfter = 2
  )
  $Y.Value += $SpaceBefore
  $lbl = [Terminal.Gui.Label]::new($Text)
  $lbl.X = $X
  $lbl.Y = $Y.Value
  $View.Add($lbl)
  $Y.Value += $SpaceAfter
}

function New-SearchTab {
  param([object]$Data, [hashtable]$Config)

  $objectType = if ($Config.ObjectType) {
    $Config.ObjectType
  } elseif ($Data.ObjectClass -eq 'user') { 'User'
  } elseif ($Data.ObjectClass -eq 'group') { 'Group'
  } elseif ($Data.ObjectClass -eq 'computer') { 'Computer'
  } else { 'Object'
  }

  $searchTypes = $Config.SearchTypes ?? @("$objectType", "Group", "User", "Computer", "OU")

  return @{
    Name = "Search/Lookup"
    Builder = {
      param($view, $data, $state)
      $y = 1

      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Domain:" -FieldName 'txtSearchDomain' -State $state -Value $Script:CurrentDomain -Width 30
      $nameLabel = "$objectType Name:"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label $nameLabel -FieldName 'txtSearchName' -State $state -Value $data.Name -Width 30

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
      $lblFilter.X = 48; $lblFilter.Y = 1
      $view.Add($lblFilter)

      $state.txtSearchFilter = [Terminal.Gui.TextField]::new("")
      $state.txtSearchFilter.X = 62; $state.txtSearchFilter.Y = 1; $state.txtSearchFilter.Width = 20
      $view.Add($state.txtSearchFilter)

      $state.txtSearchFilter.add_TextChanged({
        if ($state.currentSearchOutputLines) {
          $search = $state.txtSearchFilter.Text.ToString().Trim()
          if ($search) { $state.txtSearchOutput.Text = ($state.currentSearchOutputLines | Where-Object { $_ -match "(?i)$search" }) -join "`n"
          } else { $state.txtSearchOutput.Text = $state.currentSearchOutputLines -join "`n" }
        }
      }.GetNewClosure())

      $lblResult = [Terminal.Gui.Label]::new("Properties:")
      $lblResult.X = 2; $lblResult.Y = $y
      $view.Add($lblResult)
      $y += 1

      $state.txtSearchOutput = [Terminal.Gui.TextView]::new()
      $state.txtSearchOutput.X = 2; $state.txtSearchOutput.Y = $y
      $state.txtSearchOutput.Width = [Terminal.Gui.Dim]::Fill(2)
      $state.txtSearchOutput.Height = [Terminal.Gui.Dim]::Fill(4)
      $state.txtSearchOutput.ReadOnly = $true
      $state.txtSearchOutput.WordWrap = $false
      $view.Add($state.txtSearchOutput)

      script:Set-ObjectCheckboxes -View $view -State $state -Data $data -ObjectType $objectType -Mode 'Create'

      $btnSearch = [Terminal.Gui.Button]::new("Search")
      $btnSearch.X = 48; $btnSearch.Y = 3
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

          script:Set-ObjectCheckboxes -State $state -Data $data -ObjectType $objectType -Mode 'Update'
        }
      }.GetNewClosure())
    }
  }
}
### end of known source of truth

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
    Debug-Log ": New-PropertiesDialog called for: $Title" -Type "Debug"

    ## CAPTURE FUNCTIONS FOR CLOSURES
    $debugLogFunc = ${function:Debug-Log}
    $showModalFunc = ${function:Show-Modal}

    ## Shared state
    $sharedState = @{}

    ## Auto-add search tab if requested
    if ($IncludeSearchTab) {
      Debug-Log ": Adding search tab to tabs array" -Type "Debug"
      $searchTab = New-SearchTab -Data $Data -Config $SearchTabConfig
      $Tabs += $searchTab
    }

    ## Create buttons
    $btnOK = [Terminal.Gui.Button]::new("OK")
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnApply = [Terminal.Gui.Button]::new("Apply")

    ## Button handlers with captured functions
    $btnOK.add_Clicked({
      & $debugLogFunc ": OK clicked" -Type "Info"
      if ($OnOK) { & $OnOK $Data $sharedState }
      [Terminal.Gui.Application]::RequestStop()
    }.GetNewClosure())

    $btnCancel.add_Clicked({
      & $debugLogFunc ": Cancel clicked" -Type "Info"
      [Terminal.Gui.Application]::RequestStop()
    }.GetNewClosure())

    $btnApply.add_Clicked({
      & $debugLogFunc ": Apply clicked" -Type "Info"
      if ($OnApply) {
        try {
          & $OnApply $Data $sharedState
        } catch {
          & $debugLogFunc ": Apply failed: $($_.Exception.Message)" -Type "Error"
          [Terminal.Gui.Application]::MainLoop.Invoke({
            & $showModalFunc "Error" "Failed to apply changes:`n$($_.Exception.Message)"
          })
        }
      }
    }.GetNewClosure())

    ## Create dialog with buttons
    Debug-Log ": Creating dialog with buttons" -Type "Debug"
    $dialog = [Terminal.Gui.Dialog]::new($Title, $Width, $Height, $btnOK, $btnCancel, $btnApply)

    ## Create TabView
    Debug-Log ": Creating TabView" -Type "Debug"
    $tabView = [Terminal.Gui.TabView]::new()
    $tabView.X = 0
    $tabView.Y = 0
    $tabView.Width = [Terminal.Gui.Dim]::Fill()
    $tabView.Height = [Terminal.Gui.Dim]::Fill(1)

    ## Build each tab
    foreach ($tabDef in $Tabs) {
      Debug-Log ": Creating tab: $($tabDef.Name)" -Type "Info"

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
          Debug-Log ": Builder completed for tab: $($tabDef.Name)" -Type "Debug"
        } catch {
          Debug-Log ": Error building tab $($tabDef.Name): $($_.Exception.Message)" -Type "Error"
          Debug-Log ": Error line: $($_.InvocationInfo.ScriptLineNumber)" -Type "Error"
          Debug-Log ": Stack: $($_.ScriptStackTrace)" -Type "Error"
          throw
        }
      }

      $tab.View = $view
      $tabView.AddTab($tab, $false)
    }

    Debug-Log ": Adding TabView to dialog" -Type "Debug"
    $dialog.Add($tabView)

    Debug-Log ": All tabs added, running dialog" -Type "Success"
    [Terminal.Gui.Application]::Run($dialog)
    Debug-Log ": Dialog closed normally" -Type "Info"

  } catch {
    Debug-Log ": Exception in New-PropertiesDialog: $($_.Exception.Message)" -Type "Error"
    Debug-Log ": Stack: $($_.ScriptStackTrace)" -Type "Error"
    Show-Modal "Error" "Failed to display dialog:`n$($_.Exception.Message)"
  }
}

## -----------------------{ Pretty Themes Selection }----------------------
function Show-ThemeSelector {
  ## --- Theme list ---
  $themes = @("british", "class91", "dark", "db-1980s", "dsb", "gemstones", "intercity-swallow", "irn-bru", "light", "matrix", "ns", "network-southeast", "panam", "procomm", "scotrail", "twa", "viarail", "viarail-soft")

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
  $rdoLeft.add_SelectedItemChanged({ if ($rdoLeft.SelectedItem -ge 0) { $rdoRight.SelectedItem = -1 } })
  $rdoRight.add_SelectedItemChanged({ if ($rdoRight.SelectedItem -ge 0) { $rdoLeft.SelectedItem = -1  }})

  ## --- Apply Button ---
  $btnApply = [Terminal.Gui.Button]::new("Apply")
  $btnApply.add_Clicked({
    $sel = if ($rdoLeft.SelectedItem -ge 0) { $leftThemes[$rdoLeft.SelectedItem]
    } elseif ($rdoRight.SelectedItem -ge 0) { $rightThemes[$rdoRight.SelectedItem]
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
  }).GetNewClosure()
  $dlg.AddButton($btnApply)

  ## --- Cancel Button ---
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
  $dlg.AddButton($btnCancel)

  ## --- Run the dialog ---
  [Terminal.Gui.Application]::Run($dlg)
}

## ==================== STALE ACCOUNT DETECTION ====================
function Find-StaleAccounts {
  <#
  .SYNOPSIS
    Find accounts that haven't been used in a specified number of days
  .PARAMETER DaysSinceLogon
    Number of days since last logon to consider account stale
  .PARAMETER ObjectType
    'User', 'Computer', or 'Both'

  .PARAMETER ShowDialog
  If specified, shows interactive dialog with results

  .PARAMETER ExportPath
  Optional - export results to CSV

  .EXAMPLE
  Find-StaleAccounts -DaysSinceLogon 90 -ObjectType 'User' -ShowDialog

  .EXAMPLE
  Find-StaleAccounts -DaysSinceLogon 180 -ObjectType 'Computer' -ExportPath "stale_computers.csv"

  .EXAMPLE
  Find-StaleAccounts -DaysSinceLogon 90 -ObjectType 'Both' -ShowDialog
  #>

  param(
    [Parameter(Mandatory=$false)]
    [int]$DaysSinceLogon = 90,
    [Parameter(Mandatory=$false)]
    [ValidateSet('User', 'Computer', 'Both')]
    [string]$ObjectType = 'User',
    [switch]$ShowDialog,
    [string]$ExportPath
  )

  ## ==================== PARAMETER SELECTION DIALOG ====================
  if ($ShowDialog -and -not $PSBoundParameters.ContainsKey('DaysSinceLogon')) {
    ## Show dialog to select parameters
    $paramDlg = [Terminal.Gui.Dialog]::new("Stale Account Detection", 60, 16)

    $lblTitle = [Terminal.Gui.Label]::new("Configure stale account search:")
    $lblTitle.X = 2
    $lblTitle.Y = 1
    $paramDlg.Add($lblTitle)

    ## Days threshold
    $lblDays = [Terminal.Gui.Label]::new("Days since last logon:")
    $lblDays.X = 2
    $lblDays.Y = 3
    $paramDlg.Add($lblDays)

    $txtDays = [Terminal.Gui.TextField]::new("90")
    $txtDays.X = 25
    $txtDays.Y = 3
    $txtDays.Width = 10
    $paramDlg.Add($txtDays)

    ## Object type selection
    $lblType = [Terminal.Gui.Label]::new("Search for:")
    $lblType.X = 2
    $lblType.Y = 5
    $paramDlg.Add($lblType)

    $rdoType = [Terminal.Gui.RadioGroup]::new(@("Users", "Computers", "Both"))
    $rdoType.X = 2
    $rdoType.Y = 6
    $rdoType.SelectedItem = 0  # Users by default
    $paramDlg.Add($rdoType)

    ## Search button
    $btnSearch = [Terminal.Gui.Button]::new("Search")
    $btnSearch.add_Clicked({
      ## Validate days input
      $daysValue = 90
      if (-not [int]::TryParse($txtDays.Text.ToString(), [ref]$daysValue) -or $daysValue -lt 1) {
        Show-Modal "Invalid Input" "Days must be a positive number"
        return
      }

      ## Get object type
      $typeValue = switch ($rdoType.SelectedItem) {
        0 { 'User' }
        1 { 'Computer' }
        2 { 'Both' }
      }

      [Terminal.Gui.Application]::RequestStop()

      ## Run search with selected parameters
      Find-StaleAccounts -DaysSinceLogon $daysValue -ObjectType $typeValue -ShowDialog
    }.GetNewClosure())

    $paramDlg.AddButton($btnSearch)

    ## Cancel button
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
    $paramDlg.AddButton($btnCancel)
    [Terminal.Gui.Application]::Run($paramDlg)
    return
  }
  Debug-Log ": Starting stale account detection - $ObjectType, $DaysSinceLogon days threshold" -Type "Info"
  $cutoffDate = (Get-Date).AddDays(-$DaysSinceLogon)
  $staleAccounts = @()

  ## ==================== FIND STALE USERS ====================
  if ($ObjectType -eq 'User' -or $ObjectType -eq 'Both') {
    if ($Script:DemoMode) {
      ## Demo mode - use simulated LastLogonDate
      foreach ($user in $Script:Users) {
        ## Generate pseudo-random last logon based on name hash
        $hash = 0
        foreach ($char in $user.Name.ToCharArray()) { $hash += [int]$char }
        $daysAgo = ($hash % 365) + 1  # 1-365 days ago
        $lastLogon = (Get-Date).AddDays(-$daysAgo)

        if ($lastLogon -lt $cutoffDate) {
          $staleAccounts += [PSCustomObject]@{
          ObjectType = 'User'
          Name = $user.Name
          SamAccountName = $user.SamAccountName
          Enabled = $user.Enabled
          Disabled = $user.Disabled
          LastLogon = $lastLogon
          DaysSinceLogon = $daysAgo
          OU = ($user.OU -join ';')
          Department = $user.Department
          Title = $user.Title
        }
      }
    }
    } else {
    ## Production mode - query AD
    try {
      $lastLogonDate = $cutoffDate.ToFileTime()
      $users = Get-ADUser -Filter "LastLogonTimeStamp -lt $lastLogonDate" -Properties LastLogonTimeStamp,Department,Title,Enabled
      foreach ($user in $users) {
      $lastLogon = if ($user.LastLogonTimeStamp) {
        [DateTime]::FromFileTime($user.LastLogonTimeStamp)
      } else {
        [DateTime]::MinValue
      }

      $daysAgo = ((Get-Date) - $lastLogon).Days
      $staleAccounts += [PSCustomObject]@{
        ObjectType = 'User'
        Name           = $user.Name
        SamAccountName = $user.SamAccountName
        Enabled        = $user.Enabled
        Disabled       = -not $user.Enabled
        LastLogon      = $lastLogon
        DaysSinceLogon = $daysAgo
        OU             = $user.DistinguishedName -replace '^CN=[^,]+,', ''
        Department     = $user.Department
        Title          = $user.Title
        }
      }
    } catch {
      Debug-Log ": Error querying stale users: $($_.Exception.Message)" -Type "Error"
      Show-Modal "Query Error" "Failed to query stale users:`n`n$($_.Exception.Message)"
      return
    }
  }
  }

  ## ==================== FIND STALE COMPUTERS ====================
  if ($ObjectType -eq 'Computer' -or $ObjectType -eq 'Both') {
    if ($Script:DemoMode) {
      ## Demo mode - use simulated LastLogonDate
      foreach ($computer in $Script:Computers) {
        ## Generate pseudo-random last logon based on name hash
        $hash = 0
        foreach ($char in $computer.Name.ToCharArray()) { $hash += [int]$char }
        $daysAgo   = ($hash % 365) + 1  # 1-365 days ago
        $lastLogon = (Get-Date).AddDays(-$daysAgo)

        if ($lastLogon -lt $cutoffDate) {
          $staleAccounts += [PSCustomObject]@{
          ObjectType      = 'Computer'
          Name            = $computer.Name
          SamAccountName  = $computer.SamAccountName
          Enabled         = $computer.Enabled
          Disabled        = $computer.Disabled
          LastLogon       = $lastLogon
          DaysSinceLogon  = $daysAgo
          OU              = 'Computers'
          OperatingSystem = $computer.OperatingSystem
          Location        = $computer.Location
        }
      }
    }
  } else {
    ## Production mode - query AD
    try {
      $lastLogonDate = $cutoffDate.ToFileTime()
      $computers = Get-ADComputer -Filter "LastLogonTimeStamp -lt $lastLogonDate" -Properties LastLogonTimeStamp,OperatingSystem,Location,Enabled
      foreach ($computer in $computers) {
        $lastLogon = if ($computer.LastLogonTimeStamp) {
        [DateTime]::FromFileTime($computer.LastLogonTimeStamp)
      } else {
        [DateTime]::MinValue
      }

      $daysAgo = ((Get-Date) - $lastLogon).Days
      $staleAccounts += [PSCustomObject]@{
        ObjectType      = 'Computer'
        Name            = $computer.Name
        SamAccountName  = $computer.SamAccountName
        Enabled         = $computer.Enabled
        Disabled        = -not $computer.Enabled
        LastLogon       = $lastLogon
        DaysSinceLogon  = $daysAgo
        OU              = $computer.DistinguishedName -replace '^CN=[^,]+,', ''
        OperatingSystem = $computer.OperatingSystem
        Location        = $computer.Location
      }
    }
  } catch {
      Debug-Log ": Error querying stale computers: $($_.Exception.Message)" -Type "Error"
      Show-Modal "Query Error" "Failed to query stale computers:`n`n$($_.Exception.Message)"
      return
    }
  }
  }

  ## Sort by days since logon (most stale first)
  $staleAccounts = $staleAccounts | Sort-Object -Property DaysSinceLogon -Descending
  Debug-Log ": Found $($staleAccounts.Count) stale accounts" -Type "Info"

  ## ==================== EXPORT TO CSV ====================
  if ($ExportPath) {
    try {
      $staleAccounts | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
      Debug-Log ": Exported stale accounts to $ExportPath" -Type "Success"
      Show-Modal "Export Complete" "Exported $($staleAccounts.Count) stale accounts to:`n`n$ExportPath"
    } catch {
      Debug-Log ": Export failed: $($_.Exception.Message)" -Type "Error"
      Show-Modal "Export Failed" "Failed to export:`n`n$($_.Exception.Message)"
    }
  }

  ## ==================== SHOW DIALOG ====================
  if ($ShowDialog) {
    $dlg = [Terminal.Gui.Dialog]::new("Stale Account Report", 100, 30)

    ## Header info
    $lblInfo = [Terminal.Gui.Label]::new("Accounts not used in $DaysSinceLogon+ days (Cutoff: $($cutoffDate.ToString('yyyy-MM-dd')))")
    $lblInfo.X = 2
    $lblInfo.Y = 1
    $dlg.Add($lblInfo)

    $lblCount = [Terminal.Gui.Label]::new("Found: $($staleAccounts.Count) stale accounts")
    $lblCount.X = 2
    $lblCount.Y = 2
    $dlg.Add($lblCount)

    ## Results list
    $lstResults = [Terminal.Gui.ListView]::new()
    $lstResults.X = 2
    $lstResults.Y = 4
    $lstResults.Width = [Terminal.Gui.Dim]::Fill(2)
    $lstResults.Height = [Terminal.Gui.Dim]::Fill(5)

    ## Format display strings
    $displayItems = @()
    foreach ($account in $staleAccounts) {
      $statusIcon = if ($account.Enabled) { "✓" } else { "⊗" }
      $lastLogonStr = if ($account.LastLogon -eq [DateTime]::MinValue) {
        "NEVER"
      } else {
        $account.LastLogon.ToString('yyyy-MM-dd')
      }
      $displayItems += "[$statusIcon] $($account.ObjectType) | $($account.Name) | Last: $lastLogonStr ($($account.DaysSinceLogon)d ago)"
    }
    if ($displayItems.Count -eq 0) { $displayItems = @("(No stale accounts found)") }

    $lstResults.SetSource($displayItems)
    $dlg.Add($lstResults)

    ## Action buttons
    $y = [Terminal.Gui.Pos]::Bottom($dlg) - 3

    ## Export button
    $btnExport = [Terminal.Gui.Button]::new(2, $y, "Export CSV")
    $btnExport.add_Clicked({
      $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
      $filename = "stale_accounts_${DaysSinceLogon}days_$timestamp.csv"

      try {
        $staleAccounts | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
        Show-Modal "Export Complete" "Exported to:`n`n$filename"
        Debug-Log ": Exported stale accounts to $filename" -Type "Success"
      } catch {
        Show-Modal "Export Failed" "Failed to export:`n`n$($_.Exception.Message)"
      }
    }.GetNewClosure())
    $dlg.Add($btnExport)

    ## Disable Selected button (if items selected)
    $btnDisable = [Terminal.Gui.Button]::new(20, $y, "Disable Selected")
    $btnDisable.add_Clicked({
      if ($lstResults.SelectedItem -lt 0) {
        Show-Modal "No Selection" "Please select an account"
        return
      }

      $selected = $staleAccounts[$lstResults.SelectedItem]

      ## Find actual object in script arrays
      $obj = $null
      if ($selected.ObjectType -eq 'User') { $obj = $Script:Users | Where-Object { $_.SamAccountName -eq $selected.SamAccountName } | Select-Object -First 1
      } else { $obj = $Script:Computers | Where-Object { $_.SamAccountName -eq $selected.SamAccountName } | Select-Object -First 1 }

      if ($obj) {
        Manage-AccountStatus -Objects @($obj) -Action 'Disable' -Reason "Stale account (not used in $DaysSinceLogon+ days)"

        ## Refresh display
        $selected.Enabled = $false
        $selected.Disabled = $true
        $displayItems[$lstResults.SelectedItem] = "[⊗] $($selected.ObjectType) | $($selected.Name) | Last: $($selected.LastLogon.ToString('yyyy-MM-dd')) ($($selected.DaysSinceLogon)d ago)"
        $lstResults.SetSource($displayItems)
      }
    }.GetNewClosure())
    $dlg.Add($btnDisable)

    ## Close button
    $btnClose = [Terminal.Gui.Button]::new(44, $y, "Close")
    $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
    $dlg.Add($btnClose)
    [Terminal.Gui.Application]::Run($dlg)
  }
  return $staleAccounts
}

## ==================== COPY USER/GROUP (TEMPLATE-BASED) ====================
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
  Debug-Log ": Copy $objectType initiated - Template: $($SourceObject.Name)" -Type "Info"

  ## ==================== INTERACTIVE DIALOG ====================
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
          $last = $Matches[2].ToLower()
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

  ## ==================== DIRECT CREATION ====================
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
      ## ==================== DEMO MODE - COPY USER ====================
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

      Debug-Log ": Created user '$NewName' (demo mode) - Template: $($SourceObject.Name)" -Type "Success"
      if ($CopyMemberships) { Debug-Log ":   Copied $($newUser.Groups.Count) group memberships" -Type "Info" }
      Show-Modal "User Created" "Successfully created user '$NewName'`n`nLogin: $samAccountName`nEmail: $emailAddress$(if ($CopyMemberships) { "`n`nCopied $($newUser.Groups.Count) group memberships" } else { '' })"

      ## ==================== DEMO MODE - COPY GROUP ====================
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
        Debug-Log ": Created group '$NewName' (demo mode) - Template: $($SourceObject.Name)" -Type "Success"
        if ($CopyMemberships) { Debug-Log ":   Copied $($newGroup.Members.Count) members" -Type "Info" }
        Show-Modal "Group Created" "Successfully created group '$NewName'$(if ($CopyMemberships) { "`n`nCopied $($newGroup.Members.Count) members" } else { '' })"
      }

    } else {
      ## ==================== PRODUCTION MODE ====================
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
          Debug-Log ":   Copied $($sourceGroups.Count) group memberships" -Type "Info"
        }
        Debug-Log ": Created user '$NewName' in AD - Template: $($SourceObject.Name)" -Type "Success"
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
          Debug-Log ":   Copied $($sourceMembers.Count) members" -Type "Info"
        }

        Debug-Log ": Created group '$NewName' in AD - Template: $($SourceObject.Name)" -Type "Success"
        Show-Modal "Group Created" "Successfully created group '$NewName' in AD$(if ($CopyMemberships) { "`n`nCopied $($sourceMembers.Count) members" } else { '' })"
      }
    }

    ## Refresh UI
    Refresh-Data -domain $Script:CurrentDomain
    Build-Tree -domain $Script:CurrentDomain
    if ($Script:FilterStatusLabel) { Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel }
  } catch {
    Debug-Log ": Failed to create $objectType '$NewName': $($_.Exception.Message)" -Type "Error"
    Show-Modal "Creation Failed" "Failed to create $objectType '$NewName':`n`n$($_.Exception.Message)"
  }
}

## Fail-safe Demo data
function Load-DefaultDemoData {

  ## Demo forest/domain scaffold
  $Script:ForestName    = "jukebox.example"
  $Script:RootDomain    = "example.com"
  $Script:Domains       = @('example.com','example.net','example.org')
  $Script:CurrentDomain = $Script:RootDomain
  $Script:Sites         = @('Default-First-Site-Name','Branch-Office-Site','Remote-Site')

  ## Demo users (simplified example, add all your full users)
  ## ------------------ Define Demo Users ------------------

  $Script:rawUsers = @(
    ## ========== Eurythmics – UK/Scotland/Aberdeen (Satellite Office) ==========
    @{ Name = 'Annie Lennox'             ; SamAccountName = 'annie.lennox'      ; UserPrincipalName = 'annie.lennox@example.org'       ; OU = @('Locations','UK','Scotland','Aberdeen','Eurythmics')                           ; Groups = @('Eurythmics','Vocalists','Keyboardists')   ; Title = 'Lead Vocalist/Keyboardist'    ; Email = 'annie.lennox@example.org'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Aberdeen Office'   ; Phone = '+44 1224 496 010' ; MobilePhone = '+44 7700 941001' ; Street = '210 Union Street'       ; City = 'Aberdeen'   ; PostalCode = 'AB10 1TL' ; Company = 'Example Music Ltd'     ; Manager = ''                ; Description = 'Lead vocalist and co-founder of Eurythmics'                   },
    @{ Name = 'Dave Stewart'             ; SamAccountName = 'dave.stewart'      ; UserPrincipalName = 'dave.stewart@example.org'       ; OU = @('Locations','UK','Scotland','Aberdeen','Eurythmics')                           ; Groups = @('Eurythmics','Guitarists','Producers')     ; Title = 'Guitarist and Producer'       ; Email = 'dave.stewart@example.org'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Aberdeen Office'   ; Phone = '+44 1224 496 011' ; MobilePhone = '+44 7700 941002' ; Street = '18 Rosemount Place'     ; City = 'Aberdeen'   ; PostalCode = 'AB25 2XP' ; Company = 'Example Music Ltd'     ; Manager = 'Annie Lennox'    ; Description = 'Guitarist, songwriter and producer for Eurythmics'            },

    ## ========== Deacon Blue – UK/Scotland/Dundee ==========
    @{ Name = 'Ricky Ross'               ; SamAccountName = 'ricky.ross'        ; UserPrincipalName = 'ricky.ross@example.net'         ; OU = @('Locations','UK','Scotland','Dundee','Deacon Blue')                            ; Groups = @('Deacon Blue','Sales')                     ; Title = 'Lead Vocalist'                ; Email = 'ricky.ross@example.net'          ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Dundee Office'     ; Phone = '+44 1632 496 001' ; MobilePhone = '+44 7700 392001' ; Street = '1 Tannadice Street'     ; City = 'Dundee'     ; PostalCode = 'DD3 7JW'  ; Company = 'Example Music Ltd'     ; Manager = ''                ; Description = 'Lead vocalist for Deacon Blue'                                },
    @{ Name = 'Lorraine McIntosh'        ; SamAccountName = 'lorraine.mcintosh' ; UserPrincipalName = 'lorraine.mcintosh@example.net'  ; OU = @('Locations','UK','Scotland','Dundee','Deacon Blue')                            ; Groups = @('Deacon Blue','Sales')                     ; Title = 'Vocalist'                     ; Email = 'lorraine.mcintosh@example.net'   ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Dundee Office'     ; Phone = '+44 1632 496 002' ; MobilePhone = '+44 7700 392002' ; Street = '1 Tannadice Street'     ; City = 'Dundee'     ; PostalCode = 'DD3 7JW'  ; Company = 'Example Music Ltd'     ; Manager = 'Ricky Ross'      ; Description = 'Vocalist for Deacon Blue'                                     },
    @{ Name = 'Dougie Vipond'            ; SamAccountName = 'dougie.vipond'     ; UserPrincipalName = 'dougie.vipond@example.net'      ; OU = @('Locations','UK','Scotland','Dundee','Deacon Blue')                            ; Groups = @('Deacon Blue','Sales')                     ; Title = 'Drummer'                      ; Email = 'dougie.vipond@example.net'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Dundee Office'     ; Phone = '+44 1632 496 003' ; MobilePhone = '+44 7700 392003' ; Street = '1 Tannadice Street'     ; City = 'Dundee'     ; PostalCode = 'DD3 7JW'  ; Company = 'Example Music Ltd'     ; Manager = 'Ricky Ross'      ; Description = 'Drummer for Deacon Blue'                                      },
    @{ Name = 'James Prime'              ; SamAccountName = 'james.prime'       ; UserPrincipalName = 'james.prime@example.net'        ; OU = @('Locations','UK','Scotland','Dundee','Deacon Blue')                            ; Groups = @('Deacon Blue','Sales')                     ; Title = 'Keyboardist'                  ; Email = 'james.prime@example.net'         ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Dundee Office'     ; Phone = '+44 1632 496 004' ; MobilePhone = '+44 7700 392004' ; Street = '1 Tannadice Street'     ; City = 'Dundee'     ; PostalCode = 'DD3 7JW'  ; Company = 'Example Music Ltd'     ; Manager = 'Ricky Ross'      ; Description = 'Keyboardist for Deacon Blue'                                  },
    @{ Name = 'Ewen Vernal'              ; SamAccountName = 'ewen.vernal'       ; UserPrincipalName = 'ewen.vernal@example.net'        ; OU = @('Locations','UK','Scotland','Dundee','Deacon Blue')                            ; Groups = @('Deacon Blue','Sales')                     ; Title = 'Bassist'                      ; Email = 'ewen.vernal@example.net'         ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Dundee Office'     ; Phone = '+44 1632 496 005' ; MobilePhone = '+44 7700 392005' ; Street = '1 Tannadice Street'     ; City = 'Dundee'     ; PostalCode = 'DD3 7JW'  ; Company = 'Example Music Ltd'     ; Manager = 'Ricky Ross'      ; Description = 'Bassist for Deacon Blue'                                      },
    @{ Name = 'Graeme Kelling'           ; SamAccountName = 'graeme.kelling'    ; UserPrincipalName = 'graeme.kelling@example.net'     ; OU = @('Locations','UK','Scotland','Dundee','Deacon Blue')                            ; Groups = @('Deacon Blue','Sales')                     ; Title = 'Guitarist'                    ; Email = 'graeme.kelling@example.net'      ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Dundee Office'     ; Phone = '+44 1632 496 006' ; MobilePhone = '+44 7700 392006' ; Street = '1 Tannadice Street'     ; City = 'Dundee'     ; PostalCode = 'DD3 7JW'  ; Company = 'Example Music Ltd'     ; Manager = 'Ricky Ross'      ; Description = 'Guitarist for Deacon Blue'                                    },

    ## ========== Marillion (UK/Scotland/Edinburgh) ==========
    @{ Name = 'Derek Dick'               ; SamAccountName = 'fish'              ; UserPrincipalName = 'fish@example.com'               ; OU = @('Locations','UK','Scotland','Edinburgh','Marillion')                           ; Groups = @('Marillion','Vocalists')                   ; Title = 'Lead Vocalist'                ; Email = 'fish@example.com'                ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $true  ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0221' ; MobilePhone = '+44 7700 222221' ; Street = '22 Tynecastle Street'   ; City = 'Edinburgh'  ; PostalCode = 'EH1 2BB'   ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Former lead vocalist (Fish) for Marillion (1981-1988)'        },
    @{ Name = 'Steve Rothery'            ; SamAccountName = 'steve.rothery'     ; UserPrincipalName = 'steve.rothery@example.com'      ; OU = @('Locations','UK','Scotland','Edinburgh','Marillion')                           ; Groups = @('Marillion','Guitarists')                  ; Title = 'Lead Guitarist'               ; Email = 'steve.rothery@example.com'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0222' ; MobilePhone = '+44 7700 222222' ; Street = '22 Tynecastle Street'   ; City = 'Edinburgh'  ; PostalCode = 'EH1 2BB'   ; Company = 'Example Music Ltd'    ; Manager = 'Derek Dick'      ; Description = 'Lead guitarist and founding member of Marillion'              },
    @{ Name = 'Pete Trewavas'            ; SamAccountName = 'pete.trewavas'     ; UserPrincipalName = 'pete.trewavas@example.com'      ; OU = @('Locations','UK','Scotland','Edinburgh','Marillion')                           ; Groups = @('Marillion','Guitarists')                  ; Title = 'Bassist'                      ; Email = 'pete.trewavas@example.com'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0223' ; MobilePhone = '+44 7700 222223' ; Street = '22 Tynecastle Street'   ; City = 'Edinburgh'  ; PostalCode = 'EH1 2BB'   ; Company = 'Example Music Ltd'    ; Manager = 'Derek Dick'      ; Description = 'Bassist and founding member of Marillion'                     },
    @{ Name = 'Mark Kelly'               ; SamAccountName = 'mark.kelly'        ; UserPrincipalName = 'mark.kelly@example.com'         ; OU = @('Locations','UK','Scotland','Edinburgh','Marillion')                           ; Groups = @('Marillion','Keyboards')                   ; Title = 'Keyboardist'                  ; Email = 'mark.kelly@example.com'          ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0224' ; MobilePhone = '+44 7700 222224' ; Street = '22 Tynecastle Street'   ; City = 'Edinburgh'  ; PostalCode = 'EH1 2BB'   ; Company = 'Example Music Ltd'    ; Manager = 'Derek Dick'      ; Description = 'Keyboardist and founding member of Marillion'                 },
    @{ Name = 'Ian Mosley'               ; SamAccountName = 'ian.mosley'        ; UserPrincipalName = 'ian.mosley@example.com'         ; OU = @('Locations','UK','Scotland','Edinburgh','Marillion')                           ; Groups = @('Marillion','Percussion')                  ; Title = 'Drummer'                      ; Email = 'ian.mosley@example.com'          ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0225' ; MobilePhone = '+44 7700 222225' ; Street = '22 Tynecastle Street'   ; City = 'Edinburgh'  ; PostalCode = 'EH1 2BB'   ; Company = 'Example Music Ltd'    ; Manager = 'Derek Dick'      ; Description = 'Drummer for Marillion (joined 1984)'                          },

    ## ========== Proclaimers (UK/Scotland/Edinburgh) ==========
    @{ Name = 'Craig Reid'               ; SamAccountName = 'craig.reid'        ; UserPrincipalName = 'craig.reid@example.net'         ; OU = @('Locations','Canada','Ontario','Brockville','The Proclaimers')                 ; Groups = @('The Proclaimers','VPN-Users')             ; Title = 'Vocalist / Guitarist'         ; Email = 'craig.reid@example.net'          ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0101' ; MobilePhone = '+44 7700 496011' ; Street = '12 Albion Place'        ; City = 'Edinburgh'  ; PostalCode = 'EH7 5DG'   ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description='Founding member of The Proclaimers'                             },
    @{ Name = 'Charlie Reid'             ; SamAccountName = 'charlie.reid'      ; UserPrincipalName = 'charlie.reid@example.net'       ; OU = @('Locations','USA','Florida','Miami','The Proclaimers')                         ; Groups = @('The Proclaimers','VPN-Users')             ; Title = 'Vocalist / Guitarist'         ; Email = 'charlie.reid@example.net'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0102' ; MobilePhone = '+44 7700 496012' ; Street = '12 Albion Place'        ; City = 'Edinburgh'  ; PostalCode = 'EH7 5DG'   ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description='Founding member of The Proclaimers'                             },

    ## ========== Ultravox – UK/Scotland/Edinburgh (And Vienna) ==========
    @{ Name = 'Midge Ure'                ; SamAccountName = 'midge.ure'         ; UserPrincipalName = 'midge.ure@example.org'          ; OU = @('Locations','UK','Scotland','Edinburgh','Ultravox')                            ; Groups = @('Ultravox','Sales')                        ; Title = 'Vocalist / Guitarist'         ; Email = 'midge.ure@example.org'           ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0226' ; MobilePhone = ''                ; Street = '22 Tynecastle Street'   ; City = 'Edinburgh'  ; PostalCode = 'EH1 2BB'   ; Company = 'Example Music UK'     ; Manager = ''                ; Description = 'Vocalist and guitarist for Ultravox'                          },
    @{ Name = 'Billy Currie'             ; SamAccountName = 'billy.currie'      ; UserPrincipalName = 'billy.currie@example.org'       ; OU = @('Locations','UK','Scotland','Edinburgh','Ultravox')                            ; Groups = @('Ultravox','Sales')                        ; Title = 'Keyboardist / Violinist'      ; Email = 'billy.currie@example.org'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Edinburgh Office'  ; Phone = '+44 131 496 0227' ; MobilePhone = ''                ; Street = '22 Tynecastle Street'   ; City = 'Edinburgh'  ; PostalCode = 'EH1 2BB'   ; Company = 'Example Music UK'     ; Manager = 'Midge Ure'       ; Description = 'Keyboardist and violinist for Ultravox'                       },
    @{ Name = 'Chris Cross'              ; SamAccountName = 'chris.cross'       ; UserPrincipalName = 'chris.cross@example.org'        ; OU = @('Locations','Austria','Vienna','Ultravox')                                     ; Groups = @('Ultravox','Sales')                        ; Title = 'Bassist'                      ; Email = 'chris.cross@example.org'         ; Country = 'AT' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Vienna Office'     ; Phone = '+44 131 496 0228' ; MobilePhone = ''                ; Street = 'Am Hauptbahnhof'        ; City = 'Vienna'     ; PostalCode = '1100 Wien' ; Company = 'Example Music AT'     ; Manager = 'Midge Ure'       ; Description = 'Bassist for Ultravox'                                         },
    @{ Name = 'Warren Cann'              ; SamAccountName = 'warren.cann'       ; UserPrincipalName = 'warren.cann@example.org'        ; OU = @('Locations','Austria','Vienna','Ultravox')                                     ; Groups = @('Ultravox','Sales')                        ; Title = 'Drummer'                      ; Email = 'warren.cann@example.org'         ; Country = 'AT' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Vienna Office'     ; Phone = '+44 131 496 0229' ; MobilePhone = ''                ; Street = 'Am Hauptbahnhof'        ; City = 'Vienna'     ; PostalCode = '1100 Wien' ; Company = 'Example Music AT'     ; Manager = 'Midge Ure'       ; Description = 'Drummer for Ultravox'                                         },

    ## ========== Altered Images – UK/Scotland/Glasgow ==========
    @{ Name = 'Clare Grogan'             ; SamAccountName = 'clare.grogan'      ; UserPrincipalName = 'clare.grogan@example.org'       ; OU = @('Locations','UK','Scotland','Glasgow','Altered Images')                        ; Groups = @('Altered Images','Vocalists')              ; Title = 'Lead Vocalist'                 ; Email = 'clare.grogan@example.org'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0101' ; MobilePhone = '+44 7700 931001' ; Street = '150 Sauchiehall Street' ; City = 'Glasgow'    ; PostalCode = 'G2 3EL'    ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Lead vocalist of Altered Images'                              },
    @{ Name = 'Johnny McElhone'          ; SamAccountName = 'johnny.mcelhone'   ; UserPrincipalName = 'johnny.mcelhone@example.org'    ; OU = @('Locations','UK','Scotland','Glasgow','Altered Images')                        ; Groups = @('Altered Images','Bassists')               ; Title = 'Bassist'                       ; Email = 'johnny.mcelhone@example.org'    ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0102' ; MobilePhone = '+44 7700 931002' ; Street = '42 Hope Street'         ; City = 'Glasgow'    ; PostalCode = 'G2 6AE'    ; Company = 'Example Music Ltd'    ; Manager = 'Clare Grogan'    ; Description = 'Bassist and songwriter for Altered Images'                    },
    @{ Name = 'Jim McKinven'             ; SamAccountName = 'jim.mckinven'      ; UserPrincipalName = 'jim.mckinven@example.org'       ; OU = @('Locations','UK','Scotland','Glasgow','Altered Images')                        ; Groups = @('Altered Images','Guitarists')             ; Title = 'Guitarist'                     ; Email = 'jim.mckinven@example.org'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0103' ; MobilePhone = '+44 7700 931003' ; Street = '77 Bath Street'         ; City = 'Glasgow'    ; PostalCode = 'G2 2EN'    ; Company = 'Example Music Ltd'    ; Manager = 'Clare Grogan'    ; Description = 'Guitarist for Altered Images'                                 },
    @{ Name = 'Michael Anderson'         ; SamAccountName = 'michael.anderson'  ; UserPrincipalName = 'michael.anderson@example.org'   ; OU = @('Locations','UK','Scotland','Glasgow','Altered Images')                        ; Groups = @('Altered Images','Drummers')               ; Title = 'Drummer'                       ; Email = 'michael.anderson@example.org'   ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0104' ; MobilePhone = '+44 7700 931004' ; Street = '305 Argyle Street'      ; City = 'Glasgow'    ; PostalCode = 'G2 8DL'    ; Company = 'Example Music Ltd'    ; Manager = 'Clare Grogan'    ; Description = 'Drummer for Altered Images (aka Tich Anderson)'               },

    ## ========== Simple Minds (UK/Scotland/Glasgow) ==========
    @{ Name = 'Jim Kerr'                 ; SamAccountName = 'jkerr'             ; UserPrincipalName = 'jkerr@example.com'              ; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds')                          ; Groups = @('Simple Minds','Vocalists')                ; Title = 'Lead Vocalist'                 ; Email = 'jim.kerr@example.com'           ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0111' ; MobilePhone = '+44 7700 111111' ; Street = '1 Sauchiehall Street'   ; City = 'Glasgow'    ; PostalCode = 'G1 1AA'    ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Lead vocalist for Simple Minds'                               },
    @{ Name = 'Charlie Burchill'         ; SamAccountName = 'charlie.b'         ; UserPrincipalName = 'charlie.b@example.com'          ; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds')                          ; Groups = @('Simple Minds','Guitarists')               ; Title = 'Lead Guitarist'                ; Email = 'charlie.b@example.com'          ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0112' ; MobilePhone = '+44 7700 111112' ; Street = '1 Sauchiehall Street'   ; City = 'Glasgow'    ; PostalCode = 'G1 1AA'    ; Company = 'Example Music Ltd'    ; Manager = 'Jim Kerr'        ; Description = 'Guitarist and founding member of Simple Minds'                },
    @{ Name = 'Mel Gaynor'               ; SamAccountName = 'mel.gaynor'        ; UserPrincipalName = 'mel.gaynor@example.com'         ; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds')                          ; Groups = @('Simple Minds','Percussion')               ; Title = 'Drummer'                       ; Email = 'mel.gaynor@example.com'         ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0113' ; MobilePhone = '+44 7700 111113' ; Street = '1 Sauchiehall Street'   ; City = 'Glasgow'    ; PostalCode = 'G1 1AA'    ; Company = 'Example Music Ltd'    ; Manager = 'Jim Kerr'        ; Description = 'Drummer for Simple Minds'                                     },
    @{ Name = 'Mick MacNeil'             ; SamAccountName = 'mick.macneil'      ; UserPrincipalName = 'mick.macneil@example.com'       ; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds')                          ; Groups = @('Simple Minds','Musicians','Former Staff') ; Title = 'Keyboardist (Former)'          ; Email = 'mick.macneil@example.com'       ; Country = 'UK' ; Disabled = $true  ; Locked = $true  ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0120' ; MobilePhone = '+44 7700 111120' ; Street = '1 Sauchiehall Street'   ; City = 'Glasgow'    ; PostalCode = 'G1 1AA'    ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Former keyboardist for Simple Minds (1977-1990)'              },
    @{ Name = 'Derek Forbes'             ; SamAccountName = 'derek.forbes'      ; UserPrincipalName = 'derek.forbes@example.com'       ; OU = @('Locations','UK','Scotland','Glasgow','Simple Minds')                          ; Groups = @('Simple Minds','Musicians','Former Staff') ; Title = 'Bassist (Former)'              ; Email = 'derek.forbes@example.com'       ; Country = 'UK' ; Disabled = $true  ; Locked = $true  ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Glasgow Office'    ; Phone = '+44 141 496 0121' ; MobilePhone = '+44 7700 111121' ; Street = '1 Sauchiehall Street'   ; City = 'Glasgow'    ; PostalCode = 'G1 1AA'    ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Former bassist for Simple Minds (1977-1985)'                  },

    ## ========== Wet Wet Wet — UK/Scotland/Clydebank ==========
    @{ Name = 'Marti Pellow'             ; SamAccountName = 'marti.pellow'      ; UserPrincipalName = 'marti.pellow@example.net'       ; OU = @('Locations','UK','Scotland','Clydebank','Wet Wet Wet')                         ; Groups = @('Wet Wet Wet','Sales','VPN Users')         ; Title = 'Lead Vocalist'                 ; Email = 'marti.pellow@example.net'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Clydebank Office'  ; Phone = '+44 141 496 0201' ; MobilePhone = '+44 7700 496201' ; Street = 'Dumbarton Road 200'     ; City = 'Clydebank'  ; PostalCode = 'G81 1UE'   ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Lead vocalist of Wet Wet Wet'                                 },
    @{ Name = 'Graeme Clark'             ; SamAccountName = 'graeme.clark'      ; UserPrincipalName = 'graeme.clark@example.net'       ; OU = @('Locations','UK','Scotland','Clydebank','Wet Wet Wet')                         ; Groups = @('Wet Wet Wet','Sales','VPN Users')         ; Title = 'Bassist'                       ; Email = 'graeme.clark@example.net'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Clydebank Office'  ; Phone = '+44 141 496 0202' ; MobilePhone = '+44 7700 496202' ; Street = 'Dumbarton Road 201'     ; City = 'Clydebank'  ; PostalCode = 'G81 1UE'   ; Company = 'Example Music Ltd'    ; Manager = 'Marti Pellow'    ; Description = 'Bassist for Wet Wet Wet'                                      },
    @{ Name = 'Tommy Cunningham'         ; SamAccountName = 'tommy.cunningham'  ; UserPrincipalName = 'tommy.cunningham@example.net'   ; OU = @('Locations','UK','Scotland','Clydebank','Wet Wet Wet')                         ; Groups = @('Wet Wet Wet','Sales','VPN Users')         ; Title = 'Drummer'                       ; Email = 'tommy.cunningham@example.net'   ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Clydebank Office'  ; Phone = '+44 141 496 0203' ; MobilePhone = '+44 7700 496203' ; Street = 'Dumbarton Road 202'     ; City = 'Clydebank'  ; PostalCode = 'G81 1UE'   ; Company = 'Example Music Ltd'    ; Manager = 'Marti Pellow'    ; Description = 'Drummer for Wet Wet Wet'                                      },
    @{ Name = 'Neil Mitchell'            ; SamAccountName = 'neil.mitchell'     ; UserPrincipalName = 'neil.mitchell@example.net'      ; OU = @('Locations','UK','Scotland','Clydebank','Wet Wet Wet')                         ; Groups = @('Wet Wet Wet','Sales','VPN Users')         ; Title = 'Keyboardist'                   ; Email = 'neil.mitchell@example.net'      ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Clydebank Office'  ; Phone = '+44 141 496 0204' ; MobilePhone = '+44 7700 496204' ; Street = 'Dumbarton Road 203'     ; City = 'Clydebank'  ; PostalCode = 'G81 1UE'   ; Company = 'Example Music Ltd'    ; Manager = 'Marti Pellow'    ; Description = 'Keyboardist for Wet Wet Wet'                                  },

    ## ==========The Police - UK/England/Newcastle ==========
    @{ Name = 'Gordon Summer'            ; SamAccountName = 'sting'             ; UserPrincipalName = 'sting@example.org'              ; OU = @('Locations','UK','England','Newcastle','The Police')                           ; Groups = @('The Police','Vocalists','Bassists')        ; Title = 'Lead Vocalist & Bassist'      ; Email = 'sting@example.org'              ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Newcastle Office'  ; Phone = '+44 191 496 0001' ; MobilePhone = '+44 7700 910001' ; Street = '14 Grey Street'         ; City = 'Newcastle'  ; PostalCode = 'NE1 6BH'   ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Lead vocalist and bassist of The Police'                      },
    @{ Name = 'Andy Summers'             ; SamAccountName = 'andy.summers'      ; UserPrincipalName = 'andy.summers@example.org'       ; OU = @('Locations','UK','England','Newcastle','The Police')                           ; Groups = @('The Police','Guitarists')                  ; Title = 'Guitarist'                    ; Email = 'andy.summers@example.org'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Newcastle Office'  ; Phone = '+44 191 496 0002' ; MobilePhone = '+44 7700 910002' ; Street = '11 Dean Street'         ; City = 'Newcastle'  ; PostalCode = 'NE1 1PG'   ; Company = 'Example Music Ltd'    ; Manager = 'Sting'           ; Description = 'Guitarist for The Police'                                     },
    @{ Name = 'Stewart Copeland'         ; SamAccountName = 'stewart.copeland'  ; UserPrincipalName = 'stewart.copeland@example.org'   ; OU = @('Locations','UK','England','Newcastle','The Police')                           ; Groups = @('The Police','Drummers')                    ; Title = 'Drummer'                      ; Email = 'stewart.copeland@example.org'   ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Newcastle Office'  ; Phone = '+44 191 496 0003' ; MobilePhone = '+44 7700 910003' ; Street = '5 Collingwood Street'   ; City = 'Newcastle'  ; PostalCode = 'NE1 1JF'   ; Company = 'Example Music Ltd'    ; Manager = 'Sting'           ; Description = 'Drummer for The Police'                                       },

    ## ========== New Order – UK/England/Manchester ==========
    @{ Name = 'Bernard Sumner'           ; SamAccountName = 'bernard.sumner'    ; UserPrincipalName = 'bernard.sumner@example.org'     ; OU = @('Locations','UK','England','Manchester','New Order')                           ; Groups = @('New Order','Vocalists','Guitarists')       ; Title = 'Lead Vocalist & Guitarist'    ; Email = 'bernard.sumner@example.org'     ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Manchester Office' ; Phone = '+44 191 496 0161' ; MobilePhone = '+44 7700 951161' ; Street = 'Annalade Road 10'       ; City = 'Manchester' ; PostalCode = 'M16 9AB'   ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Lead vocalist and guitarist for New Order'                    },
    @{ Name = 'Stephen Morris'           ; SamAccountName = 'stephen.morris'    ; UserPrincipalName = 'stephen.morris@example.org'     ; OU = @('Locations','UK','England','Manchester','New Order')                           ; Groups = @('New Order','Drummers')                     ; Title = 'Drummer'                      ; Email = 'stephen.morris@example.org'     ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Manchester Office' ; Phone = '+44 191 496 0162' ; MobilePhone = '+44 7700 951162' ; Street = 'Cromwell Road 22'       ; City = 'Manchester' ; PostalCode = 'M16 9AB'   ; Company = 'Example Music Ltd'    ; Manager = 'Bernard Sumner'  ; Description = 'Drummer for New Order'                                        },
    @{ Name = 'Peter Hook'               ; SamAccountName = 'peter.hook'        ; UserPrincipalName = 'peter.hook@example.org'         ; OU = @('Locations','UK','England','Manchester','New Order')                           ; Groups = @('New Order','Bassists')                     ; Title = 'Bassist'                      ; Email = 'peter.hook@example.org'         ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Manchester Office' ; Phone = '+44 191 496 0163' ; MobilePhone = '+44 7700 951163' ; Street = 'Blake Road 5'           ; City = 'Manchester' ; PostalCode = 'M16 9AB'   ; Company = 'Example Music Ltd'    ; Manager = 'Bernard Sumner'  ; Description = 'Bassist and co-founder of New Order'                          },
    @{ Name = 'Gillian Gilbert'          ; SamAccountName = 'gillian.gilbert'   ; UserPrincipalName = 'gillian.gilbert@example.org'    ; OU = @('Locations','UK','England','Manchester','New Order')                           ; Groups = @('New Order','Keyboardists','Guitarists')    ; Title = 'Keyboardist & Guitarist'      ; Email = 'gillian.gilbert@example.org'    ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Manchester Office' ; Phone = '+44 191 496 0164' ; MobilePhone = '+44 7700 951164' ; Street = 'Chorlton Road 18'       ; City = 'Manchester' ; PostalCode = 'M16 9AB'   ; Company = 'Example Music Ltd'    ; Manager = 'Bernard Sumner'  ; Description = 'Keyboardist and guitarist for New Order'                      },

    ## ========== Echo And The Bunnymen - UK/England/Liverpool ==========
    @{ Name = 'Ian McCulloch'            ; SamAccountName = 'ian.mcculloch'     ; UserPrincipalName = 'ian.mcculloch@example.org'      ; OU = @('Locations','UK','England','Liverpool','Echo and The Bunnymen')                ; Groups = @('Echo and The Bunnymen','Vocalists')        ; Title = 'Lead Vocalist'                ; Email = 'ian.mcculloch@example.org'      ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $true  ; Department = 'Music' ; Office = 'Liverpool Office'  ; Phone = '+44 151 496 0001' ; MobilePhone = '+44 7700 900001' ; Street = '12 Mathew Street'       ; City = 'Liverpool'  ; PostalCode = 'L1 4ED'    ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Lead vocalist of Echo and The Bunnymen'                       },
    @{ Name = 'Will Sergeant'            ; SamAccountName = 'will.sergeant'     ; UserPrincipalName = 'will.sergeant@example.org'      ; OU = @('Locations','UK','England','Liverpool','Echo and The Bunnymen')                ; Groups = @('Echo and The Bunnymen','Guitarists')       ; Title = 'Guitarist'                    ; Email = 'will.sergeant@example.org'      ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Liverpool Office'  ; Phone = '+44 151 496 0002' ; MobilePhone = '+44 7700 900002' ; Street = '22 Bold Street'         ; City = 'Liverpool'  ; PostalCode = 'L1 4HR'    ; Company = 'Example Music Ltd'    ; Manager = 'Ian McCulloch'   ; Description = 'Guitarist for Echo and The Bunnymen'                          },
    @{ Name = 'Les Pattinson'            ; SamAccountName = 'les.pattinson'     ; UserPrincipalName = 'les.pattinson@example.org'      ; OU = @('Locations','UK','England','Liverpool','Echo and The Bunnymen')                ; Groups = @('Echo and The Bunnymen','Bassists')         ; Title = 'Bass Guitarist'               ; Email = 'les.pattinson@example.org'      ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Liverpool Office'  ; Phone = '+44 151 496 6003' ; MobilePhone = '+44 7700 900003' ; Street = '8 Seel Street'          ; City = 'Liverpool'  ; PostalCode = 'L1 4BE'    ; Company = 'Example Music Ltd'    ; Manager = 'Ian McCulloch'   ; Description = 'Bass guitarist for Echo and The Bunnymen'                     },
    @{ Name = 'Pete de Freitas'          ; SamAccountName = 'pete.defreitas'    ; UserPrincipalName = 'pete.defreitas@example.org'     ; OU = @('Locations','UK','England','Liverpool','Echo and The Bunnymen')                ; Groups = @('Echo and The Bunnymen','Drummers')         ; Title = 'Drummer'                      ; Email = 'pete.defreitas@example.org'     ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Liverpool Office'  ; Phone = '+44 151 496 0004' ; MobilePhone = '+44 7700 900004' ; Street = '5 Dale Street'          ; City = 'Liverpool'  ; PostalCode = 'L2 2EH'    ; Company = 'Example Music Ltd'    ; Manager = 'Ian McCulloch'   ; Description = 'Drummer for Echo and The Bunnymen'                            },

    ## ========== UB40 — UK/England/Birmingham ==========
    @{ Name = 'Ali Campbell'             ; SamAccountName = 'ali.campbell'      ; UserPrincipalName = 'ali.campbell@example.net'       ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Lead Vocalist'                ; Email = 'ali.campbell@example.net'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone = '+44 121 496 0101' ; MobilePhone = '+44 7700 921601' ; Street = '40 Broad Street'        ; City = 'Birmingham' ; PostalCode = 'B1 2EU'    ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Lead voca l ist of UB40'                                      },
    @{ Name = 'Robin Campbell'           ; SamAccountName = 'robin.campbell'    ; UserPrincipalName = 'robin.campbell@example.net'     ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Guitarist'                    ; Email = 'robin.campbell@example.net'     ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone =' +44 121 496 0102' ; MobilePhone = '+44 7700 921602' ; Street = '41 Broad Street'        ; City = 'Birmingham' ; PostalCode = 'B1 2EU'    ; Company = 'Example Music Ltd'    ; Manager = 'Ali Campbell'    ; Description = 'Guitarist for UB40'                                           },
    @{ Name = 'Brian Travers'            ; SamAccountName = 'brian.travers'     ; UserPrincipalName = 'brian.travers@example.net'      ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Saxophonist'                  ; Email = 'brian.travers@example.net'      ; Country = 'UK' ; Disabled = $true  ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone =' +44 121 496 0103' ; MobilePhone = '+44 7700 921603' ; Street = '42 Broad Street'        ; City = 'Birmingham' ; PostalCode = 'B1 2EU'    ; Company = 'Example Music Ltd'    ; Manager = 'Ali Campbell'    ; Description = 'Saxophonist for UB40 (account disabled)'                      },
    @{ Name = 'Earl Falconer'            ; SamAccountName = 'earl.falconer'     ; UserPrincipalName = 'earl.falconer@example.net'      ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Bassist'                      ; Email = 'earl.falconer@example.net'      ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone =' +44 121 496 0104' ; MobilePhone = '+44 7700 921604' ; Street = '43 Broad Street'        ; City = 'Birmingham' ; PostalCode = 'B1 2EU'    ; Company = 'Example Music Ltd'    ; Manager = 'Ali Campbell'    ; Description = 'Bassist for UB40'                                             },
    @{ Name = 'Norman Hassan'            ; SamAccountName = 'norman.hassan'     ; UserPrincipalName = 'norman.hassan@example.net'      ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Percussionist'                ; Email = 'norman.hassan@example.net'      ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone =' +44 121 496 0105' ; MobilePhone = '+44 7700 921605' ; Street = '44 Broad Street'        ; City = 'Birmingham' ; PostalCode = 'B1 2EU'    ; Company = 'Example Music Ltd'    ; Manager = 'Ali Campbell'    ; Description = 'Percussionist for UB40'                                       },
    @{ Name = 'Terence Wilson'           ; SamAccountName = 'astro.wilson'      ; UserPrincipalName = 'astro@example.net'              ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Toaster / Trumpet'            ; Email = 'astro.wilson@example.net'       ; Country = 'UK' ; Disabled = $true  ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone =' +44 121 496 0106' ; MobilePhone = '+44 7700 921606' ; Street = '45 Broad Street'        ; City = 'Birmingham' ; PostalCode = 'B1 2EU'    ; Company = 'Example Music Ltd'    ; Manager = 'Ali Campbell'    ; Description = 'Toaster and trumpeter for UB40 (account disabled)'            },
    @{ Name = 'James Brown'              ; SamAccountName = 'james.brown'       ; UserPrincipalName = 'james.brown@example.net'        ; OU = @('Locations','UK','England','Birmingham','UB40')                                ; Groups = @('UB40','Sales','VPN Users')                 ; Title = 'Drummer'                      ; Email = 'james.brown@example.net'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office='Birmingham Office'   ; Phone =' +44 121 496 0107' ; MobilePhone = '+44 7700 921607' ; Street = '46 Broad Street'        ; City = 'Birmingham' ; PostalCode = 'B1 2EU'    ; Company = 'Example Music Ltd'    ; Manager = 'Ali Campbell'    ; Description = 'Drummer for UB40'                                             },

    ## ========== Erasure (UK/England/London) ==========
    @{ Name = 'Andy Bell'                ; SamAccountName = 'andy.bell'         ; UserPrincipalName = 'andy.bell@example.com'          ; OU = @('Locations','UK','England','London','Erasure')                                 ; Groups = @('Erasure','Vocalists')                       ; Title = 'Lead Vocalist'               ; Email = 'andy.bell@example.com'          ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'London Office'     ; Phone = '+44 207 496 0011' ; MobilePhone = '+44 7700 333333' ; Street = '15 Carnaby Street'      ; City = 'London'     ; PostalCode = 'E1 7AA'    ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Lead vocalist for Erasure'                                    },
    @{ Name = 'Vince Clarke'             ; SamAccountName = 'vince.clarke'      ; UserPrincipalName = 'vince.clarke@example.com'       ; OU = @('Locations','UK','England','London','Erasure')                                 ; Groups = @('Erasure','Synth','Keyboards')               ; Title = 'Synth / Keyboardist'         ; Email = 'vince.clarke@example.com'       ; Country = 'UK' ; Disabled = $false ; Locked = $true  ; MustChangePassword = $false ; Department = 'Music' ; Office = 'London Office'     ; Phone = '+44 207 496 0012' ; MobilePhone = '+44 7700 333334' ; Street = '15 Carnaby Street'      ; City = 'London'     ; PostalCode = 'E1 7AA'    ; Company = 'Example Music Ltd'    ; Manager = 'Andy Bell'       ; Description = 'Synthesizer pioneer - member of Depeche Mode & Erasure'       },

    ## ========== Mel And Kim – UK/England/London ==========
    @{ Name = 'Melanie Appleby'          ; SamAccountName = 'melanie.appleby'   ; UserPrincipalName = 'mel.appleby@example.com'        ; OU = @('Locations','UK','England','London','Sales')                                   ; Groups = @('Mel And Kim','Sales')                       ; Title = 'Singer and Dancer'           ; Email = 'mel.appleby@example.com'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Sales' ; Office = 'London Office'     ; Phone = '+44 207 496 0013' ; MobilePhone = '+44 7700 333335' ; Street = '15 Carnaby Street'      ; City ='London'      ; PostalCode = 'E1 7AA'    ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Member of Mel and Kim, London sales team'                     },
    @{ Name = 'Kimberly Appleby'         ; SamAccountName = 'kim.appleby'       ; UserPrincipalName = 'kim.appleby@example.com'        ; OU = @('Locations','UK','England','London','Sales')                                   ; Groups = @('Mel And Kim','Sales')                       ; Title = 'Singer and Dancer'           ; Email = 'kim.appleby@example.com'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Sales' ; Office = 'London Office'     ; Phone = '+44 207 496 0014' ; MobilePhone = '+44 7700 333336' ; Street = '15 Carnaby Street'      ; City ='London'      ; PostalCode = 'E1 7AA'    ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Member of Mel and Kim, London sales team'                     },

    ## ========== Depeche Mode (UK/England/London) ==========
    @{ Name = 'Dave Gahan'               ; SamAccountName = 'dave.gahan'        ; UserPrincipalName = 'dave.gahan@example.com'         ; OU = @('Locations','UK','England','London','Depeche Mode')                            ; Groups = @('Depeche Mode','Vocalists')                  ; Title = 'Lead Vocalist'               ; Email = 'dave.gahan@example.com'         ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $true  ; Department = 'Music' ; Office = 'London Office'     ; Phone = '+44 207 496 0021' ; MobilePhone = '+44 7700 444442' ; Street = '32 Abbey Lane'          ; City = 'London'     ; PostalCode = 'EC2 1AA'   ; Company = 'Example Music Ltd'    ; Manager = 'Martin Gore'     ; Description = 'Lead vocalist for Depeche Mode'                               },
    @{ Name = 'Martin Gore'              ; SamAccountName = 'martin.gore'       ; UserPrincipalName = 'martin.gore@example.com'        ; OU = @('Locations','UK','England','London','Depeche Mode')                            ; Groups = @('Depeche Mode','Guitarists','Keyboards')     ; Title = 'Guitarist/Keyboardist'       ; Email = 'martin.gore@example.com'        ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'London Office'     ; Phone = '+44 207 496 0022' ; MobilePhone = '+44 7700 444441' ; Street = '32 Abbey Lane'          ; City = 'London'     ; PostalCode = 'EC2 1AA'   ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Guitarist, keyboardist & primary songwriter for Depeche Mode' },
    @{ Name = 'Alan Wilder'              ; SamAccountName = 'alan.wilder'       ; UserPrincipalName = 'alan.wilder@example.com'        ; OU = @('Locations','UK','England','London','Depeche Mode')                            ; Groups = @('Depeche Mode','Keyboards','Percussion')     ; Title = 'Keyboardist/Drummer'         ; Email = 'alan.wilder@example.com'        ; Country = 'UK' ; Disabled = $false ; Locked = $true  ; MustChangePassword = $false ; Department = 'Music' ; Office = 'London Office'     ; Phone = '+44 207 496 0023' ; MobilePhone = '+44 7700 444444' ; Street = '32 Abbey Lane'          ; City = 'London'     ; PostalCode = 'EC2 1AA'   ; Company = 'Example Music Ltd'    ; Manager = 'Martin Gore'     ; Description = 'Multi-instrumentalist for Depeche Mode (1982-1995, departed)' },
    @{ Name = 'Andrew Fletcher'          ; SamAccountName = 'andrew.fletcher'   ; UserPrincipalName = 'andrew.fletcher@example.com'    ; OU = @('Locations','UK','England','London','Depeche Mode')                            ; Groups = @('Depeche Mode','Keyboards')                  ; Title = 'Keyboards/Bass Synth'        ; Email = 'andrew.fletcher@example.com'    ; Country = 'UK' ; Disabled = $true  ; Locked = $true  ; MustChangePassword = $false ; Department = 'Music' ; Office = 'London Office'     ; Phone = '+44 207 496 0024' ; MobilePhone = '+44 7700 444443' ; Street = '32 Abbey Lane'          ; City = 'London'     ; PostalCode = 'EC2 1AA'   ; Company = 'Example Music Ltd'    ; Manager = 'Martin Gore'     ; Description = 'Keyboard and bass synthesizer for Depeche Mode (deceased)'    },

    ## ========== TV-2 (Danmark/Københaven) ==========
    @{ Name = 'Steffen Brandt'           ; SamAccountName = 'steffen.brandt'    ; UserPrincipalName = 'steffen.brandt@example.com'     ; OU = @('Locations','Danmark','Sjælland', 'Copenhagen','TV-2')                         ; Groups = @('TV-2','Vocalists','Guitarists')             ; Title = 'Lead Vocalist / Guitarist'   ; Email = 'steffen.brandt@example.com'     ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Copenhagen Office' ; Phone = '+45 0000 2222'    ; MobilePhone = '+45 50 12 3457'  ; Street = '1 Raadhuspladsen'       ; City = 'Copenhagen' ; PostalCode = '1550'      ; Company = 'Example Music ApS'    ; Manager = ''                ; Description = 'Frontman of TV-2'                                             },
    @{ Name = 'Hans Erik Lerchenfeldt'   ; SamAccountName = 'hans.lerchenfeldt' ; UserPrincipalName = 'hans.lerchenfeldt@example.com'  ; OU = @('Locations','Danmark','Sjælland', 'Copenhagen','TV-2')                         ; Groups = @('TV-2','Musicians')                          ; Title = 'Bassist'                     ; Email = 'hans.lerchenfeldt@example.com'  ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Copenhagen Office' ; Phone = '+45 3312 3457'    ; MobilePhone = '+45 20 11 1157'  ; Street = 'Nørrebrogade 2'         ; City = 'Copenhagen' ; PostalCode = '2200'      ; Company = 'Example Music ApS'    ; Manager = 'Steffen Brandt'  ; Description = 'Bassist for TV-2'                                             },
    @{ Name = 'Sven Gaul'                ; SamAccountName = 'sven.gaul'         ; UserPrincipalName = 'sven.gaul@example.com'          ; OU = @('Locations','Danmark','Sjælland', 'Copenhagen','TV-2')                         ; Groups = @('TV-2','Musicians')                          ; Title = 'Drummer'                     ; Email = 'sven.gaul@example.com'          ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Copenhagen Office' ; Phone = '+45 3312 3458'    ; MobilePhone = '+45 20 11 1158'  ; Street = 'Nørrebrogade 3'         ; City = 'Copenhagen' ; PostalCode = '2200'      ; Company = 'Example Music ApS'    ; Manager = 'Steffen Brandt'  ; Description = 'Drummer for TV-2'                                             },
    @{ Name = 'Georg Olesen'             ; SamAccountName = 'georg.olesen'      ; UserPrincipalName = 'georg.olesen@example.com'       ; OU = @('Locations','Danmark','Nord Jyland', 'Aarhus','TV-2')                          ; Groups = @('TV-2','Musicians','Former Staff')           ; Title = 'Guitarist (Former)'          ; Email = 'georg.olesen@example.com'       ; Country = 'DK' ; Disabled = $true  ; Locked = $true  ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Aarhus Office'     ; Phone = '+45 8612 3470'    ; MobilePhone = '+45 20 11 1170'  ; Street = 'Åboulevarden 20'        ; City = 'Aarhus'     ; PostalCode = '8000'      ; Company = 'Example Music ApS'    ; Manager = ''                ; Description = 'Former guitarist and co-founder of TV-2 (1981-2003)'          },

    ## ========== Rocazino (Danmark/Køge) ==========
    @{ Name = 'Ulla Kjaer'               ; SamAccountName = 'ulla.kjaer'        ; UserPrincipalName = 'ulla.kjaer@example.com'         ; OU = @('Locations','Danmark','Koge','Rocazino')                                       ; Groups = @('Rocazino','Vocalists')                      ; Title = 'Lead Vocalist'               ; Email = 'ulla.kjaer@example.com'         ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Koge Office'       ; Phone = '+45 0000 2234'    ; MobilePhone = '+45 3012 3456'   ; Street = '7 Torvet'               ; City = 'Koge'       ; PostalCode = '4600'      ; Company = 'Example Music ApS'    ; Manager = ''                ; Description = 'Lead vocalist for Rocazino'                                   },
    @{ Name = 'Michael Bruun'            ; SamAccountName = 'michael.bruun'     ; UserPrincipalName = 'michael.bruun@example.com'      ; OU = @('Locations','Danmark','Koge','Rocazino')                                       ; Groups = @('Rocazino','Guitarists')                     ; Title = 'Guitarist'                   ; Email = 'michael.bruun@example.com'      ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Koge Office'       ; Phone = '+45 0000 2235'    ; MobilePhone = '+45 3012 3457'   ; Street = '7 Torvet'               ; City = 'Koge'       ; PostalCode = '4600'      ; Company = 'Example Music ApS'    ; Manager = 'Ulla Kjaer'      ; Description = 'Guitarist and songwriter for Rocazino'                        },
    @{ Name = 'Jan Sivertsen'            ; SamAccountName = 'jan.sivertsen'     ; UserPrincipalName = 'jan.sivertsen@example.com'      ; OU = @('Locations','Danmark','Koge','Rocazino')                                       ; Groups = @('Rocazino','Percussion')                     ; Title = 'Drummer'                     ; Email = 'jan.sivertsen@example.com'      ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Koge Office'       ; Phone = '+45 0000 2236'    ; MobilePhone = '+45 3012 3458'   ; Street = '7 Torvet'               ; City = 'Koge'       ; PostalCode = '4600'      ; Company = 'Example Music ApS'    ; Manager = 'Ulla Kjaer'      ; Description = 'Drummer for Rocazino'                                         },

    ## ========== Whigfield (Danmark/Kørsor og Faxe) ==========
    @{ Name = 'Sannie Charlotte Carlson' ; SamAccountName = 'whigfield'         ; UserPrincipalName = 'whigfield@example.com'          ; OU = @('Locations','Danmark','Sjælland','Køge','Kørsor','Whigfield','Users')          ; Groups = @('Whigfield','Warehouse Access')              ; Title  = 'Performer'                  ; Email  = 'whigfield@example.com'         ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Kørsor Office'     ; Phone = '+45 0000 3235'    ; MobilePhone = '+45 4012 3457'   ; Street = 'Torvegade 35.'          ; City = 'Faxe'       ; PostalCode = '4640'      ; Company = 'Example Music ApS'    ; Manager = ''                ; Description = 'Danish singer-songwriter'                                     },

    ## ========== MØ (Danmark/Odense) ==========
    @{ Name = 'Karen Marie Orsted'       ; SamAccountName = 'karen.orsted'      ; UserPrincipalName = 'mo@example.com'                 ; OU = @('Locations','Danmark','Fyn', 'Odense','Mo')                                    ; Groups = @('Mo','Vocalists')                            ; Title = 'Singer / Songwriter'         ; Email = 'mo@example.com'                 ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Odense Office'     ; Phone = '+45 0000 3234'    ; MobilePhone = '+45 4012 3456'   ; Street = '22 Vestergade'          ; City = 'Odense'     ; PostalCode = '5000'      ; Company = 'Example Music ApS'    ; Manager = ''                ; Description = 'Danish singer-songwriter known internationally as MØ'         },

    ## ========== Lis Sørensen (Danmark/Bramming) ==========
    @{ Name = 'Lis Sørensen'             ; SamAccountName = 'lissorensen'       ; UserPrincipalName = 'lis.sorensen@example.com'       ; OU = @('Locations','Danmark','Syd Jyland','Bramming','Music','Lis Sørensen','Users')  ; Groups  = @('Lis Sørensen','Studio Access')             ; Title = 'Singer'                      ; Email = 'lis.sorensen@example.com'       ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Bramming Studio'   ; Phone = '+45 0000 4411'    ; MobilePhon  = '+45 4099 8822'   ; Street = 'Nørregade 12'           ; City = 'Bramming'   ; PostalCode = '1165'      ; Company = 'Example Music ApS'    ; Manager = ''                ; Description = 'Danish singer; known for the song Brændt'                     },

    ## ========== Helena Christensen (Danmark/Esjberg) ==========
    @{ Name = 'Helena Christensen'       ; SamAccountName = 'helenachristensen' ; UserPrincipalName = 'helena.christensen@example.com' ; OU = @('Locations','Danmark','Syd Jyland','Esbjerg','Creative','Photography','Users') ; Groups = @('Creative Team', 'Esbjerg Office Access')    ; Title = 'Company Photographer'        ; Email = 'helena.christensen@example.com' ; Country = 'DK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ;  Department = 'Art'  ; Office = 'Esjberg Studio'    ; Phone = '+45 00 007 788'   ; MobilePhone = '+45 4011 2233'   ; Street = 'Skolegade 18'           ; City = 'Esjberg'    ; PostalCode = '6700'      ; Company = 'Example Creative ApS' ; Manager = ''                ; Description = 'Model/Photographer working on Chris Issak Wicked Game video'  },

    ## ========== Kraftwerk (West Germany/Bonn) ==========
    @{ Name = 'Ralf Hutter'              ; SamAccountName = 'ralf.hutter'       ; UserPrincipalName = 'ralf.hutter@example.net'        ; OU = @('Locations','Germany','Bonn','Kraftwerk')                                      ; Groups = @('Kraftwerk','Vocalists','Musicians')         ; Title = 'Vocals/Synthesizer'          ; Email = 'ralf.hutter@example.net'        ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Bonn Office'       ; Phone = '+49 22 8111 111'  ; MobilePhone = '+49 1701 1111'   ; Street = 'Adenauerallee 1'        ; City = 'Bonn'       ; PostalCode = '53113'     ; Company = 'Example Music GmbH'   ; Manager = ''                ; Description = 'Co-founder and frontman of Kraftwerk'                         },
    @{ Name = 'Florian Schneider'        ; SamAccountName = 'florian.schneider' ; UserPrincipalName = 'florian.schneider@example.net'  ; OU = @('Locations','Germany','Bonn','Kraftwerk')                                      ; Groups = @('Kraftwerk','Musicians')                     ; Title = 'Synthesizer/Flute'           ; Email = 'florian.schneider@example.net'  ; Country = 'DE' ; Disabled = $true  ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Bonn Office'       ; Phone = '+49 22 8111 112'  ; MobilePhone = '+49 1701 1112'   ; Street = 'Adenauerallee 2'        ; City = 'Bonn'       ; PostalCode = '53113'     ; Company = 'Example Music GmbH'   ; Manager = 'Ralf Hutter'     ; Description = 'Co-founder of Kraftwerk (1947-2020) - Account disabled'       },
    @{ Name = 'Wolfgang Flur'            ; SamAccountName = 'wolfgang.flur'     ; UserPrincipalName = 'wolfgang.flur@example.net'      ; OU = @('Locations','Germany','Bonn','Kraftwerk')                                      ; Groups = @('Kraftwerk','Musicians','Percussionists')    ; Title = 'Electronic Drums'            ; Email = 'wolfgang.flur@example.net'      ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Bonn Office'       ; Phone = '+49 22 8111 113'  ; MobilePhone = '+49 1701 1113'   ; Street = 'Adenauerallee 3'        ; City = 'Bonn'       ; PostalCode = '53113'     ; Company = 'Example Music GmbH'   ; Manager = 'Ralf Hutter'     ; Description = 'Electronic percussionist for Kraftwerk'                       },
    @{ Name = 'Karl Bartos'              ; SamAccountName = 'karl.bartos'       ; UserPrincipalName = 'karl.bartos@example.net'        ; OU = @('Locations','Germany','Bonn','Kraftwerk')                                      ; Groups = @('Kraftwerk','Musicians','Percussionists')    ; Title = 'Electronic Percussion'       ; Email = 'karl.bartos@example.net'        ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Bonn Office'       ; Phone = '+49 22 8111 114'  ; MobilePhone = '+49 1701 1114'   ; Street = 'Adenauerallee 4'        ; City = 'Bonn'       ; PostalCode = '53113'     ; Company = 'Example Music GmbH'   ; Manager = 'Ralf Hutter'     ; Description = 'Electronic percussionist and composer for Kraftwerk'          },

    ## ========== Fancy - Germany/Munich ==========
    @{ Name = 'Manfred Segieth'         ; SamAccountName = 'fancy'              ; UserPrincipalName = 'fancy@example.net'              ; OU = @('Locations','Germany','Bayern','Munich','Fancy')                               ; Groups = @('Fancy Solo','Vocalists')                    ; Title = 'Solo Artist'                 ; Email = 'fancy@example.net'              ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Munich Office'     ; Phone = '+49 89 3333 111'  ; MobilePhone = '+49 1723 3111'   ; Street = 'Leopoldstrasse 100'     ; City = 'Munich'     ; PostalCode = '80802'     ; Company = 'Example Music GmbH'   ; Manager = ''                ; Description = 'Solo artist aus Munich, known for Lady of Ice & Italo disco'  },

    ## ==========  Nena - Germany/West Berlin ==========
    @{ Name = 'Gabriele Kerner'          ; SamAccountName = 'nena'              ; UserPrincipalName = 'nena@example.net'               ; OU = @('Locations','Germany','West Berlin','Nena')                                    ; Groups = @('Nena','Vocalists')                          ; Title = 'Lead Vocalist'               ; Email = 'nena@example.net'               ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 111'  ; MobilePhone = '+49 1712 2111'   ; Street = 'Kurfurstendamm 100'     ; City = 'Berlin'     ; PostalCode = '10709'     ; Company = 'Example Music GmbH'   ; Manager = ''                ; Description = 'Lead vocalist of Nena, known as Nena'                         },
    @{ Name = 'Carlo Karges'             ; SamAccountName = 'carlo.karges'      ; UserPrincipalName = 'carlo.karges@example.net'       ; OU = @('Locations','Germany','West Berlin','Nena')                                    ; Groups = @('Nena','Musicians')                          ; Title = 'Guitarist'                   ; Email = 'carlo.karges@example.net'       ; Country = 'DE' ; Disabled = $true  ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 112'  ; MobilePhone = '+49 1712 2112'   ; Street = 'Kurfurstendamm 101'     ; City = 'Berlin'     ; PostalCode = '10709'     ; Company = 'Example Music GmbH'   ; Manager = ''                ; Description = 'Guitarist for Nena (1951-2002) - Account disabled'            },
    @{ Name = 'Uwe Fahrenkrog-Petersen'  ; SamAccountName = 'uwe.fahrenkrog'    ; UserPrincipalName = 'uwe.fahrenkrog@example.net'     ; OU = @('Locations','Germany','West Berlin','Nena')                                    ; Groups = @('Nena','Musicians')                          ; Title = 'Keyboardist'                 ; Email = 'uwe.fahrenkrog@example.net'     ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 113'  ; MobilePhone = '+49 1712 2113'   ; Street = 'Kurfurstendamm 102'     ; City = 'Berlin'     ; PostalCode = '10709'     ; Company = 'Example Music GmbH'   ; Manager = 'Gabriele Kerner' ; Description = 'Keyboardist and songwriter for Nena'                          },
    @{ Name = 'Jurgen Dehmel'            ; SamAccountName = 'jurgen.dehmel'     ; UserPrincipalName = 'jurgen.dehmel@example.net'      ; OU = @('Locations','Germany','West Berlin','Nena')                                    ; Groups = @('Nena','Musicians')                          ; Title = 'Bassist'                     ; Email = 'jurgen.dehmel@example.net'      ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 114'  ; MobilePhone = '+49 1712 2114'   ; Street = 'Kurfurstendamm 103'     ; City = 'Berlin'     ; PostalCode = '10709'     ; Company = 'Example Music GmbH'   ; Manager = 'Gabriele Kerner' ; Description = 'Bassist for Nena'                                             },
    @{ Name = 'Rolf Brendel'             ; SamAccountName = 'rolf.brendel'      ; UserPrincipalName = 'rolf.brendel@example.net'       ; OU = @('Locations','Germany','West Berlin','Nena')                                    ; Groups = @('Nena','Musicians','Percussionists')         ; Title = 'Drummer'                     ; Email = 'rolf.brendel@example.net'       ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 115'  ; MobilePhone = '+49 1712 2115'   ; Street = 'Kurfurstendamm 104'     ; City = 'Berlin'     ; PostalCode = '10709'     ; Company = 'Example Music GmbH'   ; Manager = 'Gabriele Kerner' ; Description = 'Drummer for Nena'                                             },

    ## ========== Tangerine Dream - Germany/West Berlin ==========
    @{ Name = 'Edgar Froese'            ; SamAccountName = 'edgar.froese'       ; UserPrincipalName = 'edgar.froese@example.net'       ; OU = @('Locations','Germany','West Berlin','Tangerine Dream')                         ; Groups = @('Tangerine Dream','Musicians')               ; Title = 'Synthesizer/Guitar'          ; Email = 'edgar.froese@example.net'       ; Country = 'DE' ; Disabled = $true  ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 121'  ; MobilePhone = '+49 171 2221121' ; Street = 'Kantstrasse 50'         ; City = 'Berlin'     ; PostalCode = '10625'     ; Company = 'Example Music GmbH'   ; Manager = ''                ; Description = 'Founder of Tangerine Dream (1944-2015) - Account disabled'    },
    @{ Name = 'Christopher Franke'      ; SamAccountName = 'christopher.franke' ; UserPrincipalName = 'christopher.franke@example.net' ; OU = @('Locations','Germany','West Berlin','Tangerine Dream')                         ; Groups = @('Tangerine Dream','Musicians')               ; Title = 'Synthesizer/Drums'           ; Email = 'christopher.franke@example.net' ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 122'  ; MobilePhone = '+49 171 2221122' ; Street = 'Kantstrasse 51'         ; City = 'Berlin'     ; PostalCode = '10625'     ; Company = 'Example Music GmbH'   ; Manager = ''                ; Description = 'Synthesizer and electronic drums for Tangerine Dream'         },
    @{ Name = 'Peter Baumann'           ; SamAccountName = 'peter.baumann'      ; UserPrincipalName = 'peter.baumann@example.net'      ; OU = @('Locations','Germany','West Berlin','Tangerine Dream')                         ; Groups = @('Tangerine Dream','Musicians')               ; Title = 'Synthesizer'                 ; Email = 'peter.baumann@example.net'      ; Country = 'DE' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Berlin Office'     ; Phone = '+49 30 2222 123'  ; MobilePhone = '+49 171 2221123' ; Street = 'Kantstrasse 52'         ; City = 'Berlin'     ; PostalCode = '10625'     ; Company = 'Example Music GmbH'   ; Manager = ''                ; Description = 'Synthesizer player for Tangerine Dream'                       },

    ## ========== Fiction Factory - Scotland/Perth ==========
    @{ Name = 'Kevin Patterson'        ; SamAccountName = 'kpatterson'          ; UserPrincipalName = 'kpatterson@example.com'         ; OU = @('Locations','UK','Scotland','Perth','Fiction Factory')                         ; Groups = @('Fiction Factory','Keyboards')               ; Title='Keyboardist'                   ; Email = 'kevin.patterson@example.com'    ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Perth Office'      ; Phone = '+44 1738 496 011' ; MobilePhone = '+44 7700 173001' ; Street = '10 High Street'         ; City = 'Perth'      ; PostalCode = 'PH1 5AA'   ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Keyboardist/Synthesizers for Fiction Factory'                 },
    @{ Name = 'Eddie Jordan'           ; SamAccountName = 'ejordan'             ; UserPrincipalName = 'ejordan@example.com'            ; OU = @('Locations','UK','Scotland','Perth','Fiction Factory')                         ; Groups = @('Fiction Factory','Drummers')                ; Title='Drummer'                       ; Email = 'eddie.jordan@example.com'       ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Perth Office'      ; Phone = '+44 1738 496 012' ; MobilePhone = '+44 7700 173002' ; Street = '10 High Street'         ; City = 'Perth'      ; PostalCode = 'PH1 5AA'   ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Drummer for Fiction Factory'                                  },
    @{ Name = 'Mike Ogletree'          ; SamAccountName = 'mogletree'           ; UserPrincipalName = 'mogletree@example.com'          ; OU = @('Locations','UK','Scotland','Perth','Fiction Factory')                         ; Groups = @('Fiction Factory','Percussion')              ; Title='Percussionist'                 ; Email = 'mike.ogletree@example.com'      ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Perth Office'      ; Phone = '+44 1738 496 013' ; MobilePhone = '+44 7700 173003' ; Street = '10 High Street'         ; City = 'Perth'      ; PostalCode = 'PH1 5AA'   ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Percussionist for Fiction Factory'                            },
    @{ Name = 'Eddie Campbell'         ; SamAccountName = 'ecampbell'           ; UserPrincipalName = 'ecampbell@example.com'          ; OU = @('Locations','UK','Scotland','Perth','Fiction Factory')                         ; Groups = @('Fiction Factory','Guitarists')              ; Title='Guitarist'                     ; Email = 'eddie.campbell@example.com'     ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Perth Office'      ; Phone = '+44 1738 496 014' ; MobilePhone = '+44 7700 173004' ; Street = '10 High Street'         ; City = 'Perth'      ; PostalCode = 'PH1 5AA'   ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Guitarist for Fiction Factory'                                },
    @{ Name = 'Graham McGregor'        ; SamAccountName = 'gmcgregor'           ; UserPrincipalName = 'gmcgregor@example.com'          ; OU = @('Locations','UK','Scotland','Perth','Fiction Factory')                         ; Groups = @('Fiction Factory','Bassists')                ; Title='Bassist'                       ; Email = 'graham.mcgregor@example.com'    ; Country = 'UK' ; Disabled = $false ; Locked = $false ; MustChangePassword = $false ; Department = 'Music' ; Office = 'Perth Office'      ; Phone = '+44 1738 496 015' ; MobilePhone = '+44 7700 173005' ; Street = '10 High Street'         ; City = 'Perth'      ; PostalCode = 'PH1 5AA'   ; Company = 'Example Music Ltd'    ; Manager = ''                ; Description = 'Bassist for Fiction Factory'                                  },

    ## Dummy user- for setting up new users via template - as is common
    @{ Name = 'Template User'          ; SamAccountName = 'template.users'      ; UserPrincipalName = 'template.user@example.com'      ; OU = @('Locations','Templates','Users')                                               ; Groups = @('Disabled Users','Template Users')           ; Title = 'Template User'               ; Email = 'template.user@example.com'      ; Country = 'GL' ; Disabled = $true  ; Locked = $true  ; MustChangePassword = $true  ; Department = 'Music' ; Office = 'Nuuk, Grønland'    ; Phone = '+00 00 0000 000'  ; MobilePhone = '+00 000 0000000' ; Street = '1 Example Street'       ; City = 'Nuuk'        ; PostalCode = '00000'    ; Company = 'Example Music ApS'    ; Manager = 'Prins Knud'      ; Description = 'Counting Paperclips - En gang til Prins Knud'                 }

  ## Demo Users ending stanza
  )

  ## ------------------ Define Demo Groups ------------------
  $Script:rawDemoGroups = @(
    @{ Name = 'Simple Minds'          ; Description = 'Scottish rock band formed in Glasgow in 1977'                  ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Jim Kerr'           ; Email = 'simpleminds@example.com'    },
    @{ Name = 'Depeche Mode'          ; Description = 'English electronic music band formed in Basildon in 1980'      ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Dave Gahan'         ; Email = 'depechemode@example.com'    },
    @{ Name = 'Erasure'               ; Description = 'English synth-pop duo formed in London in 1985'                ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Andy Bell'          ; Email = 'erasure@example.com'        },
    @{ Name = 'Marillion'             ; Description = 'British rock band formed in Edinburgh in 1979'                 ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Steve Hogarth'      ; Email = 'marillion@example.com'      },
    @{ Name = 'TV-2'                  ; Description = 'Danish rock band formed in 1981'                               ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Steffen Brandt'     ; Email = 'tv2@example.com'            },
    @{ Name = 'Rocazino'              ; Description = 'Danish pop band from Koge'                                     ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Stine Bramsen'      ; Email = 'alphabeat@example.com'      },
    @{ Name = 'MØ Solo'               ; Description = 'Solo artist MØ from Odense'                                    ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Karen Marie Ørsted' ; Email = 'mo@example.com'             },
    @{ Name = 'Kraftwerk'             ; Description = 'German electronic music pioneers from Bonn, formed 1970'       ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Ralf Hutter'        ; Email = 'kraftwerk@example.net'      },
    @{ Name = 'Nena'                  ; Description = 'German Neue Deutsche Welle band from West Berlin, formed 1982' ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Gabriele Kerner'    ; Email = 'nena@example.net'           },
    @{ Name = 'Tangerine Dream'       ; Description = 'German electronic music group from West Berlin, formed 1967'   ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Christopher Franke' ; Email = 'tangerinedream@example.net' },
    @{ Name = 'Fancy Solo'            ; Description = 'Solo artist Fancy from Munich, Italo disco performer'          ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Manfred Segieth'    ; Email = 'fancy@example.net'          },
    @{ Name = 'Echo And The Bunnymen' ; Description = 'Echo And The Bunnymen Groupr'                                  ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Ian McCulloch'      ; Email = 'bunnymen@example.org'       },
    @{ Name = 'The Police'            ; Description = 'Group Catch-all for The Police'                                ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Sting'              ; Email = 'thepolice@example.org'      },
    @{ Name = 'Altered Images'        ; Description = 'New Wave band formed in Glasgow 1980'                          ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Clare Grogan'       ; Email = 'alteredimages@example.org'  },
    @{ Name = 'Eurythmics'            ; Description = 'Duo formed in Aberdeen in 1980'                                ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Annie Lennox'       ; Email = 'eurythmics@example.org'     },
    @{ Name = 'New Order'             ; Description = 'English rock band formed in Salford in 1980'                   ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Tony Wilson'        ; Email = 'neworder@example.org'       },
    @{ Name = 'Ultravox'              ; Description = 'Scottish New Wave band, classic Midge Ure era lineup'          ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Midge Ure'          ; Email = 'ultravox@example.com'       },
    @{ Name = 'Deacon Blue'           ; Description = 'Scottish pop rock band formed in Glasgow in 1985'              ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Ricky Ross'         ; Email = 'deaconblue@example.com'     },
    @{ Name = 'Mel And Kim'           ; Description = 'Melanie and Kimberly Appleby'                                  ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Pete Waterman'      ; Email = 'melandkim@example.com'      },
    @{ Name = 'UB40'                  ; Description = 'Band named after the Unemployment Benefit 40 Form in 1978'     ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Margaret Thatcher'  ; Email = 'ub40@example.net'           },
    @{ Name = 'Wet Wet Wet'           ; Description = 'Scottish pop rock band formed in Clydebank in 1982'            ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Marti Pellow'       ; Email = 'wetwetwet@example.net'      },
    @{ Name = 'The Proclaimers'       ; Description = 'Scottish folk rock duo formed in Edinburgh in 1983'            ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Craig Reid'         ; Email = 'proclaimers@example.net'    },
    @{ Name = 'Fiction Factory'       ; Description = 'Scottish synth band formed in Perth, Scotland in 1982'         ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Mike Ogletree'      ; Email ='fictionfactory@example.net'  },

    ## Occupation groups. Normally used for distirb lists in Exchange (365)
    @{ Name = 'Vocalists'             ; Description = 'Lead singers and vocalists across all bands'                   ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = ''                   ; Email = 'vocalists@example.com'      },
    @{ Name = 'Keyboards'             ; Description = 'Keyboard And synthesizers'                                     ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = ''                   ; Email = 'synths@example.com'         },
    @{ Name = 'Musicians'             ; Description = 'Instrumentalists'                                              ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = ''                   ; Email = 'instruments@example.com'    },
    @{ Name = 'Guitarists'            ; Description = 'Guitar and bass players'                                       ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = ''                   ; Email = 'guitars@example.com'        },
    @{ Name = 'Percussionists'        ; Description = 'Drummers and percussion specialists'                           ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = ''                   ; Email = 'percussion@example.com'     },
    @{ Name = 'Marketing'             ; Description = 'Marketing and Creattive Group      '                           ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Helena Christensen' ; Email = 'marketing@example.com'      },
    ## Company Lawyer
    @{ Name = 'Legal'                 ; Description = 'Company Lawyer Messers Sue, Grabbit & Runne'                   ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'Arkell v Pressdram' ; Email = 'legal@example.com'          },

    ## Non-band based groups
    @{ Name = 'Sales'                 ; Description = 'Remote VPN users and road warriors in sales'                   ; Type = 'Security' ; Scope = 'Global' ; ManagedBy ='IT Admin'            ; Email='sales@example.com'            },
    @{ Name = 'Former Staff'          ; Description = 'Former band members and staff'                                 ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = ''                   ; Email = ''                           },
    @{ Name = 'VPN-Users'             ; Description = 'Remote access users (touring staff, laptops and tablets)'      ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = 'IT Operations'      ; Email = 'vpn-users@example.net'      },
    @{ Name = 'Disabled Users'        ; Description = 'Administratively disabled user accounts'                       ; Type = 'Security' ; Scope = 'Global' ; ManagedBy = ''                   ; Email = ''                           }
  )

  ## ------------------ Define Demo Domain Controllers ------------------
  ## Updated Demo Domain Controllers - Matching Production AD Structure

  $Script:rawDCs = @(
    ## ===== UK DCs (example.com) =====
    @{ Name = 'EXAGLADC01' ; SamAccountName = 'EXAGLADC01$' ; DNSHostName = 'EXAGLADC01.example.com' ; HostName = 'EXAGLADC01.example.com' ; Site = 'GLA' ; Location = 'Glasgow, Scotland'    ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.4.20'    ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @('Schema Master', 'Domain Naming Master', 'PDC Emulator')  ; OperationMasterRoles = @('Schema Master', 'Domain Naming Master', 'PDC Emulator')     ; LastReplication = (Get-Date).AddMinutes(-12) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-45)    ; LastBootUpTime = (Get-Date).AddDays(-45)   ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '45 GB' ; Free = '75 GB' ; PercentFree = 62} ; 'SYSVOL' = @{Total = '50 GB'  ; Used = '8 GB'  ; Free = '42 GB'  ; PercentFree = 84} }  ; ReplicationPartners = @('EXALNDDC01', 'EXAEDIDC01')                              },
    @{ Name = 'EXAEDIDC01' ; SamAccountName = 'EXAEDIDC01$' ; DNSHostName = 'EXAEDIDC01.example.com' ; HostName = 'EXAEDIDC01.example.com' ; Site = 'EDI' ; Location = 'Edinburgh, Scotland'  ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2019 Standard' ; OperatingSystemVersion = '10.0 (17763)' ; IPv4Address = '192.168.3.20'    ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @('PDC Emulator', 'RID Master', 'Infrastructure Master')    ; OperationMasterRoles = @()                                                            ; LastReplication = (Get-Date).AddMinutes(-18) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-67)    ; LastBootUpTime = (Get-Date).AddDays(-67)   ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '38 GB' ; Free = '82 GB' ; PercentFree = 68} ; 'SYSVOL' = @{Total = '50 GB'  ; Used = '6 GB'  ; Free = '44 GB'  ; PercentFree = 88} }  ; ReplicationPartners = @('EXAGLADC01', 'EXALNDDC01','EXANEWDC01', 'EXALIVDC01')   },
    @{ Name = 'EXALNDDC01' ; SamAccountName = 'EXALNDDC01$' ; DNSHostName = 'EXALNDDC01.example.com' ; HostName = 'EXALNDDC01.example.com' ; Site = 'LND' ; Location = 'London, England'      ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.2.20'    ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @('RID Master', 'Infrastructure Master')                    ; OperationMasterRoles = @('RID Master', 'Infrastructure Master')                       ; LastReplication = (Get-Date).AddMinutes(-8)  ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-23)    ; LastBootUpTime = (Get-Date).AddDays(-23)   ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '52 GB' ; Free = '68 GB' ; PercentFree = 57} ; 'SYSVOL' = @{Total = '50 GB'  ; Used = '12 GB' ; Free = '38 GB'  ; PercentFree = 76} }  ; ReplicationPartners = @('EXAGLADC01', 'EXAEDIDC01', 'EXAKBRDC01', 'EXACPHDC01')  },

    ## ===== UK DCs (example.org) =====
    @{ Name = 'EXANEWDC01' ; SamAccountName = 'EXANEWDC01$' ; DNSHostName = 'EXANEWDC01.example.org' ; HostName = 'EXANEWDC01.example.org' ; Site = 'NEW' ; Location = 'Newcastle, UK'        ; Domain = 'example.org' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.91.10'   ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @()                                                         ; OperationMasterRoles = @()                                                            ; LastReplication = (Get-Date).AddMinutes(-13) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-42)    ; LastBootUpTime = (Get-Date).AddDays(-42)   ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '24 GB' ; Free = '96 GB' ; PercentFree = 80} ; 'SYSVOL' = @{Total = '50 GB'  ; Used = '8 GB'  ; Free = '42 GB'  ; PercentFree = 84} }  ; ReplicationPartners = @('EXALIVDC01', 'EXAEDIDC01', 'EXAGLADC01', 'EXALNDDC01') },
    @{ Name = 'EXALIVDC01' ; SamAccountName = 'EXALIVDC01$' ; DNSHostName = 'EXALIVDC01.example.org' ; HostName = 'EXALIVDC01.example.org' ; Site = 'LIV' ; Location = 'Liverpool, UK'        ; Domain = 'example.org' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2025 Standard' ; OperatingSystemVersion = '10.0 (26100)' ; IPv4Address = '192.168.51.11'   ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @()                                                         ; OperationMasterRoles = @()                                                            ; LastReplication = (Get-Date).AddMinutes(-15) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-12)    ; LastBootUpTime = (Get-Date).AddDays(-12)   ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '46 GB' ; Free = '74 GB' ; PercentFree = 62} ; 'SYSVOL' = @{Total = '50 GB'  ; Used = '8 GB'  ; Free = '42 GB'  ; PercentFree = 84} }  ; ReplicationPartners = @('EXANEWDC01', 'EXAEDIDC01')                             },
    @{ Name = 'EXAMCRDC01' ; SamAccountName = 'EXAMCRDC01$' ; DNSHostName = 'EXAMCRDC01.example.org' ; HostName = 'EXAMCRDC01.example.org' ; Site = 'MCR' ; Location = 'Manchester, UK'       ; Domain = 'example.org' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.61.10'   ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @('PDC Emulator', 'RID Master', 'Infrastructure Master')    ; OperationMasterRoles = @('PDC Emulator', 'RID Master', 'Infrastructure Master')       ; LastReplication = (Get-Date).AddMinutes(-12) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-7)     ; LastBootUpTime = (Get-Date).AddDays(-7)    ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '80 GB' ; Used = '49 GB' ; Free = '40 GB' ; PercentFree = 50} ; 'SYSVOL' = @{Total = '50 GB'  ; Used = '28 GB' ; Free = '22 GB'  ; PercentFree = 44} }  ; ReplicationPartners = @('EXAGLADC01', 'EXAEDIDC01', 'EXAKBRDC01', 'EXACPHDC01') },
    @{ Name = 'EXAMCRDC02' ; SamAccountName = 'EXAMCRDC02$' ; DNSHostName = 'EXAMCRDC02.example.org' ; HostName = 'EXAMCRDC02.example.org' ; Site = 'MCR' ; Location = 'Manchester, UK'       ; Domain = 'example.org' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.61.11'   ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @()                                                         ; OperationMasterRoles = @()                                                            ; LastReplication = (Get-Date).AddMinutes(-6)  ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-3)     ; LastBootUpTime = (Get-Date).AddDays(-3)    ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '100 GB'; Used = '78 GB' ; Free = '22 GB' ; PercentFree = 22} ; 'SYSVOL' = @{Total = '15 GB'  ; Used = '10 GB' ; Free = '5 GB'   ; PercentFree = 33} }  ; ReplicationPartners = @('EXALNDDC01', 'EXAKBRDC01', 'EXAODEDC01')               },

    ## ===== UK DCs (example.net) =====
    @{ Name = 'EXADCBIR01' ; SamAccountName = 'EXADCBIR01$' ; DNSHostName = 'EXADCBIR01.example.net' ; HostName = 'EXADCBIR01.example.net' ; Site = 'BIR' ; Location = 'Birmingham, UK'       ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.121.10'  ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @()                                                         ; OperationMasterRoles = @()                                                            ; LastReplication = (Get-Date).AddMinutes(-46) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-120)   ; LastBootUpTime = (Get-Date).AddDays(-120)  ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '58 GB' ; Free = '62 GB' ; PercentFree = 52} ; 'SYSVOL' = @{Total = '50 GB'  ; Used = '11 GB' ; Free = '39 GB'  ; PercentFree = 78} }  ; ReplicationPartners = @('EXABONDC01', 'EXAMUCDC01')                             },
    @{ Name = 'EXADCBIR02' ; SamAccountName = 'EXADCBIR02$' ; DNSHostName = 'EXADCBIR02.example.net' ; HostName = 'EXADCBIR02.example.net' ; Site = 'BIR' ; Location = 'Birmingham, UK'       ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.121.11'  ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @()                                                         ; OperationMasterRoles = @()                                                            ; LastReplication = (Get-Date).AddMinutes(-3)  ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-11)    ; LastBootUpTime = (Get-Date).AddDays(-11)   ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '500 GB'; Used = '10 GB' ; Free = '490 GB'; PercentFree = 98} ; 'SYSVOL' = @{Total = '100GB'  ; Used = '6 GB'  ; Free = '94 GB'  ; PercentFree = 94} }  ; ReplicationPartners = @('EXABERDC01', 'EXABRKDC01')                             },
    @{ Name = 'EXADCCLY01' ; SamAccountName = 'EXADCCLY01$' ; DNSHostName = 'EXADCCLY01.example.net' ; HostName = 'EXADCCLY01.example.net' ; Site = 'CLY' ; Location = 'Clydebank, UK'        ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.141.9'   ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @()                                                         ; OperationMasterRoles = @()                                                            ; LastReplication = (Get-Date).AddMinutes(-38) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-18)    ; LastBootUpTime = (Get-Date).AddDays(-18)   ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '80 GB' ; Used = '61 GB' ; Free = '19 GB' ; PercentFree = 24} ; 'SYSVOL' = @{Total = '200GB'  ; Used = '18 GB' ; Free = '182GB'  ; PercentFree = 91} }  ; ReplicationPartners = @('EXAGLADC01', 'EXAEDIDC01', 'EXAKBRDC01', 'EXACPHDC01') },
    @{ Name = 'EXADCEDI03' ; SamAccountName = 'EXADCEDI03$' ; DNSHostName = 'EXADCEDI03.example.net' ; HostName = 'EXADCEDI03.example.net' ; Site = 'EDI' ; Location = 'Edinburgh, UK'        ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.131.10'  ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $false ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @('RID Master', 'Infrastructure Master')                    ; OperationMasterRoles = @('RID Master', 'Infrastructure Master')                       ; LastReplication = (Get-Date).AddMinutes(-17) ; ReplicationHealth = 'Unhealthy'             ; LastBoot = (Get-Date).AddDays(-60)    ; LastBootUpTime = (Get-Date).AddDays(-60)   ; Services = @{DNS = 'Running'; DFSR = 'Stopped' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '114 GB'; Free = '6 GB'  ; PercentFree = 5 } ; 'SYSVOL' = @{Total = '40 GB'  ; Used = '32 GB' ; Free = '8 GB'   ; PercentFree = 20} }  ; ReplicationPartners = @('EXALIVDC01', 'EXAEDIDC01', 'EXAGLADC01', 'EXALNDDC01') },
    @{ Name = 'EXADCCLY02' ; SamAccountName = 'EXADCCLY02$' ; DNSHostName = 'EXADCCLY02.example.net' ; HostName = 'EXADCCLY02.example.net' ; Site = 'CLY' ; Location = 'Clydebank Office'     ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.141.11'  ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @()                                                         ; OperationMasterRoles = @()                                                            ; LastReplication = (Get-Date).AddMinutes(-31) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-4)     ; LastBootUpTime = (Get-Date).AddDays(-4)    ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '150 GB'; Used = '123 GB'; Free = '27 GB' ; PercentFree = 18} ; 'SYSVOL' = @{Total = '60 GB'  ; Used = '29 GB' ; Free = '31 GB'  ; PercentFree = 52} }  ; ReplicationPartners = @('EXADCCLY01', 'EXAGLADC01')                             },
    @{ Name = 'EXADCDUN01' ; SamAccountName = 'EXADCDUN01$' ; DNSHostName = 'EXADCDUN01.example.net' ; HostName = 'EXADCDUN01.example.net' ; Site = 'DUN' ; Location = 'Dundee, UK'           ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPAddress   = '192.168.138.9'   ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @()                                                         ; OperationMasterRoles = @()                                                            ; LastReplication = (Get-Date).AddMinutes(-31) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-4)     ; LastBootUpTime = (Get-Date).AddDays(-4)    ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '200 GB'; Used = '105 GB'; Free = '95 GB' ; PercentFree = 42} ; 'SYSVOL' = @{Total = '50 GB'  ; Used = '20 GB' ; Free = '30 GB'  ; PercentFree = 40} }  ; ReplicationPartners = @('EXADCCLY01', 'EXAGLADC01', 'EXAEDIDC01', 'EXADCEDI02') },
    @{ Name = 'EXADCPER01' ; SamAccountName = 'EXADCPER01$' ; DNSHostName = 'EXADCPER01.example.net' ; HostName = 'EXADCPER01.example.net' ; Site = 'PER' ; Location = 'Perth, UK'            ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPAddress   = '192.168.173.10'  ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @()                                                         ; OperationMasterRoles = @()                                                            ; LastReplication = (Get-Date).AddMinutes(-25) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-5)     ; LastBootUpTime = (Get-Date).AddDays(-5)    ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '200 GB'; Used = '110 GB'; Free = '90 GB' ; PercentFree = 45} ; 'SYSVOL' = @{Total = '50 GB'  ; Used='22 GB'   ; Free = '28 GB'  ; PercentFree = 44} }  ; ReplicationPartners = @('EXADCEDI03', 'EXAGLADC01', 'EXAEDIDC01', 'EXADCLND01') },

    ## ===== Danmark DCs (example.com) =====
    @{ Name = 'EXACPHDC01' ; SamAccountName ='EXACPHDC01$'  ; DNSHostName = 'EXACPHDC01.example.com' ; HostName = 'EXACPHDC01.example.com' ; Site = 'CPH' ; Location = 'Copenhagen, Danmark'  ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.6.20'    ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @()                                                         ; OperationMasterRoles = @()                                                            ; LastReplication = (Get-Date).AddMinutes(-41) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-34)    ; LastBootUpTime = (Get-Date).AddDays(-34)   ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '41 GB' ; Free = '79 GB' ; PercentFree = 66} ; 'SYSVOL' = @{Total = '50 GB'  ; Used = '9 GB'  ; Free = '41 GB'  ; PercentFree = 82} }  ; ReplicationPartners = @('EXALNDDC01', 'EXAKBRDC01', 'EXAODEDC01')               },
    @{ Name = 'EXAKBRDC01' ; SamAccountName ='EXAKBRDC01$'  ; DNSHostName = 'EXAKBRDC01.example.com' ; HostName = 'EXAKBRDC01.example.com' ; Site = 'KGE' ; Location = 'Køge, Danmark'        ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2016 Standard' ; OperatingSystemVersion = '10.0 (14393)' ; IPv4Address = '192.168.5.20'    ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $false ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @()                                                         ; OperationMasterRoles = @()                                                            ; LastReplication = (Get-Date).AddDays(-27)    ; ReplicationHealth = 'Warning - Out of Sync' ; LastBoot = (Get-Date).AddDays(-156)   ; LastBootUpTime = (Get-Date).AddDays(-156)  ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '100 GB'; Used = '78 GB' ; Free = '22 GB' ; PercentFree = 22} ; 'SYSVOL' = @{Total = '40 GB'  ; Used = '28 GB' ; Free = '12 GB'  ; PercentFree = 30} }  ; ReplicationPartners = @('EXALNDDC01', 'EXACPHDC01')                             },
    @{ Name = 'EXAODEDC01' ; SamAccountName ='EXAODEDC01$'  ; DNSHostName = 'EXAODEDC01.example.com' ; HostName = 'EXAODEDC01.example.com' ; Site = 'ODE' ; Location = 'Odense, Danmark'      ; Domain = 'example.com' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.7.20'    ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @('PDC Emulator', 'RID Master', 'Infrastructure Master')    ; OperationMasterRoles = @()                                                            ; LastReplication = (Get-Date).AddMinutes(-39) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-28)    ; LastBootUpTime = (Get-Date).AddDays(-28)   ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'} ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '42 GB' ; Free = '78 GB' ; PercentFree = 65} ; 'SYSVOL' = @{Total = '50 GB'  ; Used = '7 GB'  ; Free = '43 GB'  ; PercentFree = 86} }  ; ReplicationPartners = @('EXACPHDC01', 'EXAKBRDC01')                             },

    ## ===== Germany DCs (example.net) =====
    @{ Name = 'EXABONDC01' ; SamAccountName = 'EXABONDC01$' ; DNSHostName = 'EXABONDC01.example.net' ; HostName = 'EXABONDC01.example.net' ; Site = 'BON' ; Location = 'Bonn, Germany'        ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.22.20'   ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @('Schema Master', 'Domain Naming Master')                  ; OperationMasterRoles = @('Schema Master', 'Domain Naming Master')                     ; LastReplication = (Get-Date).AddMinutes(-11) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-38)    ; LastBootUpTime = (Get-Date).AddDays(-38)   ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'}  ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '48 GB'; Free = '72 GB' ; PercentFree = 60} ; 'SYSVOL' = @{Total = '50 GB'  ; Used = '10 GB' ; Free = '40 GB'  ; PercentFree = 80} }  ; ReplicationPartners = @('EXABERDC01')                                           },
    @{ Name = 'EXABERDC01' ; SamAccountName = 'EXABERDC01$' ; DNSHostName = 'EXABERDC01.example.net' ; HostName = 'EXABERDC01.example.net' ; Site = 'BER' ; Location = 'West Berlin, Germany' ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2019 Standard' ; OperatingSystemVersion = '10.0 (17763)' ; IPv4Address = '192.168.30.20'   ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @('PDC Emulator', 'RID Master', 'Infrastructure Master')    ; OperationMasterRoles = @('PDC Emulator', 'RID Master', 'Infrastructure Master')       ; LastReplication = (Get-Date).AddMinutes(-9)  ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-51)    ; LastBootUpTime = (Get-Date).AddDays(-51)   ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'}  ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '55 GB'; Free = '65 GB' ; PercentFree = 54} ; 'SYSVOL' = @{Total = '50 GB'  ; Used = '14 GB' ; Free = '36 GB'  ; PercentFree = 72} }  ; ReplicationPartners = @('EXABONDC01', 'EXAMUCDC01')                             },
    @{ Name = 'EXAMUCDC01' ; SamAccountName = 'EXAMUCDC01$' ; DNSHostName = 'EXAMUCDC01.example.net' ; HostName = 'EXAMUCDC01.example.net' ; Site = 'MUC' ; Location = 'Munich, Germany'      ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.89.20'   ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @()                                                         ; OperationMasterRoles = @('Schema Master', 'Domain Naming Master')                     ; LastReplication = (Get-Date).AddMinutes(-13) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-42)    ; LastBootUpTime = (Get-Date).AddDays(-42)   ; Services = @{DNS = 'Running'; DFSR = 'Running' ; Netlogon = 'Running' ; KDC = 'Running'}  ; DiskSpace = @{ 'C:' = @{Total = '120 GB'; Used = '44 GB'; Free = '76 GB' ; PercentFree = 63} ; 'SYSVOL' = @{Total = '50 GB'  ; Used = '8 GB'  ; Free = '42 GB'  ; PercentFree = 84} }  ; ReplicationPartners = @('EXABONDC01', 'EXABERDC01')                             },

    ## ===== Canada DCs (example.net) ====
    @{ Name = 'EXABRKDC01' ; SamAccountName = 'EXABRKDC01$' ; DNSHostName = 'EXABRKDC01.example.net' ; HostName = 'EXABRKDC01.example.net' ; Site = 'BRK' ; Location = 'Brockville, Ontario'  ; Domain = 'example.net' ; Forest = 'jukebox.example' ; OS = 'Windows Server 2022 Standard' ; OperatingSystemVersion = '10.0 (20348)' ; IPv4Address = '192.168.136.20'  ; IPv6Address = $null ; Enabled = $true ; IsGlobalCatalog = $true  ; IsReadOnly = $false ; LdapPort = 389 ; SslPort = 636 ; FSMORoles = @()                                                         ; OperationMasterRoles = @('Schema Master', 'Domain Naming Master')                     ; LastReplication = (Get-Date).AddMinutes(-19) ; ReplicationHealth = 'Healthy'               ; LastBoot = (Get-Date).AddDays(-12)    ; LastBootUpTime = (Get-Date).AddDays(-12)   ; Services = @{DNS = 'Stopped'; DFSR = 'Running' ; Netlogon = 'Stopped' ; KDC = 'Stopped'}  ; DiskSpace = @{ 'C:' = @{Total = '50 GB' ; Used = '24 GB'; Free = '26 GB' ; PercentFree = 54}  ; 'SYSVOL' = @{Total = '30 GB'; Used = '15 GB'  ; Free = '15 GB'  ; PercentFree = 50} }  ; ReplicationPartners = @('EXAEDIDC01', 'EXAODNDC01')                             }
  )
  Debug-Log ": Loaded $($Script:rawDCs.Count) demo domain controllers" -Type "Info"

  ## --------{ Demo Computers (Workstations, Laptops, Printers) }--------
  $Script:rawComputers = @(

    ## Aberdeen, Scotland
    @{ Name = 'EXAPHNABD001' ; SamAccountName = 'EXAPHNABD001$' ; Type = 'Phones'        ; Role = 'PHN' ; Site = 'ABD' ; Location = 'Aberdeen, UK'                 ; DNSHostName = 'EXAPHNABD001.example.org' ; OU = @('Locations','UK','Scotland','Aberdeen','Devices')            ; OS = 'iOS'                               ; OperatingSystemVersion = '17.x'                 ; Description = 'Corporate iPhone – Annie Lennox'                ; LastLogon = 'N/A'                      ; IPAddress = 'N/A'             ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  };
    @{ Name = 'EXAPHNABD002' ; SamAccountName = 'EXAPHNABD002$' ; Type = 'Phones'        ; Role = 'PHN' ; Site = 'ABD' ; Location = 'Aberdeen, UK'                 ; DNSHostName = 'EXAPHNABD002.example.org' ; OU = @('Locations','UK','Scotland','Aberdeen','Devices')            ; OS = 'iOS'                               ; OperatingSystemVersion = '17.x'                 ; Description = 'Corporate iPhone – Dave Stewart'                ; LastLogon = 'N/A'                      ; IPAddress = 'N/A'             ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXAMBPABD001' ; SamAccountName = 'EXAMBPABD001$' ; Type = 'Computer'      ; Role = 'MBP' ; Site = 'ABD' ; Location = 'Aberdeen, UK'                 ; DNSHostName = 'EXAMBPABD001.example.org' ; OU = @('Locations','UK','Scotland','Aberdeen','Computers')          ; OS = 'macOS Sonoma'                      ; OperatingSystemVersion = '14.2.1'               ; Description = 'MacBook – Annie Lennox (Aberdeen satellite)'    ; LastLogon = (Get-Date).AddDays(-16)    ; IPAddress = '192.168.224.137' ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXAMBPABD002' ; SamAccountName = 'EXAMBPABD002$' ; Type = 'Computer'      ; Role = 'MBP' ; Site = 'ABD' ; Location = 'Aberdeen, UK'                 ; DNSHostName = 'EXAMBPABD002.example.org' ; OU = @('Locations','UK','Scotland','Aberdeen','Computers')          ; OS = 'macOS Sonoma'                      ; OperatingSystemVersion = '14.2.1'               ; Description = 'MacBook – Dave Stewart (Aberdeen satellite)'    ; LastLogon = (Get-Date).AddDays(-17)    ; IPAddress = '192.168.224.124' ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },

    ## Dundee Infra
    @{ Name = 'EXASURDUN001' ; SamAccountName = 'EXASURDUN001$' ; Type = 'Computer'      ; Role = 'SUR' ; Site = 'DUN' ; Location = 'Dundee, UK'                   ; DNSHostName = 'EXASURDUN001.example.net' ; OU = @('Locations','UK','Scotland','Dundee','Laptops')              ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '24H2'                 ; Description = 'Laptop – touring staff'                         ; LastLogon = (Get-Date).AddDays(-1)     ; IPAddress = '192.168.138.51'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASURDUN002' ; SamAccountName = 'EXASURDUN002$' ; Type = 'Computer'      ; Role = 'SUR' ; Site = 'DUN' ; Location = 'Dundee, UK'                   ; DNSHostName = 'EXASURDUN002.example.net' ; OU = @('Locations','UK','Scotland','Dundee','Laptops')              ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '24H2'                 ; Description = 'Laptop – touring staff'                         ; LastLogon = (Get-Date).AddDays(-2)     ; IPAddress = '192.168.138.52'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXAPHNDUN001' ; SamAccountName = 'EXAPHNDUN001$' ; Type = 'Phone'         ; Role = 'PHN' ; Site = 'DUN' ; Location = 'Dundee, UK'                   ; DNSHostName = 'EXAPHNDUN001.example.net' ; OU = @('Locations','UK','Scotland','Dundee','Phones')               ; OS = 'iOS 17.2'                          ; OperatingSystemVersion = '17.2'                 ; Description = 'iPhone – touring staff'                         ; LastLogon = 'N/A'                      ; IPAddress = $null             ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXAPHNDUN002' ; SamAccountName = 'EXAPHNDUN002$' ; Type = 'Phone'         ; Role = 'PHN' ; Site = 'DUN' ; Location = 'Dundee, UK'                   ; DNSHostName = 'EXAPHNDUN002.example.net' ; OU = @('Locations','UK','Scotland','Dundee','Phones')               ; OS = 'iOS 17.2'                          ; OperatingSystemVersion = '17.2'                 ; Description = 'iPhone – touring staff'                         ; LastLogon = 'N/A'                      ; IPAddress = $null             ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

    ## Perth, Scotland
    @{ Name = 'EXAPRNPER001' ; SamAccountName = 'EXAPRNPER001$' ; Type = 'Hardware'      ; Role = 'PRN' ; Site = 'PER' ; Location = 'Perth, UK'                    ; DNSHostName = 'EXAPRNPER001.example.net' ; OU = @('Locations','UK','Scotland','Perth','Infrastructure')        ; OS = 'HP MFP LaserJet'                   ; OperatingSystemVersion = 'N/A'                  ; Description = 'Multi-function printer / scanner'               ; LastLogon = 'N/A'                      ; IPAddress = '192.168.173.20'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXASBCPER001' ; SamAccountName = 'EXASBCPER001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'PER' ; Location = 'Perth, UK'                    ; DNSHostName = 'EXASBCPER001.example.net' ; OU = @('Locations','UK','Scotland','Perth','Servers')               ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – VOIP gateway for Perth'               ; LastLogon = 'N/A'                      ; IPAddress = '192.168.173.30'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXANIXPER001' ; SamAccountName = 'EXANIXPER001$' ; Type = 'Unix/Linux'    ; Role = 'NIX' ; Site = 'PER' ; Location = 'Perth, UK'                    ; DNSHostName = 'EXANIXPER001.example.net' ; OU = @('Locations','UK','Scotland','Perth','Servers')               ; OS = 'Solaris 11.5'                      ; OperatingSystemVersion = '11.5'                 ; Description = 'MIDI/Music archive server for Fiction Factory'  ; LastLogon = 'N/A'                      ; IPAddress = '192.168.173.40'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXANASPER001' ; SamAccountName = 'EXANASPER001$' ; Type = 'Hardware'      ; Role = 'NAS' ; Site = 'PER' ; Location = 'Perth, UK'                    ; DNSHostName = 'EXANASPER001.example.net' ; OU = @('Locations','UK','Scotland','Perth','Servers')               ; OS = 'Synology DSM 7.1'                  ; OperatingSystemVersion = '7.1'                  ; Description = 'File storage for user profiles & music archive' ; LastLogon = 'N/A'                      ; IPAddress = '192.168.173.50'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXACVMPER001' ; SamAccountName = 'EXACVMPER001$' ; Type = 'Hardware'      ; Role = 'DON' ; Site = 'PER' ; Location = 'Perth, UK Break Room'         ; DNSHostName = 'EXACVMPER001.example.net' ; OU = @('Locations','UK','Scotland','Perth','Infrastructure')        ; OS = 'Smart Bakery Vending Machine'      ; OperatingSystemVersion = 'SP100'                ; Description = 'Scone Palace vending machine (Easter egg)'      ; LastLogon = 'N/A'                      ; IPAddress = '192.168.173.60'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXAMBPPER001' ; SamAccountName = 'EXAMBPPER001$' ; Type = 'Computer'      ; Role = 'MBP' ; Site = 'PER' ; Location = 'Perth, UK'                    ; DNSHostName = 'EXAMBPPER001.example.net' ; OU = @('Locations','UK','Scotland','Perth','Computers')             ; OS = 'macOS Ventura'                     ; OperatingSystemVersion = '13.5'                 ; Description = 'MacBook Pro assigned to staff'                  ; LastLogon = (Get-Date).AddDays(-2)     ; IPAddress = '192.168.173.70'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'admin'         ; 'msLAPS-Password' = 'Per#MBP!01'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(30).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-10) ; LAPSPasswordExpiration = (Get-Date).AddDays(20) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
    @{ Name = 'EXASURPER001' ; SamAccountName = 'EXASURPER001$' ; Type = 'Computer'      ; Role = 'SUR' ; Site = 'PER' ; Location = 'Perth, UK'                    ; DNSHostName = 'EXASURPER001.example.net' ; OU = @('Locations','UK','Scotland','Perth','Computers')             ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '24H2'                 ; Description = 'Surface Laptop assigned to staff'               ; LastLogon = (Get-Date).AddDays(-3)     ; IPAddress = '192.168.173.71'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Per#SUR!01'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(25).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-5)  ; LAPSPasswordExpiration = (Get-Date).AddDays(25) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
    @{ Name = 'EXAPHNPER001' ; SamAccountName = 'EXAPHNPER001$' ; Type = 'VOIP'          ; Role = 'PHN' ; Site = 'PER' ; Location = 'Perth, UK'                    ; DNSHostName = 'EXAPHNPER001.example.net' ; OU = @('Locations','UK','Scotland','Perth','Phones')                ; OS = 'Yealink T46G'                      ; OperatingSystemVersion = 'N/A'                  ; Description = 'Desk VOIP phone'                                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.173.80'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  }
    @{ Name = 'EXAPHNPER002' ; SamAccountName = 'EXAPHNPER002$' ; Type = 'VOIP'          ; Role = 'PHN' ; Site = 'PER' ; Location = 'Perth, UK'                    ; DNSHostName = 'EXAPHNPER001.example.net' ; OU = @('Locations','UK','Scotland','Perth','Phones')                ; OS = 'Yealink T46G'                      ; OperatingSystemVersion = 'N/A'                  ; Description = 'Desk VOIP phone'                                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.173.80'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  }
    @{ Name = 'EXAPHNPER003' ; SamAccountName = 'EXAPHNPER003$' ; Type = 'VOIP'          ; Role = 'PHN' ; Site = 'PER' ; Location = 'Perth, UK'                    ; DNSHostName = 'EXAPHNPER001.example.net' ; OU = @('Locations','UK','Scotland','Perth','Phones')                ; OS = 'Yealink T46G'                      ; OperatingSystemVersion = 'N/A'                  ; Description = 'Desk VOIP phone'                                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.173.80'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  }
    @{ Name = 'EXAPHNPER004' ; SamAccountName = 'EXAPHNPER004$' ; Type = 'VOIP'          ; Role = 'PHN' ; Site = 'PER' ; Location = 'Perth, UK'                    ; DNSHostName = 'EXAPHNPER001.example.net' ; OU = @('Locations','UK','Scotland','Perth','Phones')                ; OS = 'Yealink T46G'                      ; OperatingSystemVersion = 'N/A'                  ; Description = 'Desk VOIP phone'                                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.173.80'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  }

    ## Edinburgh, Scotland
    @{ Name = 'EXAWKSEDI001' ; SamAccountName = 'EXAWKSEDI001$' ; Type = 'Computer'      ; Role = 'WKS' ; Site = 'EDI' ; Location = 'Edinburgh, UK'                ; DNSHostName = 'EXAWKSEDI001.example.org' ; OU = @('Locations','UK','Scotland','Edinburgh','Computers')         ; OS = 'Windows 10 Pro'                    ; OperatingSystemVersion = '22H2'                 ; Description = 'Shared desktop workstation'                     ; LastLogon = (Get-Date).AddDays(-17)    ; IPAddress = '192.168.131.150' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Edi#Wks!01'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(3).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-27) ; LAPSPasswordExpiration = (Get-Date).AddDays(33) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
    @{ Name = 'EXALAPEDI098' ; SamAccountName = 'EXALAPEDI098$' ; Type = 'Computer'      ; Role = 'LAP' ; Site = 'EDI' ; Location = 'Edinburgh, UK'                ; DNSHostName = 'EXALAPEDI098.example.com' ; OU = @('Locations','UK','Scotland','Edinburgh','Computers')         ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '24H2'                 ; Description = 'Pool laptop – Edinburgh Office'                 ; LastLogon = (Get-Date).AddDays(-3)     ; IPAddress = '192.168.131.108' ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Edi#Soon!98'    ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(1).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-25) ; LAPSPasswordExpiration = (Get-Date).AddDays(55) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
    @{ Name = 'EXARACEDI001' ; SamAccountName = 'EXARACEDI001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'EDI' ; Location = 'Edinburgh, UK'                ; DNSHostName = 'EXARACEDI001.example.net' ; OU = @('Locations','UK','Scotland','Edinburgh','Infrastructure')    ; OS = 'Dell iDRAC9'                       ; OperatingSystemVersion = 'N/A'                  ; Description = 'iDRAC – root / calvin'                          ; LastLogon = 'N/A'                      ; IPAddress = '192.168.131.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXASWIEDI001' ; SamAccountName = 'EXASWIEDI001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'EDI' ; Location = 'Edinburgh, UK'                ; DNSHostName = 'EXASWIEDI001.example.net' ; OU = @('Locations','UK','Scotland','Edinburgh','Network')           ; OS = 'Cisco Catalyst 2960X'              ; OperatingSystemVersion = 'N/A'                  ; Description = 'Floor switch – admin:cisco / EDI2960!'          ; LastLogon = 'N/A'                      ; IPAddress = '192.168.131.5'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXASWIEDI002' ; SamAccountName = 'EXASWIEDI001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'EDI' ; Location = 'Edinburgh, UK'                ; DNSHostName = 'EXASWIEDI002.example.net' ; OU = @('Locations','UK','Scotland','Edinburgh','Network')           ; OS = 'Cisco Catalyst 2960X'              ; OperatingSystemVersion = 'N/A'                  ; Description = 'Cisco Catalyst 48-port - admin:cisco123'        ; LastLogon = 'N/A'                      ; IPAddress = '192.168.131.6'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXASBCEDI001' ; SamAccountName = 'EXASBCEDI001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'EDI' ; Location = 'Edinburgh, UK'                ; DNSHostName = 'EXASBCEDI001.example.net' ; OU = @('Locations','UK','Scotland','Edinburgh','Servers')           ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-EDI-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.131.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    ## Internet Connected cofee pot https://en.wikipedia.org/wiki/Trojan_Room_coffee_pot?useskin=vector
    @{ Name = 'EXATEAEDI001' ; SamAccountName = 'EXATEAEDI001$' ; Type = 'Siemens EQ700' ; Role = 'TEA' ; Site = 'EDI' ; Location = 'Edinburgh Office Break Room'  ; DNSHostName = 'EXATEAEDI001.example.com' ; OU = @('Locations','UK','Scotland','Edinburgh','Computers')         ; OS = 'Smart Bean to Cup Coffee Machine'  ; OperatingSystemVersion = 'TP713GB9'             ; Description = 'Coffee Machine in Edinburgh office kitchen'     ; LastLogon = 'N/A'                      ; IPAddress = '192.168.131.16'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },

    ## From Miami to Canada B-)
    @{ Name = 'EXALAPBRK001' ; SamAccountName = 'EXALAPBRK001$' ; Type='Computer'        ; Role='LAP'   ; Site = 'BRK'  ; Location = 'Brockille, ON, CA'           ; DNSHostName = 'EXALAPBRK001.example.net' ; OU = @('Locations','CA','Ontario','Brockville','The Proclaimers')   ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '24H2'                 ; Description = 'Tour laptop (Canada)'                           ; LastLogon = (Get-Date).AddDays(-11)    ; IPAddress = '10.20.10.21'     ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mcr#Lap@18'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(7).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-15) ; LAPSPasswordExpiration = (Get-Date).AddDays(15) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
    @{ Name = 'EXALAPMIA001' ; SamAccountName = 'EXALAPMIA001$' ; Type='Computer'        ; Role='LAP'   ; Site = 'MIA'  ; Location = 'Miami, FL, US'               ; DNSHostName = 'EXALAPMIA001.example.net' ; OU = @('Locations','US','Florida','Miami','The Proclaimers'     )   ; OS = 'macOS Sonoma'                      ; OperatingSystemVersion = '25.2'                 ; Description = 'Tour MacBook (Miami)'                           ; LastLogon = (Get-Date).AddDays(-14)    ; IPAddress = '10.30.10.21'     ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mcr#Lap@18'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(7).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-15) ; LAPSPasswordExpiration = (Get-Date).AddDays(15) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },

    ## Glasgow
    @{ Name = 'EXAWKSGLA001' ; SamAccountName = 'EXAWKSGLA001$' ; Type = 'Workstation'   ; Role = 'WKS' ; Site = 'GLA' ; Location = 'Glasgow, Scotland'            ; DNSHostName = 'EXAWKSGLA001.example.com' ; OU = @('Locations','UK','Scotland','Glasgow','Computers')           ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Hot desk workstation - Glasgow office'          ; LastLogon = (Get-Date).AddHours(-2)    ; IPAddress = '192.168.141.150' ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Gla#Wks!41'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(9).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-12) ; LAPSPasswordExpiration = (Get-Date).AddDays(9)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
    @{ Name = 'EXAWKSGLA002' ; SamAccountName = 'EXAWKSGLA002$' ; Type = 'Workstation'   ; Role = 'WKS' ; Site = 'GLA' ; Location = 'Glasgow, Scotland'            ; DNSHostName = 'EXAWKSGLA002.example.com' ; OU = @('Locations','UK','Scotland','Glasgow','Computers')           ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Hot desk workstation - Glasgow office'          ; LastLogon = (Get-Date).AddDays(-1)     ; IPAddress = '192.168.141.151' ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Gla#Wks!73'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(14).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-8)  ; LAPSPasswordExpiration = (Get-Date).AddDays(14) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
    @{ Name = 'EXALAPGLA001' ; SamAccountName = 'EXALAPGLA001$' ; Type = 'Laptop'        ; Role = 'LAP' ; Site = 'GLA' ; Location = 'Glasgow, Scotland'            ; DNSHostName = 'EXALAPGLA001.example.com' ; OU = @('Locations','UK','Scotland','Glasgow','Computers')           ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Dell Latitude laptop - Pool device'             ; LastLogon = (Get-Date).AddHours(-5)    ; IPAddress = '192.168.141.152' ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Gla#Lap@19'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(6).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-10) ; LAPSPasswordExpiration = (Get-Date).AddDays(6)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
    @{ Name = 'EXAPRNGLA001' ; SamAccountName = 'EXAPRNGLA001$' ; Type = 'Printer'       ; Role = 'PRN' ; Site = 'GLA' ; Location = 'Glasgow, Scotland'            ; DNSHostName = 'EXAPRNGLA001.example.com' ; OU = @('Locations','UK','Scotland','Glasgow','Computers')           ; OS = 'Printer'                           ; OperatingSystemVersion = 'N/A'                  ; Description = 'HP LaserJet Pro - Main floor'                   ; LastLogon = (Get-Date).AddMinutes(-30) ; IPAddress = '192.168.141.16'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },

    ## Clydebank Infra
    @{ Name = 'EXASWICLY001' ; SamAccountName = 'EXASWICLY001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXASWICLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Network')           ; OS = 'Cisco Catalyst 9300'               ; OperatingSystemVersion = 'N/A'                  ; Description = 'Cisco Catalyst 48-port (admin:catalyst80s)'     ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXAFWLCLY001' ; SamAccountName = 'EXAFWLCLY001$' ; Type = 'Router'        ; Role = 'RTR' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXAFWLCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Network')           ; OS = 'FortiOS'                           ; OperatingSystemVersion = '7.6.5 build 3651'     ; Description = 'Firewall / VPN Gateway'                         ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.1'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASRVCLY001' ; SamAccountName = 'EXASRVCLY001$' ; Type = 'Server'        ; Role = 'SRV' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXASRVCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Servers')           ; OS = 'Rocky Linux'                       ; OperatingSystemVersion = 'N/A'                  ; Description = 'Database Server running Oracle DB'              ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.20'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXAILOCLY001' ; SamAccountName = 'EXAILOCLY001$' ; Type = 'RAC'           ; Role = 'ILO' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXARACCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Infrastructure')    ; OS = 'HPE iLO5'                          ; OperatingSystemVersion = 'N/A'                  ; Description = 'HP iLO in Clydebank (Administrator:@nG3l3yE$)'  ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.30'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASBCCLY001' ; SamAccountName = 'EXASBCCLY001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXASBCCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Servers')           ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-CLY-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASURCLY001' ; SamAccountName = 'EXASURCLY001$' ; Type = 'Computer'      ; Role = 'LAP' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXASURCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Wet Wet Wet')       ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '24H2'                 ; Description = 'Microsoft Surface – touring'                    ; LastLogon = (Get-Date).AddDays(-2)     ; IPAddress = '192.168.141.51'  ; Domain = 'example.net' ; 'msLAPS-AccountName' ='Administrator'  ; 'msLAPS-Password' ='Cly%SUR@07'      ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(22)                 ; LAPSPasswordLastSet = (Get-Date).AddDays(8)   ; LAPSPasswordExpiration = (Get-Date).AddDays(30) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
    @{ Name = 'EXAPHNGLA001' ; SamAccountName = 'EXAPHNCLY001$' ; Type = 'Phone'         ; Role = 'PHN' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXAPHNCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Phones') ;          ; OS = 'Apple iOS'                         ; OperatingSystemVersion = '26.1'                 ; Description = 'Android phone – touring handset'                ; LastLogon = 'N/A'                      ; IPAddress =  $null            ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXATABGLA001' ; SamAccountName = 'EXATABCLY001$' ; Type = 'Tablet'        ; Role = 'TAB' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXATABCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Computer')          ; OS = 'Apple iOD'                         ; OperatingSystemVersion = '26.2'                 ; Description = 'Android tablet – setlists (service account)'    ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.62'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASWICLY001' ; SamAccountName = 'EXASWICLY001$' ; Type = 'Switch'        ; Role = 'SWI' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXASWICLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','Network')           ; OS = 'Cisco Catalyst 9300'               ; OperatingSystemVersion = 'N/A'                  ; Description = 'TPLink 48port managed switch; admin:tplink1987' ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.20'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXARACCLY001' ; SamAccountName = 'EXARACCLY001$' ; Type = 'RAC'           ; Role = 'RAC' ; Site = 'CLY' ; Location = 'Clydebank, UK'                ; DNSHostName = 'EXARACCLY001.example.net' ; OU = @('Locations','UK','Scotland','Clydebank','BMC')               ; OS = 'HP iLO5'                           ; OperatingSystemVersion = 'N/A'                  ; Description = 'Dell DRAC for EXADCCLY001; dracadmin:WetWet87'  ; LastLogon = 'N/A'                      ; IPAddress = '192.168.141.30'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

    ## Newcastle, UK
    @{ Name = 'EXASRVNEW001' ; SamAccountName = 'EXASRVNEW001$' ; Type = 'Server'        ; Role = 'SRV' ; Site = 'NEW' ; Location = 'Newcastle, UK'                ; DNSHostName = 'EXASRVNEW001.example.org' ; OU = @('Locations','UK','England','Newcastle','Computers')          ; OS = 'Windows Server 2022'               ; OperatingSystemVersion = '21H2'                 ; Description = 'File and print server for Newcastle office'     ; LastLogon = (Get-Date).AddDays(-5)     ; IPAddress = '192.168.191.21'  ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'New#Srv!01'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(30).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-21) ; LAPSPasswordExpiration = (Get-Date).AddDays(13) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
    @{ Name = 'EXARACNEW001' ; SamAccountName = 'EXARACNEW001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'NEW' ; Location = 'Newcastle, UK'                ; DNSHostName = 'EXARACNEW001.example.net' ; OU = @('Locations','UK','England','Newcastle','Infrastructure')     ; OS = 'Dell iDRAC9'                       ; OperatingSystemVersion = 'N/A'                  ; Description = 'iDRAC – root / new-rac-01!'                     ; LastLogon = 'N/A'                      ; IPAddress = '192.168.191.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASWINEW001' ; SamAccountName = 'EXASWINEW001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'NEW' ; Location = 'Newcastle, UK'                ; DNSHostName = 'EXASWINEW001.example.net' ; OU = @('Locations','UK','England','Newcastle','Network')            ; OS = 'TP-Link JetStream'                 ; OperatingSystemVersion = 'N/A'                  ; Description = 'Access switch – admin / NEWsw!'                 ; LastLogon = 'N/A'                      ; IPAddress = '192.168.191.5'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASBCNEW001' ; SamAccountName = 'EXASBCNEW001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'NEW' ; Location = 'Newcastle, UK'                ; DNSHostName = 'EXASBCNEW001.example.net' ; OU = @('Locations','UK','England','Newcastle','Servers')            ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-NEW-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.191.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXAWKSNEW099' ; SamAccountName = 'EXAWKSNEW099$' ; Type = 'Computer'      ; Role = 'WKS' ; Site = 'NEW' ; Location = 'Newcastle, UK'                ; DNSHostName = 'EXAWKSNEW099.example.org' ; OU = @('Locations','UK','England','Newcastle','Computers')          ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '24H2'                 ; Description = 'Newcastle office PC'                            ; LastLogon = (Get-Date).AddDays(-3)     ; IPAddress = '192.168.191.161' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'New#Expired!99' ; 'msLAPS-PasswordExpirationTime'=(Get-Date).AddDays(-1).ToFileTimeUtc()   ; LAPSPasswordLastSet = (Get-Date).AddDays(-31) ; LAPSPasswordExpiration = (Get-Date).AddDays(-1) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },

    ## Manchester, UK
    @{ Name = 'EXALAPMCR001' ; SamAccountName = 'EXALAPMCR001$' ; Type = 'Laptop'        ; Role = 'LAP' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXALAPMCR001.example.org' ; OU = @('Locations','UK','England','Manchester','Computers')         ; OS = 'Windows 11 Enterprise'             ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Management laptop – Manchester office'          ; LastLogon = (Get-Date).AddDays(-12)    ; IPAddress = '192.168.228.19'  ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mcr#Lap@18'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(7).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-20) ; LAPSPasswordExpiration = (Get-Date).AddDays(7)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
    @{ Name = 'EXALAPMCR002' ; SamAccountName = 'EXALAPMCR002$' ; Type = 'Laptop'        ; Role = 'LAP' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXALAPMCR002.example.org' ; OU = @('Locations','UK','England','Manchester','Computers')         ; OS = 'Windows 11 Enterprise'             ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Mobile staff laptop – Manchester office'        ; LastLogon = (Get-Date).AddDays(-24)    ; IPAddress = '192.168.161.150' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mcr#Lap@92'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(4).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-26) ; LAPSPasswordExpiration = (Get-Date).AddDays(4)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
    @{ Name = 'EXAWKSMCR001' ; SamAccountName = 'EXAWKSMCR001$' ; Type = 'Computer'      ; Role = 'WKS' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXAWKSMCR001.example.org' ; OU = @('Locations','UK','England','Manchester','Computers')         ; OS = 'Windows 10 Enterprise'             ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Front desk workstation – Manchester office'     ; LastLogon = (Get-Date).AddDays(-40)    ; IPAddress = '192.168.161.152' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mcr#Wks!09'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(12).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-19) ; LAPSPasswordExpiration = (Get-Date).AddDays(12) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
    @{ Name = 'EXAWKSMCR002' ; SamAccountName = 'EXAWKSMCR002$' ; Type = 'Computer'      ; Role = 'WKS' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXAWKSMCR002.example.org' ; OU = @('Locations','UK','England','Manchester','Computers')         ; OS = 'Windows 10 Enterprise'             ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Finance workstation – Manchester office'        ; LastLogon = (Get-Date).AddDays(-9)     ; IPAddress = '192.168.161.153' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mcr#Wks!55'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(8).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-22) ; LAPSPasswordExpiration = (Get-Date).AddDays(8)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
    @{ Name = 'EXAPRNMCR001' ; SamAccountName = 'EXAPRNMCR001$' ; Type = 'Printer'       ; Role = 'PRN' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXAPRNMCR001.example.org' ; OU = @('Locations','UK','England','Manchester','Computers')         ; OS = 'Embedded Printer Firmware'         ; OperatingSystemVersion = 'N/A'                  ; Description = 'Main network printer – Manchester office'       ; LastLogon = (Get-Date).AddDays(-4)     ; IPAddress = '192.168.161.16'  ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXARACMCR001' ; SamAccountName = 'EXARACMCR001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXARACMCR001.example.net' ; OU = @('Locations','UK','England','Manchester','Infrastructure')    ; OS = 'HPE iLO5'                          ; OperatingSystemVersion = 'N/A'                  ; Description = 'iLO – Administrator / mcr-ilo-01'               ; LastLogon = 'N/A'                      ; IPAddress = '192.168.161.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASWIMCR001' ; SamAccountName = 'EXASWIMCR001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXASWIMCR001.example.net' ; OU = @('Locations','UK','England','Manchester','Network')           ; OS = 'Cisco Catalyst 9300'               ; OperatingSystemVersion = 'N/A'                  ; Description = 'Distribution switch – admin:cisco / MCR9300!'   ; LastLogon = 'N/A'                      ; IPAddress = '192.168.161.5'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASBCMCR001' ; SamAccountName = 'EXASBCMCR001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'MCR' ; Location = 'Manchester, UK'               ; DNSHostName = 'EXASBCMCR001.example.net' ; OU = @('Locations','UK','England','Manchester','Servers')           ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-MCR-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.161.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  }

    ## Liverpool, UK
    @{ Name = 'EXASVRLIV001' ; SamAccountName = 'EXASVRLIV001$' ; Type = 'Server'        ; Role = 'SVR' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXASVRLIV001.example.org' ; OU = @('Locations','UK','England','Liverpool','Computers')          ; OS = 'Windows Server 2022'               ; OperatingSystemVersion = '21H2'                 ; Description = 'File Server for Liverpool site'                 ; LastLogon = (Get-Date).AddDays(-8)     ; IPAddress = '192.168.151.10'  ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = '$3s$4m3BuN*'    ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(30).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-1)  ; LAPSPasswordExpiration = (Get-Date).AddDays(30) ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXAMBPLIV001' ; SamAccountName = 'EXAMBPLIV001$' ; Type = 'Macbook'       ; Role = 'MBP' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXAMBPLIV001.example.org' ; OU = @('Locations','UK','England','Liverpool','Computers')          ; OS = 'MacOS Tahoe'                       ; OperatingSystemVersion = '26.0.1'               ; Description = 'Macbook Pro 2024 - Pool device'                 ; LastLogon = (Get-Date).AddDays(-14)    ; IPAddress = '192.168.151.150' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Liv#Mac@31'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(13).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-21) ; LAPSPasswordExpiration = (Get-Date).AddDays(13) ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  }, ## If TPM is disabled we don't have the other stuff
    @{ Name = 'EXAMACLIV001' ; SamAccountName = 'EXAMACLIV001$' ; Type = 'iMac'          ; Role = 'MAC' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXAMACLIV001.example.org' ; OU = @('Locations','UK','England','Newcastle','Computers')          ; OS = 'MacOS Tahoe'                       ; OperatingSystemVersion = '26.0.1'               ; Description = 'Macbook Pro 2024 - Pool device'                 ; LastLogon = (Get-Date).AddDays(-40)    ; IPAddress = '192.168.151.152' ; Domain = 'example.org' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Liv#Mac@77'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(17).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-35) ; LAPSPasswordExpiration = (Get-Date).AddDays(17) ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $false }, ## Disabled for maintenance
    @{ Name = 'EXARDRLIV002' ; SamAccountName = 'EXARDRLIV002$' ; Type = 'Security'      ; Role = 'RDR' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXARDRLIV002.example.org' ; OU = @('Locations','UK','England','Liverpool','Computers')          ; OS = 'HID Signo Reader'                  ; OperatingSystemVersion = 'HID-SIGNO-LIV-338572' ; Description = 'Electronic badge reader for access control'     ; LastLogon = (Get-Date).AddDays(-9)     ; IPAddress = '192.168.31.16'   ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXABPSLIV001' ; SamAccountName = 'EXABPSLIV001$' ; Type = 'Security'      ; Role = 'BPS' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXABPSLIV001.example.org' ; OU = @('Locations','UK','England','Liverpool','Computers')          ; OS = 'HID Omnikey 5427'                  ; OperatingSystemVersion = 'OMNI5427-LIV-449821'  ; Description = 'Badge programming workstation'                  ; LastLogon = (Get-Date).AddDays(-23)    ; IPAddress = '192.168.151.17'  ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXARACLIV001' ; SamAccountName = 'EXARACLIV001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXARACLIV001.example.net' ; OU = @('Locations','UK','England','Liverpool','Infrastructure')     ; OS = 'HPE iLO5'                          ; OperatingSystemVersion = 'N/A'                  ; Description = 'iLO – admin / liv-ilo-pass'                     ; LastLogon = 'N/A'                      ; IPAddress = '192.168.151.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASWILIV001' ; SamAccountName = 'EXASWILIV001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXASWILIV001.example.net' ; OU = @('Locations','UK','England','Liverpool','Network')            ; OS = 'Cisco Catalyst 9200'               ; OperatingSystemVersion = 'N/A'                  ; Description = 'Core switch – admin:cisco / LIVcore!'           ; LastLogon = 'N/A'                      ; IPAddress = '192.168.151.5'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASBCLIV001' ; SamAccountName = 'EXASBCLIV001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'LIV' ; Location = 'Liverpool, UK'                ; DNSHostName = 'EXASBCLIV001.example.net' ; OU = @('Locations','UK','England','Liverpool','Servers')            ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-LIV-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.151.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

    ## Birmingham, UK
    @{ Name = 'EXASWIBIR001' ; SamAccountName = 'EXASWIBIR001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXASWIBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Network')           ; OS = 'Cisco Catalyst 9300'               ; OperatingSystemVersion = 'N/A'                  ; Description = 'Core Switch'                                    ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXAFWLBIR001' ; SamAccountName = 'EXAFWLBIR001$' ; Type = 'Router'        ; Role = 'RTR' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXAFWLBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Network')           ; OS = 'Palo Alto PanOS'                   ; OperatingSystemVersion = 'N/A'                  ; Description = 'Palo Alto Site Firewall / VPN Gateway'          ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.1'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASRVBIR001' ; SamAccountName = 'EXASRVBIR001$' ; Type = 'Server'        ; Role = 'SRV' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXASRVBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Servers')           ; OS = 'Rocky Linux'                       ; OperatingSystemVersion = 'N/A'                  ; Description = 'Roxy Linux Node runnning Oracle DB'             ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.20'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXAILOBIR001' ; SamAccountName = 'EXAILOBIR001$' ; Type = 'RAC'           ; Role = 'ILO' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXARACBIR001.example.net' ; OU = @('Locations','UK','England','Manchester','Infrastructure')    ; OS = 'HPE iLO5'                          ; OperatingSystemVersion = 'N/A'                  ; Description = 'HP iLO in Birmingham (Administrator:Ay3L0w@)'   ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.30'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASBCBIR001' ; SamAccountName = 'EXASBCBIR001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXASBCBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Servers')           ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = '3CX Debian'           ; Description = '3CX SBC – ssh:root / 3cx-BIR-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXAMBPBIR001' ; SamAccountName = 'EXAMBPBIR001$' ; Type = 'Computer'      ; Role = 'MBP' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXAMBPBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Computers')         ; OS = 'macOS Sonoma'                      ; OperatingSystemVersion = '25.2.1'               ; Description = 'MacBook Pro – touring'                          ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.41'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXATABBIR001' ; SamAccountName = 'EXATABBIR001$' ; Type = 'Tablet'        ; Role = 'TAB' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXATABLIV001.example.net' ; OU = @('Locations','UK','England','Liverpool','Tablets')            ; OS = 'Samsung Galaxy Tab A10'            ; OperatingSystemVersion = 'Lineage OS'           ; Description = 'Android tablet – setlists (service account)'    ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.61'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXAPHNBIR001' ; SamAccountName = 'EXAPHNBIR001$' ; Type = 'Phone'         ; Role = 'PHN' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXAPHNLIV001.example.net' ; OU = @('Locations','UK','England','Liverpool','Phones')             ; OS = 'Samsung Glalaxy S25 Ultra'         ; OperatingSystemVersion = 'Linage OS'            ; Description = 'Android phone – touring handset'                ; LastLogon = 'N/A'                      ; IPAddress = $null             ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASWIBIR002' ; SamAccountName = 'EXASWIBIR002$' ; Type = 'Switch'        ; Role = 'SWI' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXASWIBIR002.example.net' ; OU = @('Locations','UK','England','Birmingham','Network')           ; OS = 'N/A'                               ; OperatingSystemVersion = 'N/A'                  ; Description = 'Cisco Catalyst 48-port (admin:catalyst80s)'     ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.20'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXARACBIR001' ; SamAccountName = 'EXARACBIR001$' ; Type = 'RAC'           ; Role = 'RAC' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXARACBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Infrastructure')    ; OS = 'N/A'                               ; OperatingSystemVersion = 'N/A'                  ; Description = 'Dell DRAC Birmingham   (dracadmin:Dr@c1983)'    ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.6'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXALCDBIR001' ; SamAccountName = 'EXALCDBIR001$' ; Type = 'Wall Display'  ; Role = 'LCD' ; Site = 'BIR' ; Location = 'Birmingham, UK'               ; DNSHostName = 'EXALCDBIR001.example.net' ; OU = @('Locations','UK','England','Birmingham','Displays')          ; OS = 'NEC PlasmaSync 42MP1'              ; OperatingSystemVersion = 'PlasmaSync v1.x'      ; Description = 'LCD status display (NOC / call queue display)'  ; LastLogon = 'N/A'                      ; IPAddress = '192.168.121.30'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

    ## Avocado Central
    @{ Name = 'EXAWKSLND001' ; SamAccountName = 'EXAWKSLND001$' ; Type = 'Workstation'   ; Role = 'WKS' ; Site = 'LND' ; Location = 'London, England'              ; DNSHostName = 'EXAWKSLND001.example.com' ; OU = @('Locations','UK','England','London','Computers')             ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Hot desk workstation - London office'           ; LastLogon = (Get-Date).AddHours(-1)    ; IPAddress = '192.168.20.150'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Bon#Wks!22'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(10).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-14) ; LAPSPasswordExpiration=(Get-Date).AddDays(10)   ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
    @{ Name = 'EXAPRNLND001' ; SamAccountName = 'EXAPRNLND001$' ; Type = 'Printer'       ; Role = 'PRN' ; Site = 'LND' ; Location = 'London, England'              ; DNSHostName = 'EXAPRNLND001.example.com' ; OU = @('Locations','UK','England','London','Computers')             ; OS = 'Printer'                           ; OperatingSystemVersion = 'N/A'                  ; Description = 'Xerox WorkCentre - Reception'                   ; LastLogon = (Get-Date).AddMinutes(-15) ; IPAddress = '192.168.20.16'   ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXARTRLND001' ; SamAccountName = 'EXARTRLND001$' ; Type = 'Network'       ; Role = 'RTR' ; Site = 'LND' ; Location = 'London, UK'                   ; DNSHostName = 'EXARTRLND001.example.com' ; OU = @('Locations','UK','England','London','Computers')             ; OS = 'Cisco ISR 4331'                    ; OperatingSystemVersion = 'ISR4331-LND-552901'   ; Description = 'WAN edge router'                                ; LastLogon = (Get-Date).AddHours(-13)   ; IPAddress = '192.168.20.254'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXAFWLLND001' ; SamAccountName = 'EXAFWLLND001$' ; Type = 'Network'       ; Role = 'FWL' ; Site = 'LND' ; Location = 'London, UK'                   ; DNSHostName = 'EXAFWLLND001.example.com' ; OU = @('Locations','UK','England','London','Computers')             ; OS = 'Cisco ASA 5516-X'                  ; OperatingSystemVersion = 'ASA5516-LND-884210'   ; Description = 'Perimeter firewall and VPN gateway'             ; LastLogon = (Get-Date).AddDays(-5)     ; IPAddress = '192.168.20.1'    ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXARACLND001' ; SamAccountName = 'EXARACLND001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'LND' ; Location = 'London, UK'                   ; DNSHostName = 'EXARACLND001.example.net' ; OU = @('Locations','UK','England','London','Infrastructure')        ; OS ='Dell iDRAC9'                        ; OperatingSystemVersion = 'N/A'                  ; Description = 'Dell PowerEdge iDRAC – root / P@ssLND-RAC01'    ; LastLogon = 'N/A'                      ; IPAddress = '192.168.20.2'    ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXASWILND001' ; SamAccountName = 'EXASWILND001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'LND' ; Location = 'London, UK'                   ; DNSHostName = 'EXASWILND001.example.net' ; OU = @('Locations','UK','England','London','Network')               ; OS ='Cisco Catalyst 9300'                ; OperatingSystemVersion = 'N/A'                  ; Description = 'Core switch – admin:cisco / Sw1tchLND!'         ; LastLogon = 'N/A'                      ; IPAddress = '192.168.20.5'    ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXASBCLND001' ; SamAccountName = 'EXASBCLND001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'LND' ; Location = 'London, UK'                   ; DNSHostName = 'EXASBCLND001.example.net' ; OU = @('Locations','UK','England','London','Servers')               ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-LND-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.20.49'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXAPRNLND002' ; SamAccountName = 'EXAPRNLND002$' ; Type = 'Stenography'   ; Role = 'PRN' ; Site = 'LND' ; Location = 'London, UK'                   ; DNSHostName = 'EXAPRNLND001.example.net' ; OU = @('Locations','UK','England','London','Stenography')           ; OS = 'Embedded Firmware'                 ; OperatingSystemVersion = 'ProCAT Stylus'        ; Description = 'ProCAT Stylus Steno Writer – Court Device'      ; LastLogon = 'N/A'                      ; IPAddress = 'N/A'             ; Domain = 'example.org' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },

    ## Odense, DK
    @{ Name = 'EXAMACODE001' ; SamAccountName = 'EXAMACODE001$' ; Type = 'Computer'      ; Role = 'MAC' ; Site = 'ODE' ; Location = 'Odense, Danmark'              ; DNSHostName = 'EXAMACODE001.example.com' ; OU = @('Locations','Danmark','Odense','Computers')                  ; OS = 'MacOS Tahoe'                       ; OperatingSystemVersion = '26.0.1'               ; Description = 'Design team iMac workstation'                   ; LastLogon = (Get-Date).AddDays(-7)     ; IPAddress = '192.168.66.150'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Ode#Mac@44'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(14).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-21) ; LAPSPasswordExpiration = (Get-Date).AddDays(14) ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXAMBPODE002' ; SamAccountName = 'EXAMBPODE002$' ; Type = 'Computer'      ; Role = 'MBP' ; Site = 'ODE' ; Location = 'Odense, Danmark'              ; DNSHostName = 'EXAMBPODE002.example.com' ; OU = @('Locations','Danmark','Odense','Computers')                  ; OS = 'MacOS Tahoe'                       ; OperatingSystemVersion = '26.0.1'               ; Description = 'Executive MacBook Pro'                          ; LastLogon = (Get-Date).AddDays(-23)    ; IPAddress = '192.168.66.151'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Ode#Mac@81'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(10).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-30) ; LAPSPasswordExpiration = (Get-Date).AddDays(10) ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    ## Jukebox. This will be funny to you, if you have ever stayed at First Hotel Grand Odense Jernbanegade 18, 5000 Odense, Danmark
    @{ Name = 'EXAMUSODE001' ; SamAccountName = 'EXAMUSODE001$' ; Type = 'Jukebox'       ; Role = 'MUS' ; Site = 'ODE' ; Location = 'Odense Office – Ground Floor' ; DNSHostName = 'EXAMUSODE001.example.org' ; OU = @('Locations','Danmark','Odense','Computers')                  ; OS = 'Pureline 128V Retro Vinyl Jukebox' ; OperatingSystemVersion = 'PL128V-ODE-095823'    ; Description = 'Retro Vinyl jukebox in the Odense office.'      ; LastLogon = (Get-Date).AddHours(-4)    ; IPAddress = '192.168.66.18'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

    ## Køge, DK
    @{ Name = 'EXAWAPKGE001' ; SamAccountName = 'EXAWAPKGE001$' ; Type = 'Network'       ; Role = 'WAP' ; Site = 'KGE' ; Location = 'Køge, Danmark'                ; DNSHostName = 'EXAWAPKGE001.example.com' ; OU = @('Locations','Danmark','Køge','Computers')                    ; OS = 'Ubiquiti UniFi U6-Pro'             ; OperatingSystemVersion = 'U6P-KGE-847392'       ; Description = 'Enterprise Wi-Fi access point'                  ; LastLogon = (Get-Date).AddDays(-18)    ; IPAddress = '192.168.56.5'    ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXAPRNKGE001' ; SamAccountName = 'EXAPRNKGE001$' ; Type = 'Peripheral'    ; Role = 'PRN' ; Site = 'KGE' ; Location = 'Køge, Danmark'                ; DNSHostName = 'EXAPRNKGE001.example.com' ; OU = @('Locations','Danmark','Køge','Computers')                    ; OS = 'HP LaserJet Enterprise MFP M528'   ; OperatingSystemVersion = 'HPM528-KGE-193847'    ; Description = 'Networked multifunction printer'                ; LastLogon = (Get-Date).AddDays(-10)    ; IPAddress = '192.168.56.16'   ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

    ## København, DK
    @{ Name = 'EXACLKCPH001' ; SamAccountName = 'EXACLKCPH001$' ; Type = 'IoT'           ; Role = 'CLK' ; Site = 'CPH' ; Location = 'Christianshavn, København'    ; DNSHostName = 'EXACLKCPH001.example.com' ; OU = @('Locations','Danmark','Copenhagen','Computers')              ; OS = 'Meinberg LANTIME M300'             ; OperatingSystemVersion = 'M300-CPH-661204'      ; Description = 'Network-synchronised NTP clock'                 ; LastLogon = (Get-Date).AddDays(-4)     ; IPAddress = '192.168.228.18'  ; Domain = 'example.com' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXARACCPH001' ; SamAccountName = 'EXARACCPH001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'CPH' ; Location = 'Christianshavn, København'    ; DNSHostName = 'EXARACCPH001.example.net' ; OU = @('Locations','Danmark','Copenhagen','Infrastructure')         ; OS = 'Dell iDRAC9'                       ; OperatingSystemVersion = 'N/A'                  ; Description = 'iDRAC – root / CPH-rac01!'                      ; LastLogon = 'N/A'                      ; IPAddress = '192.168.31.2'    ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASWICPH001' ; SamAccountName = 'EXASWICPH001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'CPH' ; Location = 'Christianshavn, København'    ; DNSHostName = 'EXASWICPH001.example.net' ; OU = @('Locations','Danmark','Copenhagen','Network')                ; OS = 'TP-Link JetStream'                 ; OperatingSystemVersion = 'N/A'                  ; Description = 'Office switch – admin / CPHsw!'                 ; LastLogon = 'N/A'                      ; IPAddress = '192.168.31.5'    ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASBCCPH001' ; SamAccountName = 'EXASBCCPH001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'CPH' ; Location = 'Christianshavn, København'    ; DNSHostName = 'EXASBCCPH001.example.net' ; OU = @('Locations','Danmark','Copenhagen','Servers')                ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-CPH-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.31.49'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXATVSCPH001' ; SamAccountName = 'EXATVSCPH001$' ; Type = 'TVs'           ; Role = 'TVS' ; Site = 'CPH' ; Location = 'København Centrum office'     ; DNSHostName = 'EXATVSCPH001.example.com' ; OU = @('Locations','Danmark','Copenhagen','Computers')              ; OS = 'Bella Kronik 42X'                  ; OperatingSystemVersion = 'BTV-042001'           ; Description = 'Bella TV streams DR/TV2 in the CPH office..'    ; LastLogon = 'N/A'                      ; IPAddress = '192.168.31.17'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

    ## West Berlin, DE
    @{ Name = 'EXASRVBRD001' ; SamAccountName = 'EXASRVBRD001$' ; Type = 'Server'        ; Role = 'SRV' ; Site = 'BRD' ; Location = 'West Berlin, FRG (DE)'        ; DNSHostName = 'EXASRVBRD001.example.net' ; OU = @('Locations','Germany','West Berlin','Computers')             ; OS = 'Windows Server 2019'               ; OperatingSystemVersion = '1809'                 ; Description = 'Legacy application server (West Berlin site)'   ; LastLogon = (Get-Date).AddDays(-1)     ; IPAddress = '192.168.30.21'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Ged#Srv!19'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(30).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-25) ; LAPSPasswordExpiration = (Get-Date).AddDays(30) ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXANIXBRD002' ; SamAccountName = 'EXANIXBRD002$' ; Type = 'Unix'          ; Role = 'NIX' ; Site = 'BRD' ; Location = 'West Berlin, FRG (DE)'        ; DNSHostName = 'EXANIXBRD002.example.net' ; OU = @('Locations','Germany','West Berlin','Computers')             ; OS = 'Debian 12'                         ; OperatingSystemVersion = '12.2'                 ; Description = 'Linux server hosting internal services'         ; LastLogon = (Get-Date).AddDays(-5)     ; IPAddress = '192.168.30.22'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

    ## Munich, DE
    @{ Name = 'EXAWKSMUN001' ; SamAccountName = 'EXAWKSMUN001$' ; Type = 'Workstation'   ; Role = 'WKS' ; Site = 'BON' ; Location = 'Munich, Germany'              ; DNSHostName = 'EXAWKSMUN001.example.net' ; OU = @('Locations','Germany','Munich','Computers')                  ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Hot desk workstation - Bonn office'             ; LastLogon = (Get-Date).AddHours(-3)    ; IPAddress = '192.168.89.150'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mun#Wks!41'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(11).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-19) ; LAPSPasswordExpiration = (Get-Date).AddDays(11) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
    @{ Name = 'EXALAPMUN001' ; SamAccountName = 'EXALAPMUN001$' ; Type = 'Laptop'        ; Role = 'LAP' ; Site = 'MUN' ; Location = 'Munich, Germany'              ; DNSHostName = 'EXALAPMUN001.example.net' ; OU = @('Locations','Germany','Munich','Computers')                  ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Pool laptop – Munich office'                    ; LastLogon = (Get-Date).AddDays(-3)     ; IPAddress = '192.168.89.151'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mun#Lap!47'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(2).ToFileTimeUtc()  ; LAPSPasswordLastSet=(Get-Date).AddDays(-28)   ; LAPSPasswordExpiration = (Get-Date).AddDays(2)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
    @{ Name = 'EXALAPMUN002' ; SamAccountName = 'EXALAPMUN002$' ; Type = 'Laptop'        ; Role = 'LAP' ; Site = 'MUN' ; Location = 'Munich, Germany'              ; DNSHostName = 'EXALAPMUN002.example.net' ; OU = @('Locations','Germany','Munich','Computers')                  ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Pool laptop – Munich office'                    ; LastLogon = (Get-Date).AddDays(-95)    ; IPAddress = '192.168.89.152'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Mun#Lap!02'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(-61).ToFileTimeUtc(); LAPSPasswordLastSet=(Get-Date).AddDays(-91)   ; LAPSPasswordExpiration = (Get-Date).AddDays(-61); TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0' ; Enabled = $true  },
    @{ Name = 'EXARACMUN001' ; SamAccountName = 'EXARACMUN001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'MUN' ; Location = 'Munich, Germany'              ; DNSHostName = 'EXAracMUN01.example.net'  ; OU = @('Locations','Germany','Munich','Infrastructure')             ; OS = 'HPE iLO5'                          ; OperatingSystemVersion = 'N/A'                  ; Description = 'iLO – admin / MUN-ilo-pass'                     ; LastLogon = 'N/A'                      ; IPAddress = '192.168.89.2'    ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXASWIMUN001' ; SamAccountName = 'EXASWIMUN001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'MUN' ; Location = 'Munich, Germany'              ; DNSHostName = 'EXAswiMUN01.example.net'  ; OU = @('Locations','Germany','Munich','Network')                    ; OS = 'Cisco Catalyst 9200'               ; OperatingSystemVersion = 'N/A'                  ; Description = 'Access switch – admin:cisco / MUN9200!'         ; LastLogon = 'N/A'                      ; IPAddress = '192.168.89.5'    ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },
    @{ Name = 'EXASBCMUN001' ; SamAccountName = 'EXASBCMUN001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'MUN' ; Location = 'Munich, Germany'              ; DNSHostName = 'EXAsbcMUN01.example.net'  ; OU = @('Locations','Germany','Munich','Servers')                    ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-MUN-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.89.49'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null ; Enabled = $true  },

    ## Bonn, DE
    @{ Name = 'EXALAPBON001' ; SamAccountName = 'EXALAPBON001$' ; Type = 'Laptop'        ; Role = 'LAP' ; Site = 'BON' ; Location = 'Bonn, Germany'                ; DNSHostName = 'EXALAPBON001.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                    ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Lenovo ThinkPad - Pool device'                  ; LastLogon = (Get-Date).AddDays(-2)     ; IPAddress = '192.168.228.150' ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'Bon#Lap@64'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(5).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-18) ; LAPSPasswordExpiration = (Get-Date).AddDays(5)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $false }, ## Disabled for maintenance
    @{ Name = 'EXAWKSBON001' ; SamAccountName = 'EXAWKSBON001$' ; Type = 'Computer'      ; Role = 'WKS' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXAWKSBON001.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                    ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '23H2'                 ; Description = 'Finance department desktop workstation'         ; LastLogon = (Get-Date).AddDays(-11)    ; IPAddress = '192.168.228.151' ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'B0n#Wks!01'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(12).ToFileTimeUtc() ; LAPSPasswordLastSet = (Get-Date).AddDays(-20) ; LAPSPasswordExpiration = (Get-Date).AddDays(10) ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
    @{ Name = 'EXAWKSBON002' ; SamAccountName = 'EXAWKSBON002$' ; Type = 'Computer'      ; Role = 'WKS' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXAWKSBON002.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                    ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '10.0.22631'           ; Description = 'Finance Department Workstation'                 ; LastLogon = (Get-Date).AddHours(-2)    ; IPAddress = '192.168.228.152' ; Domain = 'example.net' ;  LAPSPassword = 'Kx9#mP2$vL5@qR8!'     ; LAPSPasswordExpiration = (Get-Date).AddDays(15) ; LAPSPasswordLastSet  = (Get-Date).AddDays(-15) ; 'ms-Mcs-AdmPwd' = 'Kx9#mP2$vL5@qR8!' ; 'ms-Mcs-AdmPwdExpirationTime' = '133789234567890123'                    ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  ; BitLockerRecoveryKey = '648392-174853-293847-483920-756483-029384-475829-384756' }, ## There's always that one guy
    @{ Name = 'EXALAPBON002' ; SamAccountName = 'EXALAPBON002$' ; Type = 'Computer'      ; Role = 'LAP' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXALAPBON002.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                    ; OS = 'Windows 11 Pro'                    ; OperatingSystemVersion = '23H2'                 ; Description = 'Assigned laptop for finance staff'              ; LastLogon = (Get-Date).AddDays(-21)    ; IPAddress = '192.168.228.153' ; Domain = 'example.net' ; 'msLAPS-AccountName' = 'Administrator' ; 'msLAPS-Password' = 'B0n#Lap!02'     ; 'msLAPS-PasswordExpirationTime' = (Get-Date).AddDays(7).ToFileTimeUtc()  ; LAPSPasswordLastSet = (Get-Date).AddDays(-14) ; LAPSPasswordExpiration = (Get-Date).AddDays(7)  ; TPMEnabled = $true  ; TPMActivated = $true  ; TPMVersion = '2.0'  ; Enabled = $true  },
    @{ Name = 'EXAVCUBON001' ; SamAccountName = 'EXAVCUBON001$' ; Type = 'AV'            ; Role = 'VCU' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXAVCUBON001.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                    ; OS = 'Poly Studio X70'                   ; OperatingSystemVersion = 'POLY-X70-BON-772190'  ; Description = 'Boardroom video conferencing system'            ; LastLogon = (Get-Date).AddDays(-5)     ; IPAddress = '192.168.228.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXACAMBON003' ; SamAccountName = 'EXACAMBON003$' ; Type = 'Security'      ; Role = 'CAM' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXACAMBON003.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                    ; OS = 'Axis P3245-LVE'                    ; OperatingSystemVersion = 'AXIS3245-BON-009381'  ; Description = 'CCTV security camera'                           ; LastLogon = (Get-Date).AddDays(-12)    ; IPAddress = '192.168.228.17'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXATVSBON001' ; SamAccountName = 'EXATVSBON001$' ; Type = 'AV'            ; Role = 'TVS' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXATVSBON001.example.net' ; OU = @('Locations','Germany','Bonn','Computers')                    ; OS = 'Samsung Business TV 65"'           ; OperatingSystemVersion = 'SAMS65-BON-771992'    ; Description = 'Networked information display'                  ; LastLogon = (Get-Date).AddDays(-13)    ; IPAddress = '192.168.151.18'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXARACBON001' ; SamAccountName = 'EXARACBON001$' ; Type = 'Hardware'      ; Role = 'RAC' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXARACBON001.example.net' ; OU = @('Locations','Germany','Bonn','Infrastructure')               ; OS = 'Dell iDRAC9'                       ; OperatingSystemVersion = 'N/A'                  ; Description = 'iDRAC – root / BON-RAC-01'                      ; LastLogon = 'N/A'                      ; IPAddress = '192.168.228.2'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASWIBON001' ; SamAccountName = 'EXASWIBON001$' ; Type = 'Network'       ; Role = 'SWI' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXASWIBON001.example.net' ; OU = @('Locations','Germany','Bonn','Network')                      ; OS = 'Cisco Catalyst 2960X'              ; OperatingSystemVersion = 'N/A'                  ; Description = 'Office switch – admin:cisco / BONsw01'          ; LastLogon = 'N/A'                      ; IPAddress = '192.168.228.5'   ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },
    @{ Name = 'EXASBCBON001' ; SamAccountName = 'EXASBCBON001$' ; Type = 'VOIP'          ; Role = 'SBC' ; Site = 'BON' ; Location = 'Room 305 Fl 3 Bonn Office'    ; DNSHostName = 'EXASBCBON001.example.net' ; OU = @('Locations','Germany','Bonn','Servers')                      ; OS = '3CX SBC Debian'                    ; OperatingSystemVersion = 'N/A'                  ; Description = '3CX SBC – ssh:root / 3cx-BON-49'                ; LastLogon = 'N/A'                      ; IPAddress = '192.168.228.49'  ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

    ## Miami Infra
    @{ Name = 'EXACOFMIA001' ; SamAccountName = 'EXACOFMIA001$' ; Type = 'Embedded'      ; Role = 'VND' ; Site = 'MIA' ; Location = 'Miami, FL'                    ; DNSHostName = 'EXACOFMIA001.example.net' ; OU = @('Locations','US','Florida','Miami','Vending Machhines')      ; OS = 'VxWorks'                           ; OperatingSystemVersion = '7.25.09'              ; Description = 'Networked Cuban Covfefe machine svc_coffee'     ; LastLogon = 'N/A'                      ; IPAddress='10.30.10.50'       ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  },

    ## Brockville, Ontario Infra
    @{ Name ='EXADONBRK001'  ; SamAccountName = 'EXADONBRK001$' ; Type = 'Embedded'      ; Role = 'VND' ; Site = 'BRK' ; Location = 'Brockville, ON'               ; DNSHostName = 'EXADONBRK001.example.net' ; OU = @('Locations','CA','Ontario','Brockville','Vending Machhines') ; OS = 'VxWorks'                           ; OperatingSystemVersion = '7.25.09'              ; Description = 'Donut vending machine (Tim Hortons compatible)' ; LastLogon = 'N/A'                      ; IPAddress='10.20.10.50'       ; Domain = 'example.net' ; 'msLAPS-AccountName' = $null           ; 'msLAPS-Password' = $null            ; 'msLAPS-PasswordExpirationTime' = $null                                  ; LAPSPasswordLastSet = $null                   ; LAPSPasswordExpiration = $null                  ; TPMEnabled = $false ; TPMActivated = $false ; TPMVersion = $null  ; Enabled = $true  }
  )

  # Convert demo objects to AD-like objects
  $converted = Convert-DataToADObjects -Users $Script:rawUsers -DCs $Script:rawDCs -Computers $Script:rawComputers -Groups $Script:rawDemoGroups -Domain $Script:CurrentDomain -BaseDN "DC=$($Script:CurrentDomain -replace '\.',',DC=')"
  $Script:DataSource = "Fallback"
  $Script:DataSourceInfo = @{
    Source       = "Fallback"
    CSVPath      = $null
    Server       = $null
    LoadedAt     = Get-Date
    IsReadOnly   = $false
    ObjectCounts = @{
      Users       = $Script:Users.Count
      Groups      = $Script:Groups.Count
      Computers   = $Script:Computers.Count
      DCs         = $Script:DCs.Count
    }
  }
  Debug-Log "Demo data loaded: $($Script:Users.Count) users, $($Script:DCs.Count) DCs, $($Script:Computers.Count) computers, $($Script:Groups.Count) groups and $($Script:ADObjects.Count) Objects" -Type 'Success'
  return $true
}

function Handle-CSVAction {

  param([ValidateSet('Import','Export')][string]$Action, [string]$CSVPath)

  switch ($Action) {
   'Import' {
     if (-not $CSVPath -or -not (Test-Path -LiteralPath $CSVPath)) {
       $Script:selectedFile = $null
       Show-FileBrowserDialog -Mode 'Open'
       if (-not $Script:selectedFile -or -not (Test-Path -LiteralPath $Script:selectedFile)) {
         Debug-Log "No valid file selected for import. Aborting." -Type 'Warn'
         Show-Modal "File Not Selected" "No CSV file selected. Import aborted."
         return $false
       }
       $CSVPath = $Script:selectedFile
       Debug-Log "File selected for import: $CSVPath" -Type 'Info'
     }

    try {
      Import-CSVData -CSVPath $CSVPath
      $Script:DataSource = "CSV"
      $Script:DataSourceInfo = @{
        Source       = "CSV"
        CSVPath      = $CSVPath
        Server       = $null
        LoadedAt     = Get-Date
        IsReadOnly   = $false
        ObjectCounts = @{
          Users      = $Script:Users.Count
          Groups     = $Script:Groups.Count
          Computers  = $Script:Computers.Count
          DCs        = $Script:DCs.Count
        }
      }
      Show-InfoPanel -UpdateOnly
      Debug-Log "Imported Users: $($Script:Users.Count) DCs: $($Script:DCs.Count) Computers: $($Script:Computers.Count), Groups: $($Script:Groups.Count)  AD Objects: $($Script:ADObjects.Count)" -Type "Success"
      return $true
    } catch {
      Debug-Log "Failed to import CSV: $($_.Exception.Message)" -Type 'Error'
      Show-Modal "CSV Import Failed" "Could not import data:`n$($_.Exception.Message)"
      return $false
      }
    }

   'Export' {
     if (-not $CSVPath) {
       $Script:selectedFile = $null
       Show-FileBrowserDialog -Mode 'Save'
       if (-not $Script:selectedFile) {
         Debug-Log "No file selected for export. Aborting." -Type 'Warn'
         return $false
       }
       $CSVPath = $Script:selectedFile
       Debug-Log "File selected for export: $CSVPath" -Type 'Info'
     }

      try {
        $allObjects = $Script:Users + $Script:Groups + $Script:Computers + $Script:DCs
        $allObjects | Export-Csv -Path $CSVPath -NoTypeInformation -Force
        Debug-Log "Exported $($allObjects.Count) objects to CSV: $CSVPath" -Type 'Success'
        return $true
      } catch {
        Debug-Log "Failed to export CSV: $($_.Exception.Message)" -Type 'Error'
        Show-Modal "CSV Export Failed" "Could not export data:`n$($_.Exception.Message)"
        return $false
      }
    }
  }
}

## ----------------------------{ Import CSV Data }-----------------------
function Import-CSVData {
  param([string]$CSVPath)

  Debug-Log "Importing CSV data from: $CSVPath" -Type "Info"
  $csvContent = Import-Csv -Path $CSVPath -Encoding UTF8 -ErrorAction Stop
  Debug-Log "Loaded $($csvContent.Count) rows from CSV" -Type "Info"

  ## Group by objectClass (with null safety)
  $grouped = $csvContent | Group-Object -Property objectClass -AsHashTable
  if (-not $grouped) {
    Debug-Log "No grouped data found, creating empty groups" -Type "Warn"
    $grouped = @{}
  }

  ## Extract Users (with null safety)
  $userRows = if ($grouped.ContainsKey('user')) { $grouped['user'] } else { @() }
  $users = if ($userRows -and $userRows.Count -gt 0) {
    $userRows | ForEach-Object {
      $dn = $_.distinguishedName
      $ouParts = @()
      if ($dn -match 'OU=') {
        $ouParts = ($dn -split '(?<!\\),' | Where-Object { $_ -match '^OU=' } | ForEach-Object { $_ -replace '^OU=', '' })
        [array]::Reverse($ouParts)
      }

      $domain = ($dn -split '(?<!\\),' | Where-Object { $_ -match '^DC=' } | ForEach-Object { $_ -replace '^DC=', '' }) -join '.'
      $groups = @()
      if ($_.memberOf) { $groups = $_.memberOf -split ';' | ForEach-Object { if ($_ -match 'CN=([^,]+)') { $matches[1] }} }
      $uac = [int]$_.userAccountControl

      @{
        Name               = $_.name
        SamAccountName     = $_.sAMAccountName
        UserPrincipalName  = $_.userPrincipalName
        Email              = $_.mail
        Title              = $_.title
        Department         = $_.description
        Office             = $_.physicalDeliveryOfficeName
        Phone              = $_.telephoneNumber
        MobilePhone        = $_.mobile
        Description        = $_.description
        OU                 = $ouParts
        Groups             = $groups
        Domain             = $domain
        Disabled           = ($uac -band 0x0002) -ne 0
        Locked             = ($uac -band 0x0010) -ne 0
        MustChangePassword = ($_.pwdLastSet -eq '0')
        Country            = $_.countryCode
        Manager            = ''
        Company            = ''
      }
    }
  } else { @() }

  ## Extract Groups (with null safety)
  $groupRows = if ($grouped.ContainsKey('group')) { $grouped['group'] } else { @() }
  $groups = if ($groupRows -and $groupRows.Count -gt 0) {
    $groupRows | ForEach-Object {
      $domain = ($_.distinguishedName -split '(?<!\\),' | Where-Object { $_ -match '^DC=' } | ForEach-Object { $_ -replace '^DC=', '' }) -join '.'
      $groupType = [int]$_.groupType
      $isSecurity = ($groupType -band 0x80000000) -ne 0
      $scope = switch ($groupType -band 0xF) {
        1 { 'Global' }
        2 { 'DomainLocal' }
        4 { 'Universal' }
        default { 'Global' }
      }

      @{
        Name = $_.name
        Description = $_.description
        Type        = if ($isSecurity) { 'Security' } else { 'Distribution' }
        Scope       = $scope
        Email       = $_.mail
        Domain      = $domain
        ManagedBy   = if ($_.managedBy -match 'CN=([^,]+)') { $matches[1] } else { '' }
      }
    }
  } else { @() }

  ## Extract Computers & DCs (with null safety) - Import all properties
  $computerRows = if ($grouped.ContainsKey('computer')) { $grouped['computer'] } else { @() }
  $computers = @()
  $dcs = @()

  if ($computerRows -and $computerRows.Count -gt 0) {
    foreach ($comp in $computerRows) {
      $dn = $comp.distinguishedName
      $ouParts = @()
      if ($dn -match 'OU=') {
        $ouParts = ($dn -split '(?<!\\),' | Where-Object { $_ -match '^OU=' } | ForEach-Object { $_ -replace '^OU=', '' })
        [array]::Reverse($ouParts)
      }

      $domain = ($dn -split '(?<!\\),' | Where-Object { $_ -match '^DC=' } | ForEach-Object { $_ -replace '^DC=', '' }) -join '.'
      $uac = [int]$comp.userAccountControl
      $isDC = ($uac -band 8192) -or ($comp.servicePrincipalName -match 'E3514235') -or ($dn -match 'OU=Domain Controllers')

      ## CREATE BASE OBJECT - START WITH ALL CSV PROPERTIES
      $computerObj = @{
        Name                   = $comp.name
        SamAccountName         = $comp.sAMAccountName
        Type                   = 'Computer'
        Role                   = if ($isDC) { 'DC' } else { 'WKS' }
        OU                     = $ouParts
        OS                     = $comp.operatingSystem
        OperatingSystemVersion = $comp.operatingSystemVersion
        DNSHostName            = $comp.dNSHostName
        Description            = $comp.description
        Enabled                = ($uac -band 0x0002) -eq 0
        Domain                 = $domain
        DistinguishedName      = $comp.distinguishedName
      }

      ## ADD ALL OTHER PROPERTIES FROM CSV (including LAPS)
      foreach ($prop in $comp.PSObject.Properties) {
        $propName = $prop.Name

        ## Skip properties we already added
        if ($propName -in @('name','sAMAccountName','operatingSystem','operatingSystemVersion','dNSHostName','description','distinguishedName','userAccountControl','servicePrincipalName')) { continue }
        ## Add everything else (including LAPS properties)
        if (-not [string]::IsNullOrWhiteSpace($prop.Value)) { $computerObj[$propName] = $prop.Value }
      }

      if ($isDC) {
        $isGC = $comp.servicePrincipalName -match 'GC/'
        $dcs += @{
          Name                   = $comp.name
          SamAccountName         = $comp.sAMAccountName
          DNSHostName            = $comp.dNSHostName
          Site                   = 'Default-First-Site-Name'
          Location               = ''
          Domain                 = $domain
          Forest                 = $domain
          OS                     = $comp.operatingSystem
          OperatingSystemVersion = $comp.operatingSystemVersion
          IPv4Address            = ''
          Enabled                = $true
          IsGlobalCatalog        = $isGC
          FSMORoles              = @()
          LastReplication        = Get-Date
          ReplicationHealth      = 'Healthy'
          LastBoot               = Get-Date
        }
      } else {
        $computers += $computerObj
      }
    }
  }

  ## Detect domain
  $importedDomain = if ($users.Count -gt 0) { $users[0].Domain
  } elseif ($dcs.Count -gt 0) { $dcs[0].Domain
  } elseif ($computers.Count -gt 0) { $computers[0].Domain
  } else { 'example.com'
  }
  Debug-Log "Detected domain: $importedDomain | Users: $($users.Count), Groups: $($groups.Count), Computers: $($computers.Count), DCs: $($dcs.Count)" -Type "Info"

  ## Update domain variables
  $Script:CurrentDomain = $importedDomain
  $Script:Domain        = $importedDomain
  $Script:Domains       = @($importedDomain)
  $Script:ForestName    = $importedDomain

  ## Convert to AD objects
  $baseDN = "DC=$($importedDomain -replace '\.',',DC=')"
  Convert-DataToADObjects -Users $users -DCs $dcs -Computers $computers -Groups $groups -Domain $importedDomain -BaseDN $baseDN

  ## Set data source tracking
  $Script:DataSource = "CSV"
  $Script:DataSourceInfo = @{
    Source = "CSV"
    CSVPath = $CSVPath
    Server = $null
    LoadedAt = Get-Date
    IsReadOnly = $false
    ObjectCounts = @{
      Users     = $Script:Users.Count
      Groups    = $Script:Groups.Count
      Computers = $Script:Computers.Count
      DCs       = $Script:DCs.Count
    }
  }
  Debug-Log "CSV import complete: $($Script:Users.Count) users, $($Script:Groups.Count) groups, $($Script:Computers.Count) computers, $($Script:DCs.Count) DCs" -Type "Success"
}

function Load-ADData {
  param(
    [string]$Domain = $null
  )

  Debug-Log "Querying live Active Directory..." -Type "Info"
  if (-not (Get-Module -ListAvailable ActiveDirectory)) {
    Debug-Log "ActiveDirectory module not available" -Type "Error"
    return $false
  }
  Import-Module ActiveDirectory -ErrorAction Stop

  try {
    $rootDSE = Get-ADRootDSE
    $baseDN = $rootDSE.defaultNamingContext
    $domainName = ($baseDN -replace 'DC=','' -replace ',', '.')
    Debug-Log "Connected to domain: $domainName" -Type "Info"

    ## ---------------- USERS ----------------
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

    ## ---------------- GROUPS ----------------
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

    ## ---------------- COMPUTERS & DCs ----------------
    $adComputers = Get-ADComputer -LDAPFilter "(objectClass=computer)" -SearchBase $baseDN -Properties * -ResultSetSize $null
    $computers = @()
    $dcs = @()
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

    ## ---------------- FINAL CONVERSION ----------------
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
    Debug-Log "Active Directory load failed: $($_.Exception.Message)" -Type "Error"
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
      Debug-Log "Updated ${label} in memory (${Script:DataSource} source): $($Object.Name)" -Type "Info"
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

## ----------------------------{ File Browser }-------------------------
function Show-FileBrowserDialog {
  param(
    [string]$StartDir = ".",
    [string]$Title = "Select File",
    [string[]]$Filter = @("*.*")
  )

  $script:selectedFile = $null  # Use script scope from the start
  $currentPath = (Resolve-Path $StartDir).Path
  $dialog = [Terminal.Gui.Dialog]::new($Title, 80, 24)

  ## Current path label
  $labelPath = [Terminal.Gui.Label]::new(2, 1, "Path: $currentPath")
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
      } catch { Show-Modal "Error" "Cannot access directory: $path" }
      if ($items.Count -eq 0) { $items.Add("(empty directory)") }
      $listView.SetSource($items)
    }

    ##TODO: Also selecting and pressing enter would be nice
    ## Double-click or Enter to select
    $listView.add_OpenSelectedItem({
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
    })

    ## Select button
    $btnSelect = [Terminal.Gui.Button]::new(2, 20, "Select")
    $btnSelect.add_Clicked({
      if ($script:selectedFile) {
        Debug-Log "File selected: $($script:selectedFile)"
        [Terminal.Gui.Application]::RequestStop()
      } else {
        Show-Modal "No Selection" "Please select a file"
      }
    }).GetNewClosure()
    $dialog.Add($btnSelect)

    ## Cancel button
    $btnCancel = [Terminal.Gui.Button]::new(15, 20, "Cancel")
    $btnCancel.add_Clicked({
      $script:selectedFile = $null
      [Terminal.Gui.Application]::RequestStop()
    }).GetNewClosure()
    $dialog.Add($btnCancel)

    ## Initial population
    Update-FileList -path $currentPath

    ## Run dialog
    [Terminal.Gui.Application]::Run($dialog)
    return $script:selectedFile  # Return the script-scoped variable
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

  Debug-Log ": Initializing Terminal.Gui framework..." -Type "Info"

  ## ==================== STEP 1: Initialise Terminal.Gui ====================
  try {
    [Terminal.Gui.Application]::Init()
    Debug-Log ": Terminal.Gui.Application Initialised" -Type "Success"
  } catch {
    Debug-Log ": FATAL - Failed to Initialise Terminal.Gui: $($_.Exception.Message)" -Type "Error"
    throw
  }

  ## Get top-level application
  $top = [Terminal.Gui.Application]::Top

  ## ==================== STEP 2: Create Main Window ====================
  $win = [Terminal.Gui.Window]::new($Title)
  $win.X = 0
  $win.Y = 0
  $win.Width  = [Terminal.Gui.Dim]::Fill()
  $win.Height = [Terminal.Gui.Dim]::Fill(1)  ## Leave room for status bar

  Debug-Log ": Main window created with title: $Title" -Type "Info"

  ## ==================== STEP 3: Apply Theme ====================
  Debug-Log ": Applying theme: $Theme" -Type "Info"
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

      Debug-Log ": Theme '$Theme' applied successfully" -Type "Success"
    } else { Debug-Log ": WARNING - Theme data is null, using defaults" -Type "Warn" }
  } catch {
    Debug-Log ": WARNING - Failed to apply theme: $($_.Exception.Message)" -Type "Warn"
  }

  ## ==================== STEP 4: Add Window to Top ====================
  $top.Add($win)
  Debug-Log ": Main window added to top-level application" -Type "Success"

  ## ==================== STEP 5: Return Components ====================
  $result = @{
    Top    = $top
    Window = $win
    Theme  = $themeData
  }

  Debug-Log ": UI Framework initialization complete" -Type "Success"
  return $result
}

## The right-click context menu
function Build-ContextMenuItems {
  <#
  .SYNOPSIS
  Build context menu items based on object type

  .DESCRIPTION
  Creates an ArrayList of menu items appropriate for the selected
  Active Directory object type (user, group, computer, DC, etc)

  .PARAMETER ObjectType
  Type of AD object ('user', 'group', 'computer', 'dc')

  .PARAMETER Object
  The actual AD object (used to check enabled/locked state)

  .RETURNS
  ArrayList of menu item strings

  .EXAMPLE
  $menuItems = Build-ContextMenuItems -ObjectType 'user' -Object $userObj
  Show-ContextMenu -menuItems $menuItems -obj $userObj -objType 'user'
  #>

  param(
    [Parameter(Mandatory=$true)]
    [string]$ObjectType,

    [Parameter(Mandatory=$true)]
    [object]$Object
  )

  $menuItems = [System.Collections.ArrayList]@()

  switch ($ObjectType) {
    'user' {
      [void]$menuItems.Add("Properties")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Reset Password")

      ## Dynamic menu items based on user state
      if ($Object.Enabled) { [void]$menuItems.Add("Disable Account")
      } else { [void]$menuItems.Add("Enable Account") }
      if ($Object.LockedOut -or $Object.Locked) { [void]$menuItems.Add("Unlock Account") }
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Move to OU...")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Delete")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Refresh")
    }

    'group' {
      [void]$menuItems.Add("Properties")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Add Member...")
      [void]$menuItems.Add("Remove Member...")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Delete")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Refresh")
    }

    'computer' {
      [void]$menuItems.Add("Properties")
      [void]$menuItems.Add("---")

      ## Dynamic menu items based on computer state
      if ($Object.Enabled) {
        [void]$menuItems.Add("Disable")
      } else {
        [void]$menuItems.Add("Enable")
      }

      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Move to OU...")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Delete")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Refresh")
    }

    'dc' {
      [void]$menuItems.Add("Properties")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Check Replication")
      [void]$menuItems.Add("View FSMO Roles")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Refresh")
    }

    'ou' {
      [void]$menuItems.Add("Properties")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("New User...")
      [void]$menuItems.Add("New Group...")
      [void]$menuItems.Add("New Computer...")
      [void]$menuItems.Add("New OU...")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Delete")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Refresh")
    }

    default {
      [void]$menuItems.Add("Properties")
      [void]$menuItems.Add("---")
      [void]$menuItems.Add("Refresh")
    }
  }

  return $menuItems
}

function Set-StatusBar {
  <#
  .SYNOPSIS
  Universal status bar management - Initialise, update, and animate

  .DESCRIPTION
  Single function that handles all status bar operations:
  - Initialise: Create status bar with F-key shortcuts and theme
  - Update: Set static messages
  - Spinner: Show animated spinner for long operations
  - Final: Mark operation complete with checkmark

  .PARAMETER Initialise
  Create and configure the status bar (call once at startup)

  .PARAMETER ThemeData
  Theme data to apply (used with -Initialise)

  .PARAMETER Message
  Status message to display (used for updates)

  .PARAMETER Spinner
  Show animated spinner (for long operations)

  .PARAMETER Final
  Mark operation as complete (shows checkmark)

  .EXAMPLE
  # Initialise at startup
  $statusBar = Set-StatusBar -Initialise -ThemeData $Script:themeData
  $top.Add($statusBar)

  .EXAMPLE
  # Show spinner
  Set-StatusBar "Loading data..." -Spinner

  .EXAMPLE
  # Show completion
  Set-StatusBar "Load complete" -Final

  .EXAMPLE
  # Static message
  Set-StatusBar "Ready"
  #>

  param(
    [Parameter(Position=0)]
    [string]$Message = "",

    [switch]$Initialise,
    [object]$ThemeData = $null,
    [switch]$Spinner,
    [switch]$Final
  )

  ## ==================== Initialise MODE ====================
  if ($Initialise) {
    Debug-Log ": Initializing status bar..." -Type "Info"

    ## Initialise spinner globals
    $Script:statusSpinner = @('|', '/', '-', '\')
    $Script:statusSpinnerIndex = 0
    $Script:SpinnerActive = $false
    $Script:SpinnerTimer = $null
    $Script:statusBaseMessage = ""
    $Script:statusPrefix = ""

    ## Create dynamic status item (right side)
    $Script:StatusItem = [Terminal.Gui.StatusItem]::new(0, "Initializing...", $null)

    ## Define F-key shortcuts (left side)
    $shortcuts = @(
      @{ Key = [Terminal.Gui.Key]::F1;  Label = "~F1~ Help";            Action = { Show-Modal "Shortcuts" "F1 - Help | F2 - Password | F3 - New | F5 - Refresh | F6- T hemes | F7 - Search | F9 - Show Menus | F10-Quit | F11 - Full Screen" } }
      @{ Key = [Terminal.Gui.Key]::F2;  Label = "~F2~ Password";        Action = { Generate-RandomPassword } }
      @{ Key = [Terminal.Gui.Key]::F3;  Label = "~F3~ New";             Action = { Show-NewObjectWizard } }
      @{ Key = [Terminal.Gui.Key]::F5;  Label = "~F5~ Refresh";         Action = { Refresh-Data -domain $Script:CurrentDomain -RebuildTree } }
      @{ Key = [Terminal.Gui.Key]::F6;  Label = "~F6~ Themes";          Action = { Show-ThemeSelector } }
      @{ Key = [Terminal.Gui.Key]::F7;  Label = "~F7~ Search";          Action = { Show-ADSearchDialog } }
      @{ Key = [Terminal.Gui.Key]::F9;  Label = "~F9~ Menus";           Action = { } }
      @{ Key = [Terminal.Gui.Key]::F10; Label = "~F10~ Quit";           Action = { [Terminal.Gui.Application]::RequestStop() } }
      @{ Key = [Terminal.Gui.Key]::F11; Label = "~F11~ Full Screen";    Action = { } }
    )

    ## Build status items array
    $items = @()
    foreach ($sc in $shortcuts) { $items += [Terminal.Gui.StatusItem]::new($sc.Key, $sc.Label, $sc.Action) }
    $items += $Script:StatusItem
    ## Create status bar
    $Script:StatusBar = [Terminal.Gui.StatusBar]::new($items)
    ## Apply theme if provided
    if ($ThemeData -and $ThemeData.StatusBar) {
      $Script:StatusBar.ColorScheme = $ThemeData.StatusBar
      Debug-Log ": Status bar theme applied" -Type "Success"
    }

    Debug-Log ": Status bar Initialised with $($shortcuts.Count) shortcuts" -Type "Success"
    return $Script:StatusBar
  }

  ## ==================== UPDATE MODE ====================

  ## Guard: Status bar must exist
  if (-not $Script:StatusItem -or -not $Script:StatusBar) {
    Debug-Log ": StatusBar not Initialised - call 'Set-StatusBar -Initialise' first!" -Type "Error"
    Debug-Log "StatusBar not Initialised. Call Set-StatusBar -Initialise before using." -Type "Error"
    return
  }

  ## Build dynamic prefix (refreshed on each call)
  ## TODO: The spinner can go in place of >>
  $prefix = ">>"

  ## ==================== SPINNER MODE ====================
  if ($Spinner) {
    $Script:SpinnerActive = $true
    $Script:statusBaseMessage = $Message
    $Script:statusSpinnerIndex = 0
    $Script:statusPrefix = $prefix  ## Cache for timer

    ## Create or reuse timer
    if (-not $Script:SpinnerTimer) {
      $Script:SpinnerTimer = [System.Timers.Timer]::new(200)
      $Script:SpinnerTimer.AutoReset = $true

      ## Timer callback - updates spinner character
      $Script:SpinnerTimer.Add_Elapsed({
        if ($Script:SpinnerActive -and $Script:StatusItem) {
          $Script:statusSpinnerIndex = ($Script:statusSpinnerIndex + 1) % 4
          $spinChar = $Script:statusSpinner[$Script:statusSpinnerIndex]
          $displayText = "$($Script:statusPrefix) | $spinChar $($Script:statusBaseMessage)"
          $Script:StatusItem.Title = $displayText

          ## Refresh display - explicit null check
          if ($Script:StatusBar) {
            try {
              $Script:StatusBar.SetNeedsDisplay()
            } catch {
              ## Ignore display errors in timer context
            }
          }
        }
      })
    }

    ## Start or restart timer
    if ($Script:SpinnerTimer.Enabled) { $Script:SpinnerTimer.Stop() }
    $Script:SpinnerTimer.Start()
    Debug-Log ": Spinner started: $Message" -Type "Info"
    return
  }

  ## ==================== STATIC/FINAL MODE ====================

  ## Stop spinner if running
  $Script:SpinnerActive = $false
  if ($Script:SpinnerTimer) { $Script:SpinnerTimer.Stop() }

  ## Build display text
  if ($Final) {
    $displayText = "$prefix | ✓ $Message"
    Debug-Log ": Final status: $Message" -Type "Success"
  }
  else {
    $displayText = "$prefix | $Message"
    Debug-Log ": Status: $Message" -Type "Info"
  }

  ## Update status bar
  $Script:StatusItem.Title = $displayText

  ## Refresh display - explicit null check
  if ($Script:StatusBar) {
    try {
      $Script:StatusBar.SetNeedsDisplay()
    } catch {
      ## Ignore display refresh errors
    }
  }
}

##  We show so many pop-up dialogs, use this function to reduce clutter
function Show-Modal {
  param(
    [string]$title,
    [string]$msg,
    [switch]$YesNo
  )

  if ($YesNo) {
    ## Returns 0 for Yes, 1 for No
    $result = [Terminal.Gui.MessageBox]::Query(60, 9, $title, $msg, "Yes", "No")
    return $result }
  else { [Terminal.Gui.MessageBox]::Query($title, $msg, @("OK")) | Out-Null }
}

## ----------------------------{ Initialise Data Source - SIMPLIFIED }----------------------
function Initialise-DataSource {
  <#
  .SYNOPSIS
  Routes to the appropriate data loading function based on context
  .DESCRIPTION
  Simple router - checks flags and parameters, calls the right loader
  No nested functions, just clear routing logic
  #>
  param(
    [string]$CSVPath = $null,
    [string]$Domain = $null,
    [ValidateSet('Import','Export')]
    [string]$Action = $null
  )
  ## If explicit action requested (Import/Export via menu)
  if ($Action) { return Handle-CSVAction -Action $Action -CSVPath $CSVPath }
  ## If CSV was already loaded, don't reload
  if ($Script:CSVDataLoaded) {
    Debug-Log "CSV data already loaded, skipping re-initialization" -Type "Info"
    return $true
  }
  ## Priority 1: Active Directory (if available and not in DemoMode)
  if ($Script:HasActiveDirectory -and -not $Script:DemoMode) {
    Debug-Log "Loading from Active Directory..." -Type "Info"
    if (Load-ADData -Domain $Domain) { return $true }
    ## If AD load failed, fall through to next option
  }
  ## Priority 2: CSV file (if path provided)
  if ($CSVPath -and (Test-Path $CSVPath)) {
    Debug-Log "Loading from CSV: $CSVPath" -Type "Info"
    try {
      Import-CSVData -CSVPath $CSVPath
      ## Set flags to prevent overwriting
      $Script:CSVDataLoaded = $true
      $Script:CSVDataPath = $CSVPath
      return $true
    } catch {
      Debug-Log "CSV import failed: $($_.Exception.Message)" -Type "Error"
      ## Fall through to demo data if in DemoMode
    }
  }
  ## Priority 3: Demo data (only if DemoMode enabled)
  if ($Script:DemoMode) {
    Debug-Log "Loading demo data..." -Type "Info"
    return Load-DefaultDemoData
  }

  ## After DCs are loaded/converted
  if ($Script:DCs.Count -gt 0) {
    ## Prefer Global Catalog DC
    $Script:CurrentDC = $Script:DCs | Where-Object { $_.IsGlobalCatalog } | Select-Object -First 1
    if (-not $Script:CurrentDC) { $Script:CurrentDC = $Script:DCs | Select-Object -First 1 }
    Debug-Log ": Set current DC to: $($Script:CurrentDC.Name)" -Type "Info"
  }

  ## Nothing worked
  Debug-Log "No data source available!" -Type "Error"
  return $false
}

## --------------------{ Dual-pane list dialog helper }--------------------

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
    [array]$Buttons = @(),
    [string]$Summary = "",

    [int]$Width = 120,
    [int]$Height = 40
  )

  $dialog = [Terminal.Gui.Dialog]::new()
  $dialog.Title = $Title
  $dialog.Width = $Width
  $dialog.Height = $Height

  $y = 1

  ## Summary
  if ($Summary) {
    $lblSummary = [Terminal.Gui.Label]::new($Summary)
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
          LeftList = $lstLeft
          LeftItems = $LeftPane.Items
          RightList = $lstRight
          RightItems = $RightPane.Items
          Dialog = $dialog
        }
      }.GetNewClosure())
    }

    $dialog.Add($button)
    $btnX += $btn.Text.Length + 7
  }

  ## Close button
  $btnClose = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = $btnX; $btnClose.Y = $y
  $btnClose.add_Clicked({ $dialog.RequestStop() }).GetNewClosure()
  $dialog.Add($btnClose)

  [Terminal.Gui.Application]::Run($dialog)
}

## -------------------------{ Get Theme }----------------------
function Get-Theme {
  <#
  .SYNOPSIS
  Get or dump theme color schemes

  .PARAMETER Mode
  Theme name to load

  .PARAMETER Dump
  If specified, dumps the current theme colors to Debug-Log instead of loading

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

  ## ==================== DUMP MODE ====================
  if ($Dump) {
    $themeName = if ($Mode) { $Mode } else { $Script:ThemeMode }
    Debug-Log ": Dumping colour scheme for theme: $themeName" -Type "Info"

    ## Dump the ColorSchemes that are actually being used
    Debug-Log ": === Script ColorScheme ===" -Type "Info"
      if ($Script:ScriptCs) {
        Debug-Log ":   Normal    : $($Script:ScriptCs.Normal)" -Type "Info"
        Debug-Log ":   Focus     : $($Script:ScriptCs.Focus)" -Type "Info"
        Debug-Log ":   HotNormal : $($Script:ScriptCs.HotNormal)" -Type "Info"
        Debug-Log ":   HotFocus  : $($Script:ScriptCs.HotFocus)" -Type "Info"
        Debug-Log ":   Disabled  : $($Script:ScriptCs.Disabled)" -Type "Info"
      } else {
        Debug-Log ":   ScriptCs is null!" -Type "Warn"
      }

      Debug-Log ": === Main Window ColorScheme ===" -Type "Info"
      if ($Script:mainWindowCs) {
        Debug-Log ":   Normal    : $($Script:mainWindowCs.Normal)" -Type "Info"
        Debug-Log ":   Focus     : $($Script:mainWindowCs.Focus)" -Type "Info"
        Debug-Log ":   HotNormal : $($Script:mainWindowCs.HotNormal)" -Type "Info"
        Debug-Log ":   HotFocus  : $($Script:mainWindowCs.HotFocus)" -Type "Info"
        Debug-Log ":   Disabled  : $($Script:mainWindowCs.Disabled)" -Type "Info"
      } else {
        Debug-Log ":   mainWindowCs is null!" -Type "Warn"
      }
      return
    }

    ## ==================== LOAD MODE ====================
    if (-not $Mode) { throw "Get-Theme called with empty mode" }

    ## Initialise color schemes and Ensure ColorSchemes are instantiated
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

## -------------------------{ Apply Colours }----------------------
function Apply-Theme {
  param(
    [hashtable]$ThemeData,
    [object]$TopLevel,
    [object]$MainWindow,
    [object]$Menu,
    [object]$StatusBar
  )
  if ($null -eq $ThemeData) { return }

  ## --- Terminal.Gui base colors FIRST ---
  [Terminal.Gui.Colors]::Base     = $ThemeData.Global
  [Terminal.Gui.Colors]::Dialog   = $ThemeData.Global
  [Terminal.Gui.Colors]::Menu     = $ThemeData.Global
  [Terminal.Gui.Colors]::Error    = $ThemeData.Global
  [Terminal.Gui.Colors]::TopLevel = $ThemeData.Global

  ## --- Global / TopLevel ---
  if ($TopLevel -and $TopLevel.PSObject.Properties.Name -contains 'ColorScheme') {
    $TopLevel.ColorScheme = $ThemeData.Global
    $TopLevel.SetNeedsDisplay()
  }

  ## --- Main window ---
  if ($MainWindow -and $MainWindow.PSObject.Properties.Name -contains 'ColorScheme') {
    $MainWindow.ColorScheme = $ThemeData.MainWindow
    $MainWindow.SetNeedsDisplay()
  }

  ## --- Menu ---
  if ($Menu -and $Menu.PSObject.Properties.Name -contains 'ColorScheme') {
    $Menu.ColorScheme = $ThemeData.Global
    $Menu.SetNeedsDisplay()
  }

  ## --- StatusBar ---
  if ($StatusBar -and $StatusBar.PSObject.Properties.Name -contains 'ColorScheme') {
    $StatusBar.ColorScheme = $ThemeData.Global
    $StatusBar.SetNeedsDisplay()
  }

  ## --- Tree - USE MAINWINDOW SCHEME TO MATCH PARENT ---
  if ($Script:tree -and $Script:tree.PSObject.Properties.Name -contains 'ColorScheme') {
    $Script:tree.ColorScheme = $ThemeData.MainWindow  # ← CHANGED FROM Global
    $Script:tree.SetNeedsDisplay()
  }

  ## --- Filter Panel ---
  if ($Script:filterPanel -and $Script:filterPanel.PSObject.Properties.Name -contains 'ColorScheme') {
    $Script:filterPanel.ColorScheme = $ThemeData.MainWindow  # ← Match parent
    $Script:filterPanel.SetNeedsDisplay()
  }

  ## --- Force complete refresh ---
  [Terminal.Gui.Application]::Refresh()
}

## ------------------------{ Show progress bar }----------------------
## TODO: the spinners don't spin and status text never gets progressive updates. This is not a show stopper, but does need fixed.
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

## -------------------------{ Danske Soda vand }----------------------
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

## ~~~~~~~~~~~~~~~~~~~~~~~~~{ AD FUNCTIONS BELOW HERE }~~~~~~~~~~~~~~~~~~~~~~~~~

## -------------------------{ Cleaned up DNS Dialog }-------------------------
function Show-DNSDialog {
  param([string]$Domain)

  if (-not $Domain) { $Domain = $Script:CurrentDomain }
  Debug-Log ": Opening DNS viewer for domain: $Domain" -Type "Info"

  ## Get DNS data
  $dnsZones = @()
  $dnsRecords = @()

  if ($Script:rawDNSZones)   { $dnsZones = $Script:rawDNSZones | Where-Object { $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')") }}
  if ($Script:rawDNSRecords) { $dnsRecords = $Script:rawDNSRecords | Where-Object { $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')") }}
  Debug-Log ": Found $($dnsZones.Count) zones and $($dnsRecords.Count) records" -Type "Info"

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
          $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
          $filename = "dns_records_${Domain}_$timestamp.csv"

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
          Debug-Log ": Exported DNS records to $filename" -Type "Success"
        } catch {
          Show-Modal "Error" "Failed to export DNS records:`n$($_.Exception.Message)"
          Debug-Log ": Export failed: $($_.Exception.Message)" -Type "Error"
        }
      }
    },
    @{
      Text = "Open DNS Manager"
      OnClick = {
        param($state)
        try {
          if ($IsWindows) {
            Debug-Log ": Launching dnsmgmt.msc" -Type "Info"
            Start-Process "dnsmgmt.msc" -ErrorAction Stop
            Show-Modal "DNS Manager" "Launching DNS Manager (dnsmgmt.msc)...`n`nNote: Requires administrative privileges and DNS tools installed."
          } else {
            Show-Modal "Not Available" "DNS Manager (dnsmgmt.msc) is only available on Windows.`n`nOn Linux/macOS, use 'nsupdate' or web-based DNS management tools."
          }
        } catch {
          Show-Modal "Error" "Failed to launch DNS Manager:`n$($_.Exception.Message)`n`nEnsure DNS management tools are installed."
          Debug-Log ": Failed to launch DNS Manager: $($_.Exception.Message)" -Type "Error"
        }
      }
    }
  )

  ## Show dialog
  New-DualPaneListDialog -Title "DNS Records - $Domain" -Summary "DNS Zones: $($dnsZones.Count)  |  DNS Records: $($dnsRecords.Count)" -LeftPane @{ Title = "DNS Zones:"; Items = $dnsZones; FormatItem = $formatZone } -RightPane @{ Title = "DNS Records:"; Items = $dnsRecords; FormatItem = $formatRecord } -DetailsPane @{ Title = "Record Details:"; OnSelectionChanged = $showRecordDetails } -Buttons $buttons -Width 120 -Height 40
}

## -------------------------{ DNS Zone Details Dialog }-------------------------
function Show-DNSZoneDetailsDialog {
  param($Zone)

  $details = @"
DNS Zone Details
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Zone Name: $($Zone.Name)

Distinguished Name:
$($Zone.DN)

Zone Properties:
$(if ($Zone.Properties) { $Zone.Properties | Out-String } else { "(No additional properties)" })

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"@

  Show-Modal "DNS Zone Details" $details
}

## Part of the DNS manager dialog
function Show-DNSZoneDetailsDialog {
  param($Zone)

  Debug-Log ": Showing details for DNS zone: $($Zone.Name)" -Type "Info"

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
  Debug-Log ": Opening IPSec policies viewer for domain: $Domain" -Type "Info"

  ## Get IPSec data
  $ipsecPolicies = @()

  if ($Script:rawIPSecPolicies) {
    $ipsecPolicies = $Script:rawIPSecPolicies | Where-Object {
      $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')")
    }

    if ($ipsecPolicies.Count -eq 0) { $ipsecPolicies = $Script:rawIPSecPolicies }
  }

  Debug-Log ": Found $($ipsecPolicies.Count) IPSec object(s)" -Type "Info"

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
      $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
      $filename = "ipsec_policies_${Domain}_$timestamp.csv"

      $exportData = $items | ForEach-Object {
        [PSCustomObject]@{
          Name = if ($_.IPSecName) { $_.IPSecName } else { $_.Name }
          Type = $_.Type
          IPSecID = $_.IPSecID
          Description = $_.Description
          DN = $_.DN
        }
      }

      $exportData | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
      Show-Modal "Success" "Exported $($items.Count) IPSec object(s) to:`n$filename"
      Debug-Log ": Exported IPSec policies to $filename" -Type "Success"
    } catch {
      Show-Modal "Error" "Failed to export IPSec policies:`n$($_.Exception.Message)"
      Debug-Log ": Export failed: $($_.Exception.Message)" -Type "Error"
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
  Debug-Log ": Opening Print Queues viewer for domain: $Domain" -Type "Info"

  ## Get print queue data
  $printQueues = @()

  if ($Script:rawPrintQueues) {
    $printQueues = $Script:rawPrintQueues | Where-Object { $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')") }
    if ($printQueues.Count -eq 0) { $printQueues = $Script:rawPrintQueues }
  }
  Debug-Log ": Found $($printQueues.Count) print queue(s)" -Type "Info"
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
    $name = $printer.PrinterName ?? $printer.Name ?? "(Unnamed)"
    $server = if ($printer.ServerName) { " on $($printer.ServerName)" } else { "" }
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
          $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
          $filename = "print_queues_${Domain}_$timestamp.csv"

          $exportData = $printQueues | ForEach-Object {
            [PSCustomObject]@{
              PrinterName = $_.PrinterName ?? $_.Name
              ServerName = $_.ServerName
              ShareName = $_.ShareName
              UNCPath = if ($_.ServerName -and $_.ShareName) { "\\$($_.ServerName)\$($_.ShareName)" } else { "" }
              Location = $_.Location
              DN = $_.DN
            }
          }

          $exportData | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
          Show-Modal "Success" "Exported $($printQueues.Count) print queue(s) to:`n$filename"
          Debug-Log ": Exported print queues to $filename" -Type "Success"
        } catch {
          Show-Modal "Error" "Failed to export print queues:`n$($_.Exception.Message)"
          Debug-Log ": Export failed: $($_.Exception.Message)" -Type "Error"
        }
      }
    },
    @{
      Text = "Open Print Management"
      OnClick = {
        param($state)
        try {
          if ($IsWindows) {
            Debug-Log ": Launching printmanagement.msc" -Type "Info"
            Start-Process "printmanagement.msc" -ErrorAction Stop
            Show-Modal "Print Management" "Launching Print Management Console...`n`nNote: Requires Print Services management tools installed."
          } else {
            Show-Modal "Not Available" "Print Management Console (printmanagement.msc) is only available on Windows.`n`nOn Linux, use CUPS web interface (http://localhost:631) or cupsd commands."
          }
        } catch {
          Show-Modal "Error" "Failed to launch Print Management:`n$($_.Exception.Message)`n`nInstall Print Services management tools via Server Manager."
          Debug-Log ": Failed to launch printmanagement.msc: $($_.Exception.Message)" -Type "Error"
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
  Debug-Log ": Opening Trusts viewer for domain: $Domain" -Type "Info"
  ## Get trust data
  $trusts = @()
  if ($Script:rawTrusts) {
    $trusts = $Script:rawTrusts | Where-Object { $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')") }
    if ($trusts.Count -eq 0) { $trusts = $Script:rawTrusts }
  }

  Debug-Log ": Found $($trusts.Count) trust(s)" -Type "Info"

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
      $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
      $filename = "trusts_${Domain}_$timestamp.csv"

      $exportData = $items | ForEach-Object {
        [PSCustomObject]@{
          Partner = if ($_.Partner) { $_.Partner } else { $_.Name }
          Type = $_.Type
          Direction = $_.Direction
          Created = $_.Created
          Modified = $_.Modified
          DN = $_.DN
        }
      }

      $exportData | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
      Show-Modal "Success" "Exported $($items.Count) trust(s) to:`n$filename"
      Debug-Log ": Exported trusts to $filename" -Type "Success"
    } catch {
      Show-Modal "Error" "Failed to export trusts:`n$($_.Exception.Message)"
      Debug-Log ": Export failed: $($_.Exception.Message)" -Type "Error"
    }
  }

  ## Show dialog using existing helper
  New-ListDialog -Title "Trust Relationships - $Domain" -Items $trusts -FormatItem $formatTrust -OnView $onView -OnExport $onExport -FilterHelp "(Filter by partner name or trust type)"
}

function Test-TrustConnection {
  param($Trust)

  $partner = if ($Trust.Partner) { $Trust.Partner } else { $Trust.Name }
  Debug-Log ": Testing trust connection to $partner" -Type "Info"

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
      ## TODO: Put this in the demo data and call it like other AD objects
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
        # Production mode - actually test
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

##### <------------------------------- cleanup to here

## ----------------------------{ Audit Log viewer }-------------------------
function Show-AuditLogDialog {
  <#
  .SYNOPSIS
  Show audit log for an AD object (User, Group, Computer)

  .PARAMETER Object
  The AD object to show logs for

  .PARAMETER ObjectType
  Type of object: 'User', 'Group', 'Computer'

  .EXAMPLE
  Show-AuditLogDialog -Object $user -ObjectType 'User'
  #>

  param(
    [Parameter(Mandatory=$true)]
    [object]$Object,
    [Parameter(Mandatory=$true)]
    [ValidateSet('User', 'Group', 'Computer')]
    [string]$ObjectType
  )

  $objectName = $Object.Name
  Debug-Log ": Showing audit log for $ObjectType '$objectName'" -Type "Info"

  ## ==================== GENERATE LOG ENTRIES ====================
  $logEntries = @()

  if ($Script:DemoMode) {
    ## Demo mode - generate fake but realistic log entries
    $baseDate = Get-Date

    ## Generate pseudo-random entries based on name hash (consistent per object)
    $hash = 0
    foreach ($char in $objectName.ToCharArray()) { $hash += [int]$char }

    $entryCount = ($hash % 15) + 5  # 5-20 entries

    $actionTypes = @(
      @{ Type = "Created"; Details = "Account created" }
      @{ Type = "Modified"; Details = "DisplayName changed" }
      @{ Type = "Modified"; Details = "Description updated" }
      @{ Type = "Modified"; Details = "Email address changed" }
      @{ Type = "Modified"; Details = "Department changed" }
      @{ Type = "Password Reset"; Details = "Password changed" }
      @{ Type = "Password Reset"; Details = "Password expired and reset" }
      @{ Type = "Group Change"; Details = "Added to group" }
      @{ Type = "Group Change"; Details = "Removed from group" }
      @{ Type = "Account Status"; Details = "Account enabled" }
      @{ Type = "Account Status"; Details = "Account disabled" }
      @{ Type = "Account Status"; Details = "Account unlocked" }
      @{ Type = "Permission Change"; Details = "Permissions modified" }
      @{ Type = "Property Change"; Details = "Phone number updated" }
      @{ Type = "Property Change"; Details = "Address updated" }
    )

    $users = @("admin", "helpdesk", "system", "jsmith", "mjones", "ITAdmin")
    for ($i = 0; $i -lt $entryCount; $i++) {
      $daysAgo = ($hash + $i * 7) % 120  # 0-120 days ago
      $timestamp = $baseDate.AddDays(-$daysAgo).AddHours(-($i % 24)).AddMinutes(-($i * 13 % 60))
      $action = $actionTypes[($hash + $i) % $actionTypes.Count]
      $by = $users[($hash + $i) % $users.Count]
      $logEntries += [PSCustomObject]@{
        Timestamp = $timestamp
        Action = $action.Type
        Details = $action.Details
        By = $by
        ObjectType = $ObjectType
        ObjectName = $objectName
      }
    }

    ## Sort by timestamp descending (newest first)
    $logEntries = $logEntries | Sort-Object -Property Timestamp -Descending

  } else {
    ## Production mode - query actual AD audit logs (if available)
    ## Note: This requires audit logging to be enabled in AD and appropriate permissions
    try {
      ## Try to get audit events from Windows Event Log
      $filter = @{
        LogName = 'Security'
        ID = 4720,4722,4723,4724,4725,4726,4738,4740,4767  # AD audit event IDs
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

        $logEntries += [PSCustomObject]@{
          Timestamp = $event.TimeCreated
          Action = $action
          Details = $event.Message.Split("`n")[0]  # First line only
          By = "System"
          ObjectType = $ObjectType
          ObjectName = $objectName
        }
      }

      if ($logEntries.Count -eq 0) {
      ## No audit events found - add placeholder
        $logEntries += [PSCustomObject]@{
          Timestamp = Get-Date
          Action = "No Audit Data"
          Details = "No audit events found (audit logging may not be enabled)"
          By = "System"
          ObjectType = $ObjectType
          ObjectName = $objectName
        }
      }

    } catch {
      Debug-Log ": Failed to query audit logs: $($_.Exception.Message)" -Type "Warn"
      $logEntries += [PSCustomObject]@{
        Timestamp = Get-Date
        Action = "Error"
        Details = "Failed to retrieve audit logs: $($_.Exception.Message)"
        By = "System"
        ObjectType = $ObjectType
        ObjectName = $objectName
      }
    }
  }

  ## ==================== CREATE DIALOG ====================
  $dlg = [Terminal.Gui.Dialog]::new("Audit Log - ${ObjectType}: $objectName", 90, 30)
  $y = 1

  ## Header
  $lblHeader = [Terminal.Gui.Label]::new("Recent audit events for $objectName")
  $lblHeader.X = 2
  $lblHeader.Y = $y
  $dlg.Add($lblHeader)
  $y += 2

  ## Filter options
  $lblFilter = [Terminal.Gui.Label]::new("Filter:")
  $lblFilter.X = 2
  $lblFilter.Y = $y
  $dlg.Add($lblFilter)

  $rdoFilter = [Terminal.Gui.RadioGroup]::new(@("All", "Modifications", "Password Resets", "Group Changes", "Account Status"))
  $rdoFilter.X = 10
  $rdoFilter.Y = $y
  $rdoFilter.SelectedItem = 0
  $dlg.Add($rdoFilter)
  $y += 3

  ## Log entries list
  $lstLog = [Terminal.Gui.ListView]::new()
  $lstLog.X = 2
  $lstLog.Y = $y
  $lstLog.Width = [Terminal.Gui.Dim]::Fill(2)
  $lstLog.Height = [Terminal.Gui.Dim]::Fill(5)

  ## Format log entries for display
  function Get-FilteredEntries {
    param($filterIndex)

    $filtered = switch ($filterIndex) {
      0 { $logEntries }  # All
      1 { $logEntries | Where-Object { $_.Action -eq "Modified" -or $_.Action -eq "Property Change" } }  # Modifications
      2 { $logEntries | Where-Object { $_.Action -eq "Password Reset" } }  # Password Resets
      3 { $logEntries | Where-Object { $_.Action -eq "Group Change" } }  # Group Changes
      4 { $logEntries | Where-Object { $_.Action -eq "Account Status" } }  # Account Status
    }

    $displayItems = @()
    foreach ($entry in $filtered) {
      $timeStr = $entry.Timestamp.ToString("yyyy-MM-dd HH:mm")
      $displayItems += "$timeStr | $($entry.Action.PadRight(15)) | $($entry.Details.PadRight(35)) | By: $($entry.By)"
    }

    if ($displayItems.Count -eq 0) {
      $displayItems = @("(No entries match filter)")
    }
    return $displayItems
  }

  ## Initial load
  $displayItems = Get-FilteredEntries -filterIndex 0
  $lstLog.SetSource($displayItems)
  $dlg.Add($lstLog)

  ## Update list when filter changes
  $rdoFilter.add_SelectedItemChanged({
    $displayItems = Get-FilteredEntries -filterIndex $rdoFilter.SelectedItem
    $lstLog.SetSource($displayItems)
  }.GetNewClosure())

  ## Export button
  $btnExport = [Terminal.Gui.Button]::new("Export to CSV")
  $btnExport.add_Clicked({
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $filename = "audit_log_${objectName}_$timestamp.csv"

    try {
      $logEntries | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
      Show-Modal "Export Complete" "Audit log exported to:`n`n$filename"
      Debug-Log ": Exported audit log to $filename" -Type "Success"
    } catch {
      Show-Modal "Export Failed" "Failed to export log:`n`n$($_.Exception.Message)"
      Debug-Log ": Failed to export audit log: $($_.Exception.Message)" -Type "Error"
    }
  }.GetNewClosure())
  $dlg.AddButton($btnExport)

  ## Close button
  $btnClose = [Terminal.Gui.Button]::new("Close")
  $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
  $dlg.AddButton($btnClose)

  ## Run dialog
  [Terminal.Gui.Application]::Run($dlg)

  Debug-Log ": Audit log dialog closed" -Type "Info"
}

## TODO: This is meant to be inside functions
## View Audit Log button
<#
$btnAuditLog = [Terminal.Gui.Button]::new("View Audit Log...")
$btnAuditLog.X = 2
$btnAuditLog.Y = $y  # Use appropriate Y position
$btnAuditLog.add_Clicked({
    ## Determine object type
    $objType = if ($user) { 'User' }
               elseif ($group) { 'Group' }
               elseif ($computer) { 'Computer' }
               else { 'User' }

    ## Get the object (use whichever variable is appropriate: $user, $group, $computer)
    $obj = if ($user) { $user }
           elseif ($group) { $group }
           elseif ($computer) { $computer }
           else { $user }

    Show-AuditLogDialog -Object $obj -ObjectType $objType
}.GetNewClosure())
$view.Add($btnAuditLog)
## end of TODO ##
#>

## ==================== BULK ATTRIBUTE EDITOR ====================
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
  $hasUsers = $false
  $hasGroups = $false
  $hasComputers = $false

  foreach ($obj in $Objects) {
    if ($obj.PSObject.Properties.Match('SamAccountName') -and -not $obj.PSObject.Properties.Match('ComputerType')) { $hasUsers = $true }
    elseif ($obj.PSObject.Properties.Match('Members')) { $hasGroups = $true }
    elseif ($obj.PSObject.Properties.Match('ComputerType')) { $hasComputers = $true }
  }

  ## Define available attributes by object type
  $userAttributes =  @( 'DisplayName', 'Description', 'EmailAddress', 'Title', 'Department', 'Company', 'Manager', 'OfficePhone', 'MobilePhone', 'StreetAddress', 'City', 'PostalCode', 'Country', '--- PASSWORD OPTIONS ---', 'ChangePasswordAtLogon', 'ResetPassword' )
  $groupAttributes = @( 'Description', 'Email', 'ManagedBy' )
  $computerAttributes = @( 'Description', 'Location' )

  ## Build combined attribute list based on selected objects
  $availableAttributes = @()
  if ($hasUsers) { $availableAttributes += $userAttributes }
  if ($hasGroups) { $availableAttributes += $groupAttributes }
  if ($hasComputers) { $availableAttributes += $computerAttributes }
  $availableAttributes = $availableAttributes | Select-Object -Unique | Sort-Object

  ## ==================== INTERACTIVE DIALOG ====================
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
          $lblHelp.Text = [NStack.ustring]::Make("Value: 'true' or 'false'")
          $txtValue.Text = [NStack.ustring]::Make("true")
        } elseif ($attr -eq 'ResetPassword') {
          $lblHelp.Text = [NStack.ustring]::Make("Enter new password (leave blank for random)")
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

  ## ==================== DIRECT ATTRIBUTE UPDATE ====================
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
    Debug-Log ": Bulk attribute update cancelled by user" -Type "Info"
    return
  }

  Debug-Log ": Starting bulk attribute update: $Attribute = '$Value' on $($Objects.Count) objects" -Type "Info"

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
        Debug-Log ":   Skipped $name - attribute not valid for $objectType" -Type "Warn"
        continue
    }

    try {

        if ($Script:DemoMode) {

            ## Demo mode
            if ($Attribute -eq 'ChangePasswordAtLogon') {
                $boolValue = [bool]::Parse($Value)
                if ($obj.PSObject.Properties.Match('ChangePasswordAtLogon')) {
                    $obj.ChangePasswordAtLogon = $boolValue
                } else {
                    $obj | Add-Member -NotePropertyName 'ChangePasswordAtLogon' -NotePropertyValue $boolValue -Force
                }
                $successCount++
                Debug-Log ":   Updated $objectType '$name': ChangePasswordAtLogon = $boolValue (demo mode)" -Type "Info"

            } elseif ($Attribute -eq 'ResetPassword') {

                $passwordValue = if ([string]::IsNullOrWhiteSpace($Value)) {
                    -join ((65..90) + (97..122) + (48..57) + (33,35,36,37,38,42,63) | Get-Random -Count 12 | ForEach-Object {[char]$_})
                } else {
                    $Value
                }

                if ($obj.PSObject.Properties.Match('Password')) {
                    $obj.Password = $passwordValue
                } else {
                    $obj | Add-Member -NotePropertyName 'Password' -NotePropertyValue $passwordValue -Force
                }
                $successCount++
                Debug-Log ":   Reset password for $objectType '$name' (demo mode, password: $passwordValue)" -Type "Info"

            } else {

                ## Standard attribute
                if ($obj.PSObject.Properties.Match($Attribute)) {
                    $obj.$Attribute = $Value
                } else {
                    $obj | Add-Member -NotePropertyName $Attribute -NotePropertyValue $Value -Force
                }
                $successCount++
                Debug-Log ":   Updated $objectType '$name': $Attribute = '$Value' (demo mode)" -Type "Info"

            }

        } else {

            ## Production mode
            if ($Attribute -eq 'ChangePasswordAtLogon') {

                $boolValue = [bool]::Parse($Value)
                Set-UnifiedObject -ObjectType User -Object $obj -Properties @{ ChangePasswordAtLogon = $boolValue }
                $successCount++
                Debug-Log ":   Updated $objectType '$name': ChangePasswordAtLogon = $boolValue in AD" -Type "Success"

            } elseif ($Attribute -eq 'ResetPassword') {

                $passwordValue = if ([string]::IsNullOrWhiteSpace($Value)) {
                    $pw = -join (
                        (65..90) + (97..122) + (48..57) + (33,35,36,37,38,42,63) |
                        Get-Random -Count 12 |
                        ForEach-Object { [char]$_ }
                    )
                    ConvertTo-SecureString -String $pw -AsPlainText -Force
                } else {
                    ConvertTo-SecureString -String $Value -AsPlainText -Force
                }

                Set-ADAccountPassword -Identity $obj.SamAccountName -NewPassword $passwordValue -Reset -ErrorAction Stop
                $successCount++
                Debug-Log ":   Reset password for $objectType '$name' in AD" -Type "Success"

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
                Debug-Log ":   Updated $objectType '$name': $Attribute = '$Value' in AD" -Type "Success"

            }

        } ## End production mode

    } catch {
        $failCount++
        $errors += "${name}: $($_.Exception.Message)"
        Debug-Log ":   Failed to update $objectType '$name': $($_.Exception.Message)" -Type "Error"
    }

  } ## End foreach
}
## ----------------------------{ Apply Object Changes }-------------------------
function Apply-ObjectChanges {
    <#
    .SYNOPSIS
    Unified function to apply changes to AD objects (Users, Groups, OUs, Computers)

    .PARAMETER ObjectType
    Type of object: 'User', 'Group', 'OU', 'Computer'

    .PARAMETER Object
    The object being modified

    .PARAMETER State
    Hashtable containing all field references from the dialog

    .EXAMPLE
    Apply-ObjectChanges -ObjectType 'User' -Object $user -State $state
    #>
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('User', 'Group', 'OU', 'Computer')]
        [string]$ObjectType,

        [Parameter(Mandatory=$true)]
        [object]$Object,

        [Parameter(Mandatory=$true)]
        [hashtable]$State
    )

    Debug-Log ": Applying $ObjectType changes for: $($Object.Name)" -Type "Info"

    try {
        if ($Script:DemoMode) {
            # Demo mode - update object directly
            switch ($ObjectType) {
                'User' {
                    if ($State.txtDisplayName) { $Object.DisplayName = $State.txtDisplayName.Text.ToString() }
                    if ($State.txtDescription) { $Object.Description = $State.txtDescription.Text.ToString() }
                    if ($State.txtEmail) { $Object.EmailAddress = $State.txtEmail.Text.ToString(); $Object.mail = $State.txtEmail.Text.ToString() }
                    if ($State.txtOfficePhone) { $Object.OfficePhone = $State.txtOfficePhone.Text.ToString() }
                    if ($State.txtMobilePhone) { $Object.MobilePhone = $State.txtMobilePhone.Text.ToString() }
                    if ($State.txtStreet) { $Object.StreetAddress = $State.txtStreet.Text.ToString() }
                    if ($State.txtCity) { $Object.City = $State.txtCity.Text.ToString() }
                    if ($State.txtPostal) { $Object.PostalCode = $State.txtPostal.Text.ToString() }
                    if ($State.txtCountry) { $Object.Country = $State.txtCountry.Text.ToString() }
                    if ($State.txtTitle) { $Object.Title = $State.txtTitle.Text.ToString() }
                    if ($State.txtDept) { $Object.Department = $State.txtDept.Text.ToString() }
                    if ($State.txtCompany) { $Object.Company = $State.txtCompany.Text.ToString() }
                    if ($State.txtManager) { $Object.Manager = $State.txtManager.Text.ToString() }
                    if ($State.txtSamAccountName) { $Object.SamAccountName = $State.txtSamAccountName.Text.ToString() }
                    if ($State.txtUserPrincipalName) { $Object.UserPrincipalName = $State.txtUserPrincipalName.Text.ToString() }
                    if ($State.chkEnabled) { $Object.Enabled = $State.chkEnabled.Checked; $Object.Disabled = -not $State.chkEnabled.Checked }
                }

                'Group' {
                    if ($State.txtDescription) { $Object.Description = $State.txtDescription.Text.ToString() }
                    if ($State.txtEmail) { $Object.Email = $State.txtEmail.Text.ToString(); $Object.mail = $State.txtEmail.Text.ToString() }
                    if ($State.txtManagedBy) { $Object.ManagedBy = $State.txtManagedBy.Text.ToString() }

                    # Update in raw demo groups if exists
                    $rawGroup = $Script:rawDemoGroups | Where-Object { $_.Name -eq $Object.Name } | Select-Object -First 1
                    if ($rawGroup) {
                        if ($State.txtDescription) { $rawGroup.Description = $State.txtDescription.Text.ToString() }
                        if ($State.txtEmail) { $rawGroup.Email = $State.txtEmail.Text.ToString() }
                        if ($State.txtManagedBy) { $rawGroup.ManagedBy = $State.txtManagedBy.Text.ToString() }
                    }
                }

                'OU' {
                    $originalName = $State.originalName ?? $Object.Name
                    $newName = if ($State.txtName) { $State.txtName.Text.ToString() } else { $Object.Name }
                    $newDesc = if ($State.txtDesc) { $State.txtDesc.Text.ToString() } else { $Object.Description }

                    $isRename = $originalName -ne $newName
                    if ($isRename) {
                        Debug-Log ": Renaming OU from '$originalName' to '$newName'" -Type "Info"
                        $Object.Name = $newName

                        # Update all users that reference this OU
                        foreach ($user in $Script:Users) {
                            if ($user.OU -contains $originalName) {
                                $user.OU = $user.OU | ForEach-Object { if ($_ -eq $originalName) { $newName } else { $_ } }
                            }
                        }

                        # Update raw users
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

                        if ($State.originalName) { $State.originalName = $newName }
                    }

                    $Object.Description = $newDesc
                }

                'Computer' {
                    if ($State.txtDescription) { $Object.Description = $State.txtDescription.Text.ToString() }
                    if ($State.txtLocation) { $Object.Location = $State.txtLocation.Text.ToString() }
                }
            }

            Debug-Log "SUCCESS: $ObjectType changes applied (demo mode)" -Type "Success"
            Show-Modal "Success" "Changes applied successfully (demo mode)"

        } else {
          ## Production mode - use AD cmdlets
          switch ($ObjectType) {
            'User' {
              ## Build property hashtable from UI fields
              $setParams = @{}

              if ($State.txtDisplayName)    { $setParams['DisplayName']   = $State.txtDisplayName.Text.ToString().Trim() }
              if ($State.txtDescription)    { $setParams['Description']   = $State.txtDescription.Text.ToString().Trim() }
              if ($State.txtEmail)          { $setParams['EmailAddress']  = $State.txtEmail.Text.ToString().Trim(); $setParams['Email'] = $State.txtEmail.Text.ToString().Trim() }
              if ($State.txtOfficePhone)    { $setParams['OfficePhone']   = $State.txtOfficePhone.Text.ToString().Trim() }
              if ($State.txtMobilePhone)    { $setParams['MobilePhone']   = $State.txtMobilePhone.Text.ToString().Trim() }
              if ($State.txtStreet)         { $setParams['StreetAddress'] = $State.txtStreet.Text.ToString().Trim() }
              if ($State.txtCity)           { $setParams['City']          = $State.txtCity.Text.ToString().Trim() }
              if ($State.txtPostal)         { $setParams['PostalCode']    = $State.txtPostal.Text.ToString().Trim() }
              if ($State.txtCountry)        { $setParams['Country']       = $State.txtCountry.Text.ToString().Trim() }
              if ($State.txtTitle)          { $setParams['Title']         = $State.txtTitle.Text.ToString().Trim() }
              if ($State.txtDept)           { $setParams['Department']    = $State.txtDept.Text.ToString().Trim() }
              if ($State.txtCompany)        { $setParams['Company']       = $State.txtCompany.Text.ToString().Trim() }
              if ($State.txtManager)        { $setParams['Manager']       = $State.txtManager.Text.ToString().Trim() }
              if ($State.txtSamAccountName) { $newSam = $State.txtSamAccountName.Text.ToString().Trim() ; if ($newSam -ne $Object.SamAccountName) { $setParams['SamAccountName'] = $newSam } }
              if ($State.txtUPN)            { $newUPN = $State.txtUPN.Text.ToString().Trim() ; if ($newUPN -ne $Object.UserPrincipalName) { $setParams['UserPrincipalName'] = $newUPN }}
              if ($State.txtUserPrincipalName) { $newUPN = $State.txtUserPrincipalName.Text.ToString().Trim()
                if ($newUPN -ne $Object.UserPrincipalName) { $setParams['UserPrincipalName'] = $newUPN }
              }

              ## Apply all updates with unified function
              if ($setParams.Count -gt 0) { Set-UnifiedObject -ObjectType User -Object $Object -Properties $setParams }

              ## Handle account status
              if ($State.chkEnabled) {
                if ($State.chkEnabled.Checked -and -not $Object.Enabled) { Enable-ADAccount -Identity $Object.SamAccountName -ErrorAction Stop
                } elseif (-not $State.chkEnabled.Checked -and $Object.Enabled) { Disable-ADAccount -Identity $Object.SamAccountName -ErrorAction Stop }
              }
            ## Apply all updates with unified function
            if ($setParams.Count -gt 0) { Set-UnifiedObject -ObjectType User -Object $Object -Properties $setParams }
          }

         'Group' {
            ## Build unified property bag
            $properties = @{}
            if ($State.txtDescription) { $properties['Description'] = $State.txtDescription.Text.ToString()}
            if ($State.txtEmail -and $State.txtEmail.Text.ToString()) {
              ## AD attribute + in-memory property
              $properties['mail']  = $State.txtEmail.Text.ToString()
              $properties['Email'] = $State.txtEmail.Text.ToString()
            }
            if ($State.txtManagedBy -and $State.txtManagedBy.Text.ToString()) { $properties['ManagedBy'] = $State.txtManagedBy.Text.ToString() }
              ## Apply via unified setter
              if ($properties.Count -gt 0) { Set-UnifiedObject -ObjectType Group -Object $Object -Properties $properties }
            }

          'OU' {
            $originalName = $State.originalName ?? $Object.Name
            $newName = if ($State.txtName) { $State.txtName.Text.ToString() } else { $Object.Name }
            $newDesc = if ($State.txtDesc) { $State.txtDesc.Text.ToString() } else { "" }
            $isRename = $originalName -ne $newName
            if ($isRename) {
              $adOU = Get-ADOrganizationalUnit -Filter "Name -eq '$originalName'" -ErrorAction Stop | Select-Object -First 1
              if ($adOU) {
                Rename-ADObject -Identity $adOU.DistinguishedName -NewName $newName -ErrorAction Stop
                if ($newDesc) {
                  $newDN = "OU=$newName,$($adOU.DistinguishedName -replace '^OU=[^,]+,')"
                  Set-ADOrganizationalUnit -Identity $newDN -Description $newDesc -ErrorAction Stop
                }
                $Object.Name = $newName
                if ($State.originalName) { $State.originalName = $newName }
              }
            } else {
              $adOU = Get-ADOrganizationalUnit -Filter "Name -eq '$originalName'" -ErrorAction Stop | Select-Object -First 1
              if ($adOU -and $newDesc) { Set-ADOrganizationalUnit -Identity $adOU.DistinguishedName -Description $newDesc -ErrorAction Stop }
            }

            $Object.Description = $newDesc
          }

          'Computer' {
            ## Build unified property bag
            $properties = @{}
            if ($State.txtDescription) { $properties['Description'] = $State.txtDescription.Text.ToString() }
            if ($State.txtLocation)    { $properties['Location'] = $State.txtLocation.Text.ToString()       }
            ## Apply via unified setter
            if ($properties.Count -gt 0) { Set-UnifiedObject -ObjectType Computer -Object $Object -Properties $properties }
          }
        }

        Debug-Log "SUCCESS: $ObjectType changes applied to AD" -Type "Success"
        Show-Modal "Success" "Changes applied successfully"
      }

      ## Refresh data
      Refresh-Data -domain $Script:CurrentDomain

      ## Clear change tracking flags
      $Script:changesMade = $false
      $Script:groupChangesMade = $false
      $Script:ouChangesMade = $false
      return $true

    } catch {
      Debug-Log ": Failed to apply $ObjectType changes: $($_.Exception.Message)" -Type "Error"
      Show-Modal "Error" "Failed to apply changes:`n$($_.Exception.Message)"
      return $false
    }
}

## -------------------{ Unified Refresh Domain Data }------------------
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

  if (-not $Domain) {
    $Domain = $Script:CurrentDomain
  }

  Debug-Log ": Refreshing data for domain: $Domain" -Type "Info"

  ## Show loading dialog if requested
  $loadingDlg = $null
  if ($ShowLoadingDialog) {
    $loadingDlg = Show-LoadingDialog -Message "Refreshing Active Directory data..."
  }

  try {
    ## Demo mode - reload from raw data
    if ($Script:DemoMode) {
      Set-StatusBar "Refreshing demo data..." -spinner

      ## Reload from raw data (preserves CSV imports)
      if ($Script:DataSource -eq "CSV") {
        ## CSV data already in Script:Users/Groups/Computers/DCs
        Debug-Log ": Refreshing from CSV data source" -Type "Info"
      } else {
        ## Using default demo data
        $baseDN = "DC=$($Domain -replace '\.',',DC=')"
        Convert-DataToADObjects -Users $Script:rawUsers -DCs $Script:rawDCs -Groups $Script:rawDemoGroups -Computers $Script:rawComputers -Domain $Domain -BaseDN $baseDN
      }

      if ($RebuildTree) {
        Set-StatusBar "Rebuilding tree..." -spinner
        [Terminal.Gui.Application]::MainLoop.Invoke({
          try {
            $Script:tree.ClearObjects()
            Build-Tree -domain $Domain
            Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
            [Terminal.Gui.Application]::Refresh()
          } catch {
            Debug-Log ": Tree rebuild error: $_" -Type "Error"
          }
        })
      }

      Set-StatusBar "Demo data refreshed" -final

      if ($ShowModal) {
        Show-Modal "Refreshed" "Demo data refreshed successfully"
      }

      return $true
    }

    ## Production mode - query AD
    Set-StatusBar "Loading domain controllers..." -spinner
    $dcs = Get-ADDomainController -Filter * -Server $Domain -ErrorAction Stop
    Set-StatusBar "Loading users..." -spinner
    $users = Get-ADUser -Filter * -Server $Domain -ErrorAction Stop
    Set-StatusBar "Loading groups..." -spinner
    $groups = Get-ADGroup -Filter * -Server $Domain -ErrorAction Stop
    Set-StatusBar "Loading computers..." -spinner
    $computers = Get-ADComputer -Filter * -Server $Domain -ErrorAction Stop
    ## Validate results
    if ($null -eq $dcs -or $null -eq $users -or $null -eq $groups -or $null -eq $computers) {
      Debug-Log ": Refresh failed - one or more queries returned null" -Type "Warn"
      Set-StatusBar "Refresh failed - check logs" -final
      return $false
    }

    Debug-Log ": Loaded $($users.Count) users, $($groups.Count) groups, $($computers.Count) computers, $($dcs.Count) DCs" -Type "Info"
    ## Convert to standardized format
    Set-StatusBar "Converting data..." -spinner
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

    Convert-DataToADObjects -Users $userList -DCs $dcList -Computers $computerList -Groups $groupList -Domain $Domain -BaseDN $baseDN

    ## Update data source tracking
    $Script:DataSource = "AD"
    $Script:DataSourceInfo = @{
      Source = "AD"
      Server = $Domain
      LoadedAt = Get-Date
      IsReadOnly = $false
      ObjectCounts = @{
        Users = $Script:Users.Count
        Groups = $Script:Groups.Count
        Computers = $Script:Computers.Count
        DCs = $Script:DCs.Count
      }
    }

    ## Rebuild tree if requested
    if ($RebuildTree) {
      Set-StatusBar "Rebuilding tree..." -spinner
      [Terminal.Gui.Application]::MainLoop.Invoke({
        try {
          $Script:tree.ClearObjects()
          Build-Tree -domain $Domain
          Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
          Show-InfoPanel -UpdateOnly
          [Terminal.Gui.Application]::Refresh()
        } catch {
          Debug-Log ": Tree rebuild error: $_" -Type "Error"
        }
      })
    }

    Set-StatusBar "Refresh complete" -final
    if ($ShowModal) { Show-Modal "Refreshed" "Active Directory data refreshed successfully" }
    return $true
  } catch {
    Debug-Log ": Refresh error: $_" -Type "Error"
    Set-StatusBar "Refresh error" -final
    if ($ShowModal) { Show-Modal "Error" "Failed to refresh data:`n`n$($_.Exception.Message)" }
    return $false
  } finally {
    ## Close loading dialog if it was shown
    if ($loadingDlg) { Close-LoadingDialog $loadingDlg }
  }
}

## Is it a special day...?
## Determines which emoji to use for the Active Directory window title. Uses Unicode escapes for flags to avoid editor issues.
function Initialise-DirectoryEmoji {
  param(
    [DateTime]$Date = (Get-Date)
  )

  $month = $Date.Month
  $day   = $Date.Day

  ## Default: card index
  $emoji = "🗂️"

  ## ==================== ICON INITIALISATION ====================

if ($Script:HasTerminalIcons) {

  ## Nerd Font / Symbol icons (explicit Unicode escapes)
  ## Safe for Terminal.Gui when font supports Nerd Fonts

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
    { $month -eq 1 -and $day -eq 1   }                   { $emoji = "📅" ; break }                    ## 1st Jan New Year’s Day
    { $month -eq 1  -and $day -eq 2  }                   { $emoji = "🦄" ; break }                    ## Wild haggis Hunting
    { $month -eq 4  -and $day -eq 9  }                   { $emoji = "`u{1F1E9}`u{1F1F0}" ; break }     ## 9th Apr Danmarks besættelse (liberation day)
    { $month -eq 5  -and $day -eq 4  }                   { $emoji = "🕯️" ; break }                    ## 4th May Candle for Besættelsen
    { $month -eq 6  -and $day -eq 5  }                   { $emoji = "`u{1F1E9}`u{1F1F0}" ; break }     ## 5th June Constitution day in Denmark
    { $month -eq 5  -and $day -eq 21 }                   { $emoji = "`u{1F1EC}`u{1F1F1}" ; break }     ## 21st May Grønland Day
    { $month -eq 7  -and $day -eq 1  }                   { $emoji = "🇨🇦" ; break }                     ## 1st Jul Canada Day
    { $month -eq 7  -and $day -eq 4  }                   { $emoji = "🫖" ; break }                    ## 4th Jul Teapot (to annoy Americans)
    { $month -eq 7  -and $day -eq 29 }                   { $emoji = "`u{1F1EB}`u{1F1F4}" ; break }     ## 29th Jul Faroe Islands
    { $month -eq 11  -and $day -eq 9 }                   { $emoji = "`u{1F1E9}`u{1F1EA}" ; break }     ## 9th Nov Erich Honecker leck mich am Arsch!
    { $month -eq 11 -and $day -eq 24 }                   { $emoji = "👑" ; break }                    ## 24th Nov (Så ta'r vi den en gang til for Prins Knud B-))
    { $month -eq 11 -and $day -eq 30 }                   { $emoji = "🏴󠁧󠁢󠁳󠁣󠁴󠁿" ; break }                    ## 30th Nov St Andrew’s Day (Saltire)
    { $month -eq 12 -and ($day -eq 24 -or $day -eq 25) } { $emoji = "🎄" ; break }                    ## 24th/25th Dec Tree for Jul / Christmas
  }

if ($Script:HasTerminalIcons) {
  # Nerd Font glyphs (Terminal.Gui-safe)
  $Script:DirectoryEmoji = ""  # nf-fa-folder (U+F115)
  $Script:FileEmoji      = ""  # nf-fa-file (U+F016)
  $Script:UserEmoji      = ""  # nf-fa-user (U+F007)
} else {
  # ASCII fallback
  $Script:DirectoryEmoji = "[D]"
  $Script:FileEmoji      = "[F]"
  $Script:UserEmoji      = "[U]"
}


  Debug-Log ("Today's emoji is: $emoji") -Type "info"
  $Script:DirectoryEmoji = $emoji
}

## ----------------------------{ Convert Domain Data }-----------------------
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
  function New-FakeSid { $rid = Get-Random -Minimum 1000 -Maximum 65535 ; "S-1-5-21-{0}-{1}-{2}-{3}" -f (Get-Random -Max 999999999), (Get-Random -Max 999999999), (Get-Random -Max 999999999), $rid }

  ## ========================= Users =========================
  $convertedUsers = @()
  foreach ($user in $Users) {
    $sam = ($user.Name -replace '\s+', '.').ToLower()
    $upn = if ($user.Email) { $user.Email } else { "$sam@$Domain" }
    $userDomain = if ($user.Email -and $user.Email -match '@(.+)$') { $matches[1] } else { $Domain }

    if ($user.OU) {
      $ouChain = $user.OU | ForEach-Object { "OU=$_" }
      [array]::Reverse($ouChain)
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
    $dcDomain = if ($dc.Domain) {
      $dc.Domain
    } elseif ($dc.Location -match 'Germany') { 'example.net'
    } else { $Domain
    }

    $dn = "CN=$($dc.Name),OU=Domain Controllers,$BaseDN"

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
  ## KEY FIX: Preserve ALL properties from source hashtable
  $convertedComputers = @()
  foreach ($computer in $Computers) {
    $dn = if ($computer.OU) {
      $ouChain = $computer.OU | ForEach-Object { "OU=$_" }
      "CN=$($computer.Name)," + ($ouChain[-1..0] -join ',') + ",$BaseDN"
    } else {
      "CN=$($computer.Name),CN=Computers,$BaseDN"
    }

    $compDomain = if ($computer.Domain) {
      $computer.Domain
    } elseif ($computer.Location -match 'Germany') {
      'example.net'
    } else {
      $Domain
    }

    ## START WITH BASE PROPERTIES
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

    ## Add ALL extra properties from source (including LAPS, BitLocker, etc.)
    foreach ($prop in $computer.Keys) {
      ## Skip properties we already set
      if ($prop -in @('Name','Domain','OS','OSVersion','IPv4Address','LastLogon','Location','Description','Type','OU','Enabled')) {
        continue
      }
      ## Add everything else (this captures LAPS properties)
      if ($null -ne $computer[$prop] -and $computer[$prop] -ne '') {
        $adComputer | Add-Member -NotePropertyName $prop -NotePropertyValue $computer[$prop] -Force
      }
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

  return @{
    Users     = $convertedUsers
    DCs       = $convertedDCs
    Groups    = $convertedGroups
    Computers = $convertedComputers
  }
}

##-------------------{ Convwert Object to Tree Items }-------------------
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

  Debug-Log ": Building content for domain: $domain" -Type "Info"

  # Filter users by domain and apply combined filters
  $nameFilter = $Script:FilterOptions.NameFilter.Trim()
  $domainUsers = $Script:Users | Where-Object {
    $_.Domain -eq $domain -or (Get-DomainFromDN $_.DistinguishedName) -eq $domain
  }

  ## Apply all filters using central function
  $filteredUsers = Apply-CombinedFilters -Users $domainUsers

  Debug-Log ": Filtered to $($filteredUsers.Count) users for domain $domain" -Type "Info"

  # Node cache for fast lookup
  $nodeCache = @{}

  function Get-OrCreateChildNode {
    param(
      [Terminal.Gui.Trees.TreeNode]$Parent,
      [string]$Name,
      [string]$FullPath,
      [string]$NodeType = 'container'
    )
    if ($nodeCache.ContainsKey($FullPath)) { return $nodeCache[$FullPath] }

    # Add icon based on node type
    $displayName = if ($NodeType -eq 'ou' -and $Script:Icons.OU) {
      "$($Script:Icons.OU) $Name"
    } else {
      $Name
    }

    $newNode = [Terminal.Gui.Trees.TreeNode]::new($displayName)

    # Look up backing object for OUs
    $backingObject = $null
    if ($NodeType -eq 'ou') {
      $backingObject = $Script:rawOUs | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
      if (-not $backingObject) {
        # Create a minimal OU object if not found in rawOUs
        $backingObject = @{
          Name = $Name
          Path = $FullPath
          Description = ""
        }
      }
    }

    $newNode.Tag = @{
      Type = $NodeType
      Object = $backingObject  # Now populated for OUs
    }

    $Parent.Children = $Parent.Children ?? (New-Object 'System.Collections.ObjectModel.Collection[Terminal.Gui.Trees.ITreeNode]')
    $Parent.Children.Add($newNode)
    $nodeCache[$FullPath] = $newNode
    return $newNode
  }

  # --- Build OU hierarchy ---
  foreach ($user in $filteredUsers) {
    $ouPath = Get-OUPathFromDN $user.DistinguishedName
    if ($ouPath.Count -eq 0) { continue }

    $currentNode = $domainNode
    $pathSoFar = ""
    foreach ($ouLevel in $ouPath) {
      $pathSoFar = if ($pathSoFar) { "$pathSoFar/$ouLevel" } else { $ouLevel }
      $currentNode = Get-OrCreateChildNode -Parent $currentNode -Name $ouLevel -FullPath "$domain/$pathSoFar" -NodeType 'ou'
    }

    # Determine user status icon
    $statusIcon = if ($user.LockedOut -or $user.Locked) {
      if ($Script:Icons.Locked) { $Script:Icons.Locked } else { "🔒" }
    }
    elseif ($user.Disabled -or -not $user.Enabled) {
      if ($Script:Icons.Disabled) { $Script:Icons.Disabled } else { "⊗" }
    }
    else {
      if ($Script:Icons.User) { $Script:Icons.User } else { "○" }
    }

    $userNode = [Terminal.Gui.Trees.TreeNode]::new("$statusIcon $($user.Name)")
    # CHANGED: Wrap the AD user object in a hashtable with Type
    $userNode.Tag = @{
      Type = 'user'
      Object = $user  # Real AD user object
    }
    $currentNode.Children = $currentNode.Children ?? (New-Object 'System.Collections.ObjectModel.Collection[Terminal.Gui.Trees.ITreeNode]')
    $currentNode.Children.Add($userNode)
  }

  # --- Groups ---
  ## FIXED: Always show Groups node if groups exist, regardless of filter settings
  $domainGroups = $Script:Groups | Where-Object { $_.Domain -eq $domain -or (Get-DomainFromDN $_.DistinguishedName) -eq $domain }
  if ($domainGroups.Count -gt 0) {
    Debug-Log ": Building Groups container with $($domainGroups.Count) groups" -Type "Info"

    $groupsNode = Get-OrCreateChildNode -Parent $domainNode -Name "Groups" -FullPath "$domain/_Groups" -NodeType 'container'

    foreach ($group in $domainGroups | Sort-Object Name) {
      $groupIcon = if ($Script:Icons.Group) { $Script:Icons.Group } else { "👥" }
      $groupNode = [Terminal.Gui.Trees.TreeNode]::new("$groupIcon $($group.Name)")
      # CHANGED: Wrap the AD group object in a hashtable with Type
      $groupNode.Tag = @{
        Type = 'group'
        Object = $group  # Real AD group object
      }
      $groupNode.Children = $groupNode.Children ?? (New-Object 'System.Collections.ObjectModel.Collection[Terminal.Gui.Trees.ITreeNode]')

      # Add group members (only if ShowGroups filter is enabled)
      if ($Script:FilterOptions.ShowGroups) {
        $members = $filteredUsers | Where-Object { $_.Groups -contains $group.Name -or $_.MemberOf -contains $group.DistinguishedName } | Sort-Object Name
        foreach ($member in $members) {
          $memberStatusIcon = if ($member.LockedOut -or $member.Locked) {
            if ($Script:Icons.Locked) { $Script:Icons.Locked } else { "🔒" }
          }
          elseif ($member.Disabled -or -not $member.Enabled) {
            if ($Script:Icons.Disabled) { $Script:Icons.Disabled } else { "⊗" }
          }
          else {
            if ($Script:Icons.User) { $Script:Icons.User } else { "○" }
          }
          $memberNode = [Terminal.Gui.Trees.TreeNode]::new("$memberStatusIcon $($member.Name)")
          # CHANGED: Wrap member (user) in hashtable
          $memberNode.Tag = @{
            Type = 'user'
            Object = $member  # Real AD user object
          }
          $groupNode.Children.Add($memberNode)
        }
      }

      ## Always add group node, even if it has no members
      $groupsNode.Children.Add($groupNode)
    }

    Debug-Log ": Added Groups container with $($domainGroups.Count) groups" -Type "Success"
  }

  # --- Domain Controllers ---
  ## FIXED: Always show DCs if they exist
  if ($Script:DCs -and $Script:DCs.Count -gt 0) {
    ## Filter DCs for current domain
    $dcsInDomain = $Script:DCs | Where-Object { $_.Domain -eq $domain }

    if ($dcsInDomain.Count -gt 0) {
      Debug-Log ": Building Domain Controllers container with $($dcsInDomain.Count) DCs" -Type "Info"

      ## Create "Domain Controllers" pseudo-group node
      $dcContainerNode = [Terminal.Gui.Trees.TreeNode]::new()
      $dcContainerNode.Text = "🖥️  Domain Controllers ($($dcsInDomain.Count))"
      $dcContainerNode.Tag = @{
        Type = 'dc-container'
        Name = 'Domain Controllers'
        Object = $null  # No backing object for container
      }

      ## Initialise children collection
      $dcContainerNode.Children = $dcContainerNode.Children ?? (New-Object 'System.Collections.ObjectModel.Collection[Terminal.Gui.Trees.ITreeNode]')

      ## Add individual DCs under container
      foreach ($dc in ($dcsInDomain | Sort-Object Name)) {
        $dcNode = [Terminal.Gui.Trees.TreeNode]::new()

        ## Build display text with status indicators
        $statusIcon = if ($dc.Enabled) { "✓" } else { "⊗" }
        $gcIcon = if ($dc.IsGlobalCatalog) { "🌐" } else { "" }
        $healthIcon = switch ($dc.ReplicationHealth) {
          'Healthy' { "✓" }
          { $_ -match 'Warning' } { "▲" }
          { $_ -match 'Critical|Error|Failed' } { "✗" }
          default { "?" }
        }

        ## Show FSMO roles if any
        $fsmoText = if ($dc.FSMORoles -and $dc.FSMORoles.Count -gt 0) {
          " [FSMO: $($dc.FSMORoles.Count)]"
        } else {
          ""
        }

        $dcNode.Text = "$statusIcon $gcIcon $($dc.Name) - $($dc.Site)$fsmoText $healthIcon"
        $dcNode.Tag = @{
          Type = 'dc'
          Name = $dc.Name
          Object = $dc
        }

        $dcContainerNode.Children.Add($dcNode)
      }

      ## Add DC container to domain node
      $domainNode.Children = $domainNode.Children ?? (New-Object 'System.Collections.ObjectModel.Collection[Terminal.Gui.Trees.ITreeNode]')
      $domainNode.Children.Add($dcContainerNode)

      Debug-Log ": Added Domain Controllers container with $($dcsInDomain.Count) DCs" -Type "Success"
    }
  }

  # --- Computers ---
  ## FIXED: Always show Computers if they exist
  $domainComputers = $Script:Computers | Where-Object { $_.Domain -eq $domain -or (Get-DomainFromDN $_.DistinguishedName) -eq $domain }
  if ($domainComputers.Count -gt 0) {
    Debug-Log ": Building Computers container with $($domainComputers.Count) computers" -Type "Info"

    $computerNode = Get-OrCreateChildNode -Parent $domainNode -Name "Computers" -FullPath "$domain/_Computers" -NodeType 'container'
    $byType = $domainComputers | Group-Object ComputerType

    foreach ($typeGroup in $byType) {
      $typeNode = [Terminal.Gui.Trees.TreeNode]::new("$($typeGroup.Name) ($($typeGroup.Count))")
      # Tag for computer type containers
      $typeNode.Tag = @{
        Type = 'container'
        Object = $null
      }
      $typeNode.Children = $typeNode.Children ?? (New-Object 'System.Collections.ObjectModel.Collection[Terminal.Gui.Trees.ITreeNode]')

      foreach ($comp in $typeGroup.Group | Sort-Object Name) {
        $compIcon = if ($comp.Disabled -or -not $comp.Enabled) {
          if ($Script:Icons.Disabled) { $Script:Icons.Disabled } else { "⊗" }
        } else {
          if ($Script:Icons.Computer) { $Script:Icons.Computer } else { "💻" }
        }
        $compNode = [Terminal.Gui.Trees.TreeNode]::new("$compIcon $($comp.Name)")
        # CHANGED: Wrap computer object in hashtable
        $compNode.Tag = @{
          Type = 'computer'
          Object = $comp  # Real AD computer object
        }
        $typeNode.Children.Add($compNode)
      }

      if ($typeNode.Children.Count -gt 0) {
        $computerNode.Children.Add($typeNode)
      }
    }

    Debug-Log ": Added Computers container with $($domainComputers.Count) computers" -Type "Success"
  }

  Debug-Log ": Finished building content for domain $domain" -Type "Success"
}

##-------------------------{ Build The Tree }-------------------------
function Build-Tree {
  param([string]$domain)
  if (-not $domain) { $domain = $Script:CurrentDomain }
  Debug-Log ": Building tree..." -Type "Info"

  ## TreeView MUST already exist in 1.16
  if ($null -eq $Script:tree) {
    throw "Build-Tree failed: TreeView does not exist"
  }

  ## Clear existing objects safely
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

## -----------------------{ Filter Label Function }----------------------
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
      Debug-Log ": Creating filter status label" -Type "Info"

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

        Debug-Log ": Filter status label created successfully" -Type "Success"
        return $lblStatus

      } catch {
        Debug-Log ": ERROR creating filter status label: $($_.Exception.Message)" -Type "Error"
        return $null
      }
    }

  if ($Action -eq 'Update') {
    if (-not $Label) {
      Debug-Log ": Label parameter is null in Update action" -Type "Warn"
      return
    }

    Debug-Log ": Updating filter status label" -Type "Info"

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
    Debug-Log ": Filter status label updated: $($activeFilters.Count) filters active" -Type "Info"
  }
}

## ==================== LDAP FILTER HELPER ====================
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

    ## ==================== USER FILTERS ====================
    [Parameter(Mandatory=$true, ParameterSetName='User')]
    [ValidateSet(
      'All',
      'LockedOnly',
      'DisabledOnly',
      'EnabledOnly',
      'NeverLoggedIn',
      'NoManager',
      'PasswordExpired',
      'PasswordExpiring72h',
      'PasswordNeverExpires',
      'AccountExpired',
      'AccountExpiring30d',
      'StaleAccounts90d',
      'EmptyEmail',
      'EmptyDepartment'
    )]
    [string]$FilterType,

    [Parameter(ParameterSetName='User')]
    [array]$Users = $Script:Users,

    ## ==================== GROUP FILTERS ====================
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
          return $Users | Where-Object {
            $_.LastLogonDate -and $_.LastLogonDate -lt $staleDate
          }
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

## Info panel to reduce clutter. Call once to create, then again to update, e.g. if domain changes
function Show-InfoPanel {
  <#
  .SYNOPSIS
  Create or update the Environment Info panel
  .DESCRIPTION
  - First call: creates and returns the InfoPanel FrameView
  - Subsequent calls: updates label contents in-place
  - Safe to call after domain, forest, DC, or theme changes
  #>
  param(
    [int]$PanelWidth  = 40,
    [int]$PanelHeight = 10,
    [switch]$UpdateOnly
  )

  ## ==================== CREATE (first run) ====================
  if (-not $Script:InfoPanel) {
    $infoPanel = [Terminal.Gui.FrameView]::new("Environment Info")
    $infoPanel.Width  = $PanelWidth
    $infoPanel.Height = $PanelHeight
    $infoPanel.X = [Terminal.Gui.Pos]::AnchorEnd($PanelWidth)
    $infoPanel.Y = 1
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
    $yPos++  ## spacing

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

    ## Cache panel globally
    $Script:InfoPanel = $infoPanel
  }

  ## ==================== UPDATE (every call) ====================
  $labels = $Script:InfoPanel.Tag

  ## Get current DC name (handle both string and object)
  $dcName = if ($Script:CurrentDC) {
    if ($Script:CurrentDC -is [string]) {
      $Script:CurrentDC
    } else {
      $Script:CurrentDC.Name
    }
  } else {
    "(None)"
  }

  $labels.ForestLabel.Text       = [NStack.ustring]::Make("Forest:     $($Script:ForestName)")
  $labels.DomainLabel.Text       = [NStack.ustring]::Make("Domain:     $($Script:CurrentDomain)")
  $labels.DCLabel.Text           = [NStack.ustring]::Make("Current DC: $dcName") ## $($Script:CurrentDC is an array
  $labels.UsersLabel.Text        = [NStack.ustring]::Make("Users:      $($Script:Users.Count)")
  $labels.GroupsLabel.Text       = [NStack.ustring]::Make("Groups:     $($Script:Groups.Count)")
  $labels.ComputersLabel.Text    = [NStack.ustring]::Make("Computers:  $($Script:Computers.Count)")
  $labels.DCsLabel.Text          = [NStack.ustring]::Make("DCs:        $($Script:DCs.Count)")
  $labels.TotalObjectsLabel.Text = [NStack.ustring]::Make("Objects:    $($Script:ADObjects.Count)")
  $labels.ThemeLabel.Text        = [NStack.ustring]::Make("Theme:      $($Script:ThemeMode)")

  ## Force visual update
  $Script:InfoPanel.SetNeedsDisplay()
  ## If UpdateOnly, also refresh the whole application
  if ($UpdateOnly) { [Terminal.Gui.Application]::Refresh() }
  return $Script:InfoPanel
}

## ------------------------- Filter Panel (Add to main window)-------------------------
function Create-FilterPanel {

  ## Initialise FilterOptions with new fields
  if (-not $Script:FilterOptions) {
    Debug-Log "FilterOptions not initialised — initialising defaults" -Type "Warn"
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

  ## Create frame
  $filterFrame = [Terminal.Gui.FrameView]::new("Filters")
  $filterFrame.X = 28
  $filterFrame.Y = 1
  $filterFrame.Width = 40
  $filterFrame.Height = 26  # Increased for separator lines and better spacing
  $y = 0

  ## ==================== Name Filter with Operator ====================
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
  $cmbOperator.add_SelectedItemChanged({  $Script:FilterOptions.NameOperator = $cmbOperator.Text.ToString() }.GetNewClosure())
  $filterFrame.Add($cmbOperator)
  $y+=2

  ## Text field
  $txtNameFilter = [Terminal.Gui.TextField]::new($Script:FilterOptions.NameFilter)
  $txtNameFilter.X=1; $txtNameFilter.Y=$y; $txtNameFilter.Width=35
  $txtNameFilter.add_TextChanged({
    $Script:FilterOptions.NameFilter = $txtNameFilter.Text.ToString()
  })
  $filterFrame.Add($txtNameFilter)
  $y+=1

  ## Separator line (visual spacing)
  $lblSep1 = [Terminal.Gui.Label]::new("─────────────────────────────────────")
  $lblSep1.X=1; $lblSep1.Y=$y
  $filterFrame.Add($lblSep1)
  $y+=1

  ## ==================== Quick Filters Dropdown ====================
  $lblQuickFilter = [Terminal.Gui.Label]::new("Quick Filter:")
  $lblQuickFilter.X=1; $lblQuickFilter.Y=$y
  $filterFrame.Add($lblQuickFilter)
  $y+=1

  $cmbQuickFilter = [Terminal.Gui.ComboBox]::new()
  $cmbQuickFilter.X=1; $cmbQuickFilter.Y=$y; $cmbQuickFilter.Width=35

  $quickFilters = @(
    "All",
    "LockedOnly",
    "DisabledOnly",
    "EnabledOnly",
    "NeverLoggedIn",
    "NoManager",
    "PasswordExpired",
    "PasswordExpiring72h",
    "PasswordNeverExpires",
    "AccountExpired",
    "AccountExpiring30d",
    "StaleAccounts90d",
    "EmptyEmail",
    "EmptyDepartment"
  )

  ## Friendly display names
  $quickFilterDisplay = @(
    "All Users",
    "Locked Accounts",
    "Disabled Accounts",
    "Enabled Accounts",
    "Never Logged In",
    "No Manager Assigned",
    "Password Expired",
    "Password Expiring (72h)",
    "Password Never Expires",
    "Account Expired",
    "Account Expiring (30d)",
    "Stale Accounts (90d+)",
    "No Email Address",
    "No Department"
  )

  $cmbQuickFilter.SetSource($quickFilterDisplay)
  $cmbQuickFilter.SelectedItem = 0
  $cmbQuickFilter.add_SelectedItemChanged({
    if ($cmbQuickFilter.SelectedItem -ge 0 -and $cmbQuickFilter.SelectedItem -lt $quickFilters.Count) {
      $Script:FilterOptions.QuickFilter = $quickFilters[$cmbQuickFilter.SelectedItem]
      Debug-Log ": Quick filter changed to: $($Script:FilterOptions.QuickFilter)" -Type "Info"
    }
  })
  $filterFrame.Add($cmbQuickFilter)
  $y+=1

  ## Separator line (visual spacing)
  $lblSep2 = [Terminal.Gui.Label]::new("─────────────────────────────────────")
  $lblSep2.X=1; $lblSep2.Y=$y
  $filterFrame.Add($lblSep2)
  $y+=1

  ## ==================== Show/Hide Checkboxes ====================
  $chkEnabled = [Terminal.Gui.CheckBox]::new("Enabled Users")
  $chkEnabled.X=1; $chkEnabled.Y=$y; $chkEnabled.Checked=$Script:FilterOptions.ShowEnabledUsers
  $chkEnabled.add_Toggled({ $Script:FilterOptions.ShowEnabledUsers = $chkEnabled.Checked })
  $filterFrame.Add($chkEnabled)
  $y+=1

  $chkLocked = [Terminal.Gui.CheckBox]::new("Locked Users")
  $chkLocked.X=1; $chkLocked.Y=$y; $chkLocked.Checked=$Script:FilterOptions.ShowLockedUsers
  $chkLocked.add_Toggled({ $Script:FilterOptions.ShowLockedUsers = $chkLocked.Checked })
  $filterFrame.Add($chkLocked)
  $y+=1

  $chkDisabled = [Terminal.Gui.CheckBox]::new("Disabled Users")
  $chkDisabled.X=1; $chkDisabled.Y=$y; $chkDisabled.Checked=$Script:FilterOptions.ShowDisabledUsers
  $chkDisabled.add_Toggled({ $Script:FilterOptions.ShowDisabledUsers = $chkDisabled.Checked })
  $filterFrame.Add($chkDisabled)
  $y+=1

  $chkPwdExpiring72h = [Terminal.Gui.CheckBox]::new("Passwords Expiring in next 72h")
  $chkPwdExpiring72h.X = 1
  $chkPwdExpiring72h.Y = $y
  $chkPwdExpiring72h.Checked = $Script:FilterOptions.ShowPasswordExpiring72h
  $chkPwdExpiring72h.add_Toggled({ $Script:FilterOptions.ShowPasswordExpiring72h = $chkPwdExpiring72h.Checked })
  $filterFrame.Add($chkPwdExpiring72h)
  $y += 1

  $chkPwdExpired = [Terminal.Gui.CheckBox]::new("Users with expired passwords")
  $chkPwdExpired.X = 1
  $chkPwdExpired.Y = $y
  $chkPwdExpired.Checked = $Script:FilterOptions.ShowPasswordExpired
  $chkPwdExpired.add_Toggled({ $Script:FilterOptions.ShowPasswordExpired = $chkPwdExpired.Checked })
  $filterFrame.Add($chkPwdExpired)
  $y += 1

  $chkNoGroups = [Terminal.Gui.CheckBox]::new("Users Not in Any Groups")
  $chkNoGroups.X = 1
  $chkNoGroups.Y = $y
  $chkNoGroups.Checked = $Script:FilterOptions.ShowUsersNoGroups
  $chkNoGroups.add_Toggled({ $Script:FilterOptions.ShowUsersNoGroups = $chkNoGroups.Checked })
  $filterFrame.Add($chkNoGroups)
  $y += 1

  $chkGroups = [Terminal.Gui.CheckBox]::new("Groups")
  $chkGroups.X=1; $chkGroups.Y=$y; $chkGroups.Checked=$Script:FilterOptions.ShowGroups
  $chkGroups.add_Toggled({ $Script:FilterOptions.ShowGroups = $chkGroups.Checked })
  $filterFrame.Add($chkGroups)
  $y+=1

  $chkOUs = [Terminal.Gui.CheckBox]::new("OUs")
  $chkOUs.X=1; $chkOUs.Y=$y; $chkOUs.Checked=$Script:FilterOptions.ShowOUs
  $chkOUs.add_Toggled({ $Script:FilterOptions.ShowOUs = $chkOUs.Checked })
  $filterFrame.Add($chkOUs)
  $y+=1

  $chkDCs = [Terminal.Gui.CheckBox]::new("Domain Controllers")
  $chkDCs.X=1; $chkDCs.Y=$y; $chkDCs.Checked=$Script:FilterOptions.ShowDCs
  $chkDCs.add_Toggled({ $Script:FilterOptions.ShowDCs = $chkDCs.Checked })
  $filterFrame.Add($chkDCs)
  $y+=1

  $chkComputers = [Terminal.Gui.CheckBox]::new("Computers")
  $chkComputers.X=1; $chkComputers.Y=$y; $chkComputers.Checked=$Script:FilterOptions.ShowComputers
  $chkComputers.add_Toggled({ $Script:FilterOptions.ShowComputers = $chkComputers.Checked })
  $filterFrame.Add($chkComputers)
  $y+=1

  $chkNoLAPS = [Terminal.Gui.CheckBox]::new("Devices Without LAPS")
  $chkNoLAPS.X = 1
  $chkNoLAPS.Y = $y
  $chkNoLAPS.Checked = $Script:FilterOptions.ShowDevicesNoLAPS
  $chkNoLAPS.add_Toggled({ $Script:FilterOptions.ShowDevicesNoLAPS = $chkNoLAPS.Checked })
  $filterFrame.Add($chkNoLAPS)
  $y += 1

  $chkNoBitlocker = [Terminal.Gui.CheckBox]::new("Devices Without BitLocker Keys")
  $chkNoBitlocker.X = 1
  $chkNoBitlocker.Y = $y
  $chkNoBitlocker.Checked = $Script:FilterOptions.ShowDevicesNoBitlocker
  $chkNoBitlocker.add_Toggled({ $Script:FilterOptions.ShowDevicesNoBitlocker = $chkNoBitlocker.Checked })
  $filterFrame.Add($chkNoBitlocker)
  $y += 2

  ## Separator line (visual spacing)
  $lblSep3 = [Terminal.Gui.Label]::new("─────────────────────────────────────")
  $lblSep3.X=1; $lblSep3.Y=$y
  $filterFrame.Add($lblSep3)
  $y+=1

  ## ==================== Apply/Reset Buttons ====================
  $btnApplyFilter = [Terminal.Gui.Button]::new("Apply Filter")
  $btnApplyFilter.X=1; $btnApplyFilter.Y=$y

  ## BEFORE calling .add_Clicked(), verify button exists
  $btnApplyFilter = [Terminal.Gui.Button]::new("Apply")
  $btnApplyFilter.X = 2
  $btnApplyFilter.Y = $y

  ## ADD THIS CHECK:
  if (-not $btnApplyFilter) {
    Debug-Log ": FATAL - btnApplyFilter is null!" -Type "Error"
    return
  }

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
    Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
  }).GetNewClosure()
  $filterFrame.Add($btnApplyFilter)

  $btnResetFilter = [Terminal.Gui.Button]::new("Reset")
  $btnResetFilter.X=21; $btnResetFilter.Y=$y
  $btnResetFilter.add_Clicked({
    Debug-Log "Resetting filters..." -Type "Info"

    ## Reset all filters - Add any new ones such as bitlocker keys here
    $Script:FilterOptions.ShowDisabledUsers      = $true
    $Script:FilterOptions.ShowEnabledUsers       = $true
    $Script:FilterOptions.ShowLockedUsers        = $true
    $Script:FilterOptions.ShowGroups             = $true
    $Script:FilterOptions.ShowDCs                = $true
    $Script:FilterOptions.ShowComputers          = $true
    $Script:FilterOptions.ShowOUs                = $true
    $Script:FilterOptions.ShowDevicesNoBitlocker = $true
    $Script:FilterOptions.ShowDevicesNoLAPS      = $true
    $Script:FilterOptions.ShowUsersNoGroups      = $true
    $Script:FilterOptions.NameFilter             = ""
    $Script:FilterOptions.NameOperator           = "Contains"
    $Script:FilterOptions.QuickFilter            = "All"
    $Script:FilterOptions.SortBy                 = "Name"
    $Script:FilterOptions.SortDescending         = $false

    ## Reset UI controls - any you add above need added here too
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
    $txtNameFilter.Text          = ""
    $cmbOperator.SelectedItem    = 0
    $cmbQuickFilter.SelectedItem = 0

    ## Rebuild tree
    $rootNode = Build-Tree -domain $Script:CurrentDomain
    if ($rootNode) {
      $Script:tree.ClearObjects()
      $Script:tree.AddObject($rootNode)
      [Terminal.Gui.Application]::Refresh()
    }

    ## Update status label
    Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
  }).GetNewClosure()
  $filterFrame.Add($btnResetFilter)
  $y+=2

  ## ==================== Filter Status Label ====================
  $Script:FilterStatusLabel = Manage-FilterStatusLabel -Action 'Create' -X 1 -Y $y -InPanel
  $filterFrame.Add($Script:FilterStatusLabel)

  Debug-Log "Enhanced FilterPanel created with LDAP-style filters" -Type "Info"

  ## Apply theme
  if ($Script:themeData -and $Script:themeData.MainWindow) {
    $filterFrame.ColorScheme = $Script:themeData.MainWindow
  }

  return $filterFrame
}

## -------------------{ Quick Filter Menu (for Menu Bar) -------------------
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
      Debug-Log ": Applying quick filter: $($selected.Name)" -Type "Info"

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

## ==================== MAIN GET-ADHEALTH FUNCTION ====================
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
  - Group Policy Objects

  .PARAMETER Domain
  The domain to check. Defaults to $Script:CurrentDomain

  .EXAMPLE
  Show-ADHealthDialog

  .EXAMPLE
  Show-ADHealthDialog -Domain "contoso.com"
  #>

  [CmdletBinding()]
  param(
    [Parameter(Position=0)]
    [string]$Domain
  )

  Debug-Log ": Show-ADHealthDialog called" -Type "Info"

  ## ==================== OS & TOOL DETECTION ====================

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

  Debug-Log ": Detected OS: $($osInfo.OSName)" -Type "Info"

  ## Check tool availability
  $tools = Test-ToolsAvailability
  Debug-Log ": Tool availability checked" -Type "Info"

  ## Check AD module
  $hasADModule = $null -ne ($Script:HasActiveDirectory)
  Debug-Log ": ActiveDirectory module available: $hasADModule" -Type "Info"

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

  Debug-Log ": Checking AD Health for domain: $Domain" -Type "Info"

  ## Store tools and OS info in script scope for helper functions
  $Script:ADHealthTools = $tools
  $Script:ADHealthOS = $osInfo
  $Script:ADHealthDomain = $Domain

  ## Create main dialog
  $dialog = [Terminal.Gui.Dialog]::new("AD Health Check - $Domain", 120, 35)

  ## Create TabView
  $tabView = [Terminal.Gui.TabView]::new()
  $tabView.X = 0
  $tabView.Y = 0
  $tabView.Width = [Terminal.Gui.Dim]::Fill()
  $tabView.Height = [Terminal.Gui.Dim]::Fill(2)

  ## Tab definitions for AD health check
  $tabs = @(
    @{ Name = "System Info"; Generator = { Get-SystemInfoText } }
    @{ Name = "Domain Controllers"; Generator = { Get-DCStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "Replication"; Generator = { Get-ReplicationStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "DNS Records"; Generator = { Get-DNSStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "SYSVOL/NETLOGON"; Generator = { Get-SysvolStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "FSMO Roles"; Generator = { Get-FSMOStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "DFS Status"; Generator = { Get-DFSStatusText -Domain $Script:ADHealthDomain } }
    @{ Name = "Group Policy"; Generator = { Get-GPOStatusText -Domain $Script:ADHealthDomain } }
  )

  ## Create tabs
  foreach ($tabDef in $tabs) {
    $tab = [Terminal.Gui.TabView+Tab]::new()
    $tab.Text = [NStack.ustring]::Make($tabDef.Name)
    $tab.View = [Terminal.Gui.View]::new()
    $tab.View.Width = [Terminal.Gui.Dim]::Fill()
    $tab.View.Height = [Terminal.Gui.Dim]::Fill()

    $text = & $tabDef.Generator
    $textView = [Terminal.Gui.TextView]::new()
    $textView.X = 1
    $textView.Y = 0
    $textView.Width = [Terminal.Gui.Dim]::Fill(1)
    $textView.Height = [Terminal.Gui.Dim]::Fill()
    $textView.ReadOnly = $true
    $textView.Text = [NStack.ustring]::Make($text)
    $tab.View.Add($textView)

    $tabView.AddTab($tab, $false)
  }

  $dialog.Add($tabView)

  ## ==================== CTRL+F SEARCH ====================

  ## Add key handler for Ctrl+F
  $dialog.add_KeyPress({
    param($e)

    if ($e.KeyEvent.Key -eq [Terminal.Gui.Key]::CtrlMask -bor [Terminal.Gui.Key]::f) {
      $currentTab = $tabView.SelectedTab
      $textView = $currentTab.View.Subviews[0]

      if ($textView -and $textView -is [Terminal.Gui.TextView]) {
        ## Simple search prompt
        $searchText = Show-InputDialog -Title "Search" -Label "Search for:" -Width 50

        if ($searchText) {
          $content = $textView.Text.ToString()
          $index = $content.IndexOf($searchText, [StringComparison]::OrdinalIgnoreCase)

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
  })

  ## ==================== BUTTONS ====================

  $btnRefresh = [Terminal.Gui.Button]::new("Refresh")
  $btnRefresh.X = 2
  $btnRefresh.Y = [Terminal.Gui.Pos]::AnchorEnd(1)
  $btnRefresh.add_Clicked({
    Debug-Log ": Refreshing current tab..." -Type "Info"
    $currentTab = $tabView.SelectedTab
    $tabName = $currentTab.Text.ToString()

    ## Find matching tab definition
    $tabDef = $tabs | Where-Object { $_.Name -eq $tabName } | Select-Object -First 1
    if ($tabDef) {
      ## Refresh tools check if System Info
      if ($tabName -eq "System Info") {
        $Script:ADHealthTools = Test-ToolsAvailability
      }

      $newText = & $tabDef.Generator
      $currentTab.View.Subviews[0].Text = [NStack.ustring]::Make($newText)
    }
  }).GetNewClosure()
  $dialog.Add($btnRefresh)

  $btnExport = [Terminal.Gui.Button]::new("Export Report")
  $btnExport.X = 15
  $btnExport.Y = [Terminal.Gui.Pos]::AnchorEnd(1)
  $btnExport.add_Clicked({
    Debug-Log ": Exporting AD health report..." -Type "Info"

    $report = @"
AD HEALTH REPORT
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Domain: $Domain
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

SYSTEM INFORMATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
$(Get-SystemInfoText)

DOMAIN CONTROLLERS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
$(Get-DCStatusText -Domain $Domain)

REPLICATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
$(Get-ReplicationStatusText -Domain $Domain)

DNS RECORDS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
$(Get-DNSStatusText -Domain $Domain)

SYSVOL/NETLOGON
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
$(Get-SysvolStatusText -Domain $Domain)

FSMO ROLES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
$(Get-FSMOStatusText -Domain $Domain)

GROUP POLICY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
$(Get-GPOStatusText -Domain $Domain)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
End of Report
"@

  $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $savePath = Show-FileBrowserDialog -StartDir "." -Title "Save AD Health Report" -Filter @("*.txt", "*.log")
  if ($savePath) {
    try {
      $report | Out-File -FilePath $savePath -Encoding UTF8
      Debug-Log ": Report exported to $savePath" -Type "Success"
      Show-Modal "Export Complete" "AD Health Report saved to:`n`n$savePath"
    } catch {
      Debug-Log ": Failed to export report: $($_.Exception.Message)" -Type "Error"
      Show-Modal "Export Failed" "Failed to save report:`n`n$($_.Exception.Message)"
    }
  }
}).GetNewClosure()

  $dialog.Add($btnExport)

  $btnClose = [Terminal.Gui.Button]::new("Close")
  $btnClose.X = [Terminal.Gui.Pos]::AnchorEnd(10)
  $btnClose.Y = [Terminal.Gui.Pos]::AnchorEnd(1)
  $btnClose.add_Clicked({
    Debug-Log ": AD Health dialog closed" -Type "Info"
    [Terminal.Gui.Application]::RequestStop()
  }).GetNewClosure()
  $dialog.Add($btnClose)

  ## Run dialog
  Debug-Log ": Running AD Health dialog" -Type "Info"
  [Terminal.Gui.Application]::Run($dialog)
}


## ==================== CONTENT GENERATORS ====================

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
  $output += "  Terminal-Icons:                       $(if ($hasTerminalIcons) { '✓ Available' } else { '✗ Not Found' })"
  $output += "  NerdFonts:                            $(if ($Script:hasNerdFonts) { '✓ Available' } else { '✗ Not Found' })"
  $output += "  PSWriteColor:                         $(if ($Script:hasPSWriteColor) { '✓ Available' } else { '✗ Not Found' })"
  $output += "  Microsoft.PowerShell.ConsoleGuiTools: $(if ($Script:hasConsoleTools) { '✓ Available' } else { '✗ Not Found' })"
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

  ## -------------------- Group Policy Tools Detection --------------------
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
        Debug-Log "Unknown Windows edition. Cannot suggest Group Policy installation." -Type "Warn"
      }
    }
  } else {
    Debug-Log "Group Policy tools are Windows-only. Skipping check for $($PSVersionTable.OS)." -Type "Info"
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
        $smb = "✓ OK"
        $dns = "✓ OK"

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
    ## Production mode - check ports
    $result = Test-DCStatus -Domain $Domain

    if ($result.Summary) {
      $fmt = "{0,-20} {1,-20} {2,-16} {3,-10} {4,-6} {5,-6} {6,-6} {7,-6}"
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
      $totalCount = $result.Summary.Count

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
          if ($sourceDC.Name -ne $targetDC.Name) {
            $output += $fmt -f $sourceDC.Name, $targetDC.Name, "✓ OK", (Get-Date -Format "yyyy-MM-dd HH:mm")
          }
        }
      }

      $output += ""
      $output += "All replication partners healthy."
      $output += "No errors detected."
    } else {
      $output += "Single DC - no replication partners."
    }
  } else {
    ## Production mode - actually run repadmin
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
      foreach ($dc in $dcs) {
        $output += "  $($dc.Name): ✓ Running"
      }
    }

    $output += ""
    $output += "All DNS records appear healthy."
  } else {
    ## Production mode - actually query DNS
    $result = Test-ADDnsRecords -Domain $Domain

    $output += "Querying DNS records for domain: $Domain"
    $output += ""

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
    ## Production mode - actually check shares
    $result = Test-SysvolHealth -Domain $Domain

    $output += "Checking SYSVOL/NETLOGON shares for domain: $Domain"
    $output += ""

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


## New tab
# ===================== DFS Tab for AD Health Check =====================
# Add this to your Show-ADHealthDialog tabs array
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

        if ($link.LinkPath) {
          $output += "    Path: $($link.LinkPath)"
        }

        if ($link.TargetList) {
          $targets = $link.TargetList -split ';' | Select-Object -First 3
          $output += "    Targets:"
          foreach ($target in $targets) {
            if ($target) {
              $output += "      - $target"
            }
          }
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
      foreach ($group in $groupedDFSR) {
        $output += "  • $($group.Name): $($group.Count) object(s)"
      }

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
    ## Production mode - actually query FSMO
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

## This one is pre-merge. I believe it now needs rewritten to be msaller/cleaner
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

    foreach ($gpo in $demoGPOs) {
      $output += $fmt -f $gpo.Name, $gpo.Status
    }

    $output += ""
    $output += "All GPOs in sync."
  } else {
    ## Production mode - actually query GPOs
    $result = Test-GPOHealth -Domain $Domain

    $output += "Checking GPOs for domain: $Domain"
    $output += ""

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

## ==================== PRODUCTION CHECK FUNCTIONS ====================
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
  LINUX-BASED TOOLS FOR QUERYING WINDOWS / ACTIVE DIRECTORY
  =========================================================

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

----------{ SUMMARY }----------
Windows diagnostic tools tend to be monolithic.
Linux diagnostics are compositional, combining DNS, LDAP,
Kerberos, SMB, and RPC checks to reach the same conclusions.

This block documents the conceptual and operational mapping.

MACOS / HOMEBREW AVAILABILITY NOTES
===================================

Most Linux-side AD diagnostic tools are available on macOS via Homebrew.
However, a few are LIMITED or NOT FUNCTIONALLY EQUIVALENT on macOS.

----------{ AVAILABLE & FUNCTIONALLY EQUIVALENT ON macOS (via Homebrew) }----------
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

----------{ PARTIALLY AVAILABLE / CONDITIONAL }----------
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

----------{ LIMITED OR NOT PRACTICAL ON macOS }----------
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

----------{ SUMMARY }----------
✔ Most LDAP, DNS, Kerberos, and port-query tooling works identically
  on Linux and macOS.

⚠ Samba-based trust, DRS, and RPC tooling is the weakest area on macOS
  due to build flags, sandboxing, and Apple SMB stack differences.

❌ DFS-R diagnostics remain Windows-only regardless of platform.

This distinction matters when documenting "Linux/macOS from Windows" diagnostic parity.
#>

  ## Tool matrix per OS
  $toolMatrix = @{
    Windows = @('repadmin.exe', 'dfsrdiag.exe', 'dcdiag.exe', 'nltest.exe', 'csvde.exe', 'portqry.exe', 'ping.exe', 'netstat.exe', 'nslookup.exe')
    Linux   = @('dig', 'ldapsearch', 'ldapwhoami', 'kinit', 'klist', 'nmap', 'nc', 'netstat', 'nslookup', 'ping', 'rpcclient', 'smbclient', 'samba-tool', 'wbinfo')
    MacOS   = @('dig', 'kinit', 'klist', 'ldapsearch', 'ldapwhoami', 'nc', 'netstat', 'nmap', 'nslookup', 'ping', 'smbclient')
  }

  <#
    NB: Samba-based tools (samba-tool, wbinfo, rpcclient) are excluded because Homebrew Samba builds on macOS often lack full AD/DC
    features, conflict with Apple’s native SMB stack, or behave inconsistently across macOS releases. LDAP, DNS,  Kerberos, SMB
    client access, and port-scanning tools are stable and fully supported.
  #>


  $tools = @{}
  foreach ($tool in $toolMatrix[$os]) {
    $tools[$tool] = [bool](Get-Command $tool -ErrorAction SilentlyContinue)
  }

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

##  TODO: Never called. Perhaps this is a dupe, or an improved version
function New-ToolsAvailabilityPanel {
  param(
    [int]$X,
    [int]$Y,
    [int]$Width,
    [int]$Height
  )

  ## Detect OS for radio default
  $DetectedOS =
    if ($IsWindows) { 'Windows' }
    elseif ($IsLinux) { 'Linux' }
    elseif ($IsMacOS) { 'MacOS' }
    else { 'Windows' }

    $frame = [Terminal.Gui.FrameView]::new(
      "Tools Availability",
      $X,
      $Y,
      $Width,
      $Height
    )

    ## OS selection (non-modal)
    $osRadio = [Terminal.Gui.RadioGroup]::new(
      1,
      0,
      @('Windows','Linux','MacOS')
    )
    $osRadio.SelectedItem =
      @('Windows','Linux','MacOS').IndexOf($DetectedOS)
    $frame.Add($osRadio)

    ## Tool name column
    $toolNames = [Terminal.Gui.ListView]::new(
      1,
      4,
      22,
      -1
    )
    $frame.Add($toolNames)

    ## Status column
    $toolStatus = [Terminal.Gui.ListView]::new(
      24,
      4,
      5,
      -1
    )
    $frame.Add($toolStatus)

    ## Colour schemes
    $greenScheme = [Terminal.Gui.ColorScheme]::new()
    $greenScheme.Normal = [Terminal.Gui.Attribute]::Make(
      [Terminal.Gui.Color]::BrightGreen,
      [Terminal.Gui.Color]::Black
    )

    $redScheme = [Terminal.Gui.ColorScheme]::new()
    $redScheme.Normal = [Terminal.Gui.Attribute]::Make(
      [Terminal.Gui.Color]::BrightRed,
      [Terminal.Gui.Color]::Black
    )

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

      $names  = @()
      $status = @()

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

    $osRadio.add_SelectedItemChanged({
      Refresh-Tools
    })

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
      $dc = Get-ADDomainController -Identity $dc.HostName -Server $Domain
      $name = $dc.HostName
      $ip = if ($dc.IPv4Address) { $dc.IPv4Address } else { "N/A" }
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
      $smbOK = $false
      $dnsOK = $false

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
        $svcDNS = (Get-Service -ComputerName $name -Name "DNS" -ErrorAction SilentlyContinue).Status
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
    $health = "FAIL"
  }

  return @{
    Summary = $summary
    Details = $details
    Health  = $health
  }
}

function Test-ADReplication {
  param([string]$Domain)

  $summary = @()
  $details = ""
  $health = "OK"

  if ($Script:ADHealthTools['repadmin.exe']) {
    ## Actually run repadmin
    $replresult = Invoke-ExternalCommand -Exe "repadmin.exe" -ArgumentList "/replsummary $Domain"

    if ($replresult.ExitCode -eq 0) {
      $lines = ($replresult.StdOut -split "`r?`n") | Where-Object { $_ -match '\S' } | Select-Object -First 20
      $summary = $lines
      $details = $replresult.StdOut

      if ($replresult.StdOut -match "error|fail") {
        $health = "FAIL"
      }
    } else {
      $summary += "repadmin returned error code $($replresult.ExitCode)"
      $details = $replresult.StdErr
      $health = "WARN"
    }
  } else {
    ## Fallback to AD cmdlet
    try {
      $failures = Get-ADReplicationFailure -Scope Domain -Target $Domain -ErrorAction Stop

      if ($failures) {
        foreach ($failure in $failures) {
          $summary += "FAIL: $($failure.Server) - $($failure.FirstFailureMessage)"
        }
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
  $health = "OK"

  $srvRecords = @(
    "_ldap._tcp.dc._msdcs.$Domain"
    "_kerberos._tcp.$Domain"
    "_ldap._tcp.$Domain"
  )

  ## Actually query DNS
  foreach ($srv in $srvRecords) {
    try {
      $q = Resolve-DnsName -Name $srv -Type SRV -ErrorAction Stop
      $count = ($q | Measure-Object).Count
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
        $sysvolPath = "\\$($dc.HostName)\SYSVOL"
        $netlogonPath = "\\$($dc.HostName)\NETLOGON"

        $sysvolOK = Test-Path $sysvolPath -ErrorAction SilentlyContinue
        $netlogonOK = Test-Path $netlogonPath -ErrorAction SilentlyContinue

        $sysvolStatus = if ($sysvolOK) { "✓ Available" } else { "✗ Unavailable"; $health = "FAIL" }
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
  $health = "OK"

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
      Debug-Log ": Using imported GPO data ($($Script:rawGPOs.Count) GPOs)" -Type "Info"

      ## Filter by domain if specified
      if ($Domain) {
        $gpos = $Script:rawGPOs | Where-Object {
          $_.DN -match [regex]::Escape("DC=$($Domain -replace '\.',',DC=')")
        }
      } else {
        $gpos = $Script:rawGPOs
      }

    } else {
      ## Fall back to production query
      Debug-Log ": Querying GPOs from Active Directory for domain: $Domain" -Type "Info"
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

    $details = "GPO Details:`n`n"

    foreach ($gpo in $gpos | Select-Object -First 50) {
      ## Handle both imported GPO format and production AD GPO objects
      $gpoName = if ($gpo.DisplayName) { $gpo.DisplayName } else { $gpo.Name }
      $gpoPath = if ($gpo.GPCFileSysPath) { $gpo.GPCFileSysPath } else { "N/A" }
      $gpoVersion = if ($gpo.VersionNumber) { $gpo.VersionNumber } else { "N/A" }

      $status = "✓ Present"

      $summary += $fmt -f $gpoName, $status
      $details += "  Name: $gpoName`n"
      $details += "    Path: $gpoPath`n"
      $details += "    Version: $gpoVersion`n"

      if ($gpo.Description) {
        $details += "    Description: $($gpo.Description)`n"
      }
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

  if (-not $Domain) {
    $Domain = $Script:CurrentDomain
  }

  $result = Test-GPOHealth -Domain $Domain

  $output = @()
  $output += ""
  $output += "Group Policy Objects - Domain: $Domain"
  $output += ""

  if ($result.Summary) {
    foreach ($line in $result.Summary) {
      $output += $line
    }
  }

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
    }).GetNewClosure()

    ## --- Show Password toggle ---
    $chkShowPwd.add_Toggled({
      if ($chkShowPwd.Checked) { $txtPwd.Text = $Script:actualPassword
      } else { $txtPwd.Text = ('*' * $Script:actualPassword.Length)
      }
    })

    ## --- Copy to Clipboard ---
    $btnCopy.add_Clicked({
      if (-not $Script:actualPassword) { return }
      if ($IsWindows) { Set-Clipboard -Value $Script:actualPassword }
      elseif ($IsMacOS) { $Script:actualPassword | pbcopy }
      else { $Script:actualPassword | xsel --clipboard --input }
      Show-Modal "Copied" "Password copied to clipboard."
    }).GetNewClosure()

    ## --- Close ---
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

  # ---- Step 1: identify LAPS-enabled devices ----
  $LAPSDevices = $Computers | Where-Object {

    # Legacy LAPS
    (
      $_.PSObject.Properties['ms-Mcs-AdmPwd'] -and
      -not [string]::IsNullOrWhiteSpace($_.'ms-Mcs-AdmPwd')
    )

    -or

    # Windows LAPS (any populated msLAPS-* attribute)
    (
      $_.PSObject.Properties |
        Where-Object {
          $_.Name -like 'msLAPS-*' -and
          -not [string]::IsNullOrWhiteSpace($_.Value)
        }
    )
  }

  if (-not $LAPSDevices) {
    Debug-Log "No LAPS-enabled devices found" -Type "Warn"
    return @()
  }

  ## ---- Step 2: find properties populated on ALL LAPS devices ----
  $AllProps = $LAPSDevices[0].PSObject.Properties.Name

  $CommonProps = $AllProps | Where-Object {
    $prop = $_
    -not ($LAPSDevices | Where-Object {
      -not $_.PSObject.Properties[$prop] -or
      [string]::IsNullOrWhiteSpace($_.$prop)
    })
  }

  ## ---- Step 3: output only those properties ----
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

  Debug-Log ": Show-LAPSSearchModal called" -Type "Info"

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

  ## ==================== Data Loader ====================
  $loadComputers = {
    param($filter)

    Debug-Log ": Loading LAPS computers with filter: '$filter'" -Type "Info"
    try {
      if ($Script:DemoMode) {
        ## Demo mode - check computers for LAPS properties
        Debug-Log ": Demo mode - using sample data" -Type "Info"
        Debug-Log ": Total computers: $($Script:Computers.Count)" -Type "Debug"
        $Script:LAPSComputers = @()
        foreach ($comp in $Script:Computers) {
          ## Skip if filter doesn't match
          if ($filter -and $comp.Name -notlike "*$filter*") { continue }

          ## Check for LAPS properties
          $hasLaps = $false
          $lapsUser = ""
          $lapsPass = ""
          $lapsExpiry = ""

          ## Check Windows LAPS
          if ($comp.'msLAPS-Password') {
            $hasLaps = $true
            $lapsUser = if ($comp.'msLAPS-AccountName') { $comp.'msLAPS-AccountName' } else { 'Administrator' }
            $lapsPass = $comp.'msLAPS-Password'
            $lapsExpiry = $comp.'msLAPS-PasswordExpirationTime'
            Debug-Log ": $($comp.Name) has Windows LAPS (pass length: $($lapsPass.Length))" -Type "Debug"
          }
          ## Check Legacy LAPS
          elseif ($comp.'ms-Mcs-AdmPwd') {
            $hasLaps = $true
            $lapsUser = 'Administrator'
            $lapsPass = $comp.'ms-Mcs-AdmPwd'
            $lapsExpiry = $comp.'ms-Mcs-AdmPwdExpirationTime'
            Debug-Log ": $($comp.Name) has Legacy LAPS (pass length: $($lapsPass.Length))" -Type "Debug"
          }
          if ($hasLaps) {
            $Script:LAPSComputers += [PSCustomObject]@{
              Name = $comp.Name
              DNSHostName = $comp.DNSHostName ?? "$($comp.Name).$($Script:CurrentDomain)"
              LapsUser = $lapsUser
              LapsPass = $lapsPass
              LapsExpiry = $lapsExpiry
            }
          }
        }

        Debug-Log ": Found $($Script:LAPSComputers.Count) computers with LAPS" -Type "Info"

      } else {
        ## Production mode - detect LAPS schema once
        if (-not $Script:LAPSSchemaDetected) {
          try {
            $windowsLapsSchema = Get-ADObject -SearchBase (Get-ADRootDSE).SchemaNamingContext  -LDAPFilter "(lDAPDisplayName=msLAPS-Password)" -ErrorAction SilentlyContinue

            if ($windowsLapsSchema) {
              $Script:LAPSType = "Windows"
              $Script:LAPSProps = @('Name', 'DNSHostName', 'msLAPS-Password', 'msLAPS-PasswordExpirationTime', 'msLAPS-AccountName')
              Debug-Log ": Detected Windows LAPS schema" -Type "Info"
            } else {
              $Script:LAPSType = "Legacy"
              $Script:LAPSProps = @('Name', 'DNSHostName', 'ms-Mcs-AdmPwd', 'ms-Mcs-AdmPwdExpirationTime')
              Debug-Log ": Detected Legacy LAPS schema" -Type "Info"
            }

            $Script:LAPSSchemaDetected = $true
          } catch {
            Debug-Log ": Schema detection failed, defaulting to Legacy LAPS" -Type "Warn"
            $Script:LAPSType = "Legacy"
            $Script:LAPSProps = @('Name', 'DNSHostName', 'ms-Mcs-AdmPwd', 'ms-Mcs-AdmPwdExpirationTime')
            $Script:LAPSSchemaDetected = $true
          }
        }

        ## Query AD with correct properties
        $filterString = if ([string]::IsNullOrWhiteSpace($filter)) { "*" } else { "*$filter*" }

        Debug-Log ": Querying AD with filter: Name -like '$filterString'" -Type "Info"

        try {
          $rawComputers = Get-ADComputer -Filter "Name -like '$filterString'" -Properties $Script:LAPSProps -ErrorAction Stop
        } catch {
          Debug-Log ": AD query failed: $($_.Exception.Message)" -Type "Error"
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

        Debug-Log ": Found $($Script:LAPSComputers.Count) computers with LAPS" -Type "Info"
      }

      ## Update ListView
      if ($Script:LAPSComputers.Count -eq 0) {
        $lstComputers.SetSource(@("(No computers with LAPS found)"))
      } else {
        $computerNames = $Script:LAPSComputers | ForEach-Object { $_.Name }
        $lstComputers.SetSource($computerNames)
      }

    } catch {
      Debug-Log ": Error loading LAPS computers: $($_.Exception.Message)" -Type "Error"
      Show-Modal "Error" "Failed to load LAPS computers:`n`n$($_.Exception.Message)"
      $lstComputers.SetSource(@("(Error loading computers)"))
    }
  }

  ## Initial load
  & $loadComputers ""

  ## ==================== Buttons ====================
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

  ## ==================== View Handler ====================
  $btnView.add_Clicked({
    if ($lstComputers.SelectedItem -lt 0 -or $Script:LAPSComputers.Count -eq 0) {
      Show-Modal "No Selection" "Please select a computer from the list."
      return
    }

    try {
      $computer = $Script:LAPSComputers[$lstComputers.SelectedItem]
      Debug-Log ": Selected computer: $($computer.Name)" -Type "Info"
      Debug-Log ": LapsUser: '$($computer.LapsUser)'" -Type "Debug"
      Debug-Log ": LapsPass length: $($computer.LapsPass.Length)" -Type "Debug"
      Debug-Log ": LapsExpiry: '$($computer.LapsExpiry)'" -Type "Debug"

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
          Debug-Log ": Failed to parse expiry: $($_.Exception.Message)" -Type "Warn"
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
      $lblPasswordLabel = [Terminal.Gui.Label]::new("LAPS Password:")
      $lblPasswordLabel.X = 2; $lblPasswordLabel.Y = 5
      $detailDlg.Add($lblPasswordLabel)

      ## Ensure password is a string and create masked version
      $passStr = $computer.LapsPass.ToString()
      $masked = '*' * $passStr.Length
      $Script:LAPSPasswordRevealed = $false

      $lblPasswordValue = [Terminal.Gui.Label]::new($masked)
      $lblPasswordValue.X = 18; $lblPasswordValue.Y = 5; $lblPasswordValue.Width = 16
      $detailDlg.Add($lblPasswordValue)

      ## Show/Hide button (text-based)
      $btnReveal = [Terminal.Gui.Button]::new("👁")
      $btnReveal.X = 32; $btnReveal.Y = 5; $btnReveal.Width = 6
      $btnReveal.add_Clicked({
        if ($Script:LAPSPasswordRevealed) {
          ## Hide password
          $lblPasswordValue.Text = [NStack.ustring]::Make($masked)
          $btnReveal = [Terminal.Gui.Button]::new("👁")
          $Script:LAPSPasswordRevealed = $false
          Debug-Log ": Password hidden" -Type "Debug"
        } else {
          ## Show password
          $lblPasswordValue.Text = [NStack.ustring]::Make($passStr)
          $btnReveal.Text = [NStack.ustring]::Make("🕶️")
          $Script:LAPSPasswordRevealed = $true
          Debug-Log ": Password revealed" -Type "Debug"
        }
      })
      $detailDlg.Add($btnReveal)

      $btnCopyPassword = [Terminal.Gui.Button]::new("📋")
      $btnCopyPassword.X = 38; $btnCopyPassword.Y = 5; $btnCopyPassword.Width = 6
      $btnCopyPassword.add_Clicked({
        Set-Clipboard -Value $passStr
        Show-Modal "Copied" "Password copied to clipboard."
      })
      $detailDlg.Add($btnCopyPassword)

      ## Expiry warning
      $lblExpiry = [Terminal.Gui.Label]::new($warn)
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
            Debug-Log ": LAPS password expiration set to 0 for $($computer.Name)" -Type "Success"

            try {
              Invoke-GPUpdate -Computer $computer.Name -Force -ErrorAction Stop
              Show-Modal "Rotation Complete" "LAPS password rotation initiated!`n`nExpiration set to 0 and GP updated."
            } catch {
              Show-Modal "Rotation Requested" "LAPS will rotate on next GP refresh.`n`nExpiration set to 0."
            }
          } catch {
            Show-Modal "Error" "Failed to force rotation:`n`n$($_.Exception.Message)"
            Debug-Log ": Failed to force LAPS rotation: $($_.Exception.Message)" -Type "Error"
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
      Debug-Log ": Error showing LAPS details: $($_.Exception.Message)" -Type "Error"
    }
  }).GetNewClosure()

  [Terminal.Gui.Application]::Run($dialog)
}

## ====================================================={ Main Dialog Function }=====================================================

# DSA-TUI Object Management Module v1.0
# Create, Delete, and Move AD Objects

## -------------------------{ Create New Object Wizard }-------------------------
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
    $isUser = $rdoType.SelectedItem -eq 0
    $lblSam.Visible = $isUser
    $txtSam.Visible = $isUser
    $lblEmail.Visible = $isUser
    $txtEmail.Visible = $isUser
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

          ## Add to raw users array
          $Script:rawUsers += $newUser
          ## Reconvert to update $Script:Users with AD-like objects
          $converted = Convert-DataToADObjects -Users $Script:rawUsers -DCs $Script:rawDCs -Groups $Script:rawDemoGroups -Domain $Script:CurrentDomain
          $Script:Users = $converted.Users

          Debug-Log (": Created user $name in demo mode") -Type "Info"
        }

        "Group" {
          ## Add to RAW group data
          $newGroup = @{
            Name = $name
            Description = $displayName
            Type = 'Security'
            Scope = 'Global'
            ManagedBy = ''
            Email = ''
          }

          ## Add to raw groups array
          $Script:rawDemoGroups += $newGroup


          #Reconvert to update $Script:Groups with AD-like objects
          $converted = Convert-DataToADObjects -Users $Script:rawUsers -DCs $Script:rawDCs -Groups $Script:rawDemoGroups -Domain $Script:CurrentDomain
          $Script:Groups = $converted.Groups
          Debug-Log (": Created group $name in demo mode") -Type "Info"
        }

        "OrganizationalUnit" {
          ## For OUs, we need to track them in a structure. OUs are built from the OU arrays in users, so we could either:
          ## 1. Add a $Script:rawOUs array (cleaner)
          ## 2. Just ensure the OU path exists when we rebuild the tree

          ## For now, let's just ensure it's tracked
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
          ## Similar to users, add to a computers array
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
      Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
      [Terminal.Gui.Application]::RequestStop()

      } else {
        ## Production mode - create in AD
        switch ($objType) {
          "User" {
            $sam = $txtSam.Text.ToString().Trim()
            $email = $txtEmail.Text.ToString().Trim()
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

# -------------------------[] Change Domain Dialog ]-------------------------
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
    Debug-Log (": OK pressed, Domain = $domainString") -Type "Info"

    ## Close the dialog first
    [Terminal.Gui.Application]::RequestStop()

    ## Schedule the domain change after dialog closes
    [Terminal.Gui.Application]::MainLoop.Invoke({
      ## Save previous domain for fallback
      $Script:PreviousDomain = $Script:CurrentDomain
      Debug-Log (": Saved previous domain: $Script:PreviousDomain") -Type "Info"

      try {
        Set-StatusBar "Changing domain to $domainString..." -spinner

        ## Update the current domain
        $Script:CurrentDomain = $domainString
        $Script:Domain = $Script:CurrentDomain  ## Compatibility
        Debug-Log (": CurrentDomain set to: $Script:CurrentDomain") -Type "Info"

        ## Reset all Script variables
        Debug-Log (": Resetting Script variables...") -Type "Info"
        $Script:CurrentDC       = $null
        $Script:Users           = @()
        $Script:Groups          = @()
        $Script:DCs             = @()
        $Script:ADObjects       = @()
        $Script:SelectedObjects = @()
        $Script:SelectionMode   = $false

        ## Load domain data
        Set-StatusBar "Loading domain data for $($Script:CurrentDomain)..." -Spinner
        Debug-Log "Loading domain data for $($Script:CurrentDomain)..." -Type "Info"

        Initialise-DataSource -Domain $Script:CurrentDomain

        Debug-Log "POST-LOAD: Users=$(${Script:Users}.Count), DCs=$(${Script:DCs}.Count), Computers=$(${Script:Computers}.Count), Groups=$(${Script:Groups}.Count), Objects=$(${Script:ADObjects}.Count)" -Type"Info"
        Debug-Log "Forest/Domain initialization complete: CurrentDomain=$($Script:CurrentDomain)" -Type "Info"

        ## Build tree
        Set-StatusBar "Building tree..." -Spinner
        $rootNode = Build-Tree -domain $Script:CurrentDomain

        if ($null -ne $rootNode) {
          $Script:tree.ClearObjects()
          $Script:tree.AddObject($rootNode)
          Debug-Log ": Root node added to TreeView" -Type "Success"
          Debug-Log ": TreeView created and added to window successfully" -Type "Success"

          ## Update filter status if it exists
          if ($Script:FilterStatusLabel) {
            Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
          }

          Set-StatusBar "Ready" -Final
          Debug-Log (": Domain change successful to $domainString") -Type "Success"
        } else {
          throw "Build-Tree returned null root node"
        }

      } catch {
        Debug-Log (": Domain change error: $($_.Exception.Message)") -Type "Error"
        Debug-Log (": Falling back to previous domain: $Script:PreviousDomain") -Type "Warn"

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
          Set-StatusBar "Restoring previous domain $Script:PreviousDomain..." -Spinner
          Initialise-DataSource -Domain $Script:CurrentDomain
          Show-InfoPanel -UpdateOnly
          $rootNode = Build-Tree -domain $Script:CurrentDomain

          if ($null -ne $rootNode) {
            $Script:tree.ClearObjects()
            $Script:tree.AddObject($rootNode)
            Debug-Log ": Restored previous domain successfully" -Type "Success"
          }
        } catch {
          Debug-Log (": Failed to restore previous domain: $($_.Exception.Message)") -Type "Error"
        }

        Set-StatusBar "Domain change failed - restored to $Script:PreviousDomain" -final
        Show-Modal "Error" "Failed to load domain '$domainString'`n`nError: $($_.Exception.Message)`n`nRestored to previous domain: $Script:PreviousDomain"
      }
    })
  }).GetNewClosure()
  $dlg.Add($okBtn)

  $cancelBtn = [Terminal.Gui.Button]::new("Cancel")
  $cancelBtn.X = 25
  $cancelBtn.Y = 4
  $cancelBtn.add_Clicked({
    Debug-Log (": Cancel pressed") -Type "Info"
    [Terminal.Gui.Application]::RequestStop()
  }).GetNewClosure()
  $dlg.Add($cancelBtn)

  [Terminal.Gui.Application]::Run($dlg)
}

## -------------------------{ Change DC Dialog }-------------------------
function Show-ChangeDCDialog {
  <#
  .SYNOPSIS
  Show enhanced Domain Controller selection dialog

  .DESCRIPTION
  Displays available DCs with details (site, health, FSMO roles)
  Shows current DC with indicator
  dsa.msc-style experience
  #>

  Debug-Log ": Opening Change DC dialog" -Type "Info"

  ## Get available DCs
  $availableDCs = if ($Script:DemoMode) {
    ## Demo mode - use DCs (not rawDCs)
    if ($Script:DCs) {
      $Script:DCs | Where-Object { $_.Domain -eq $Script:CurrentDomain }
    } else {
      @()
    }
  } else {
    ## Production mode - query AD
    try {
      Get-ADDomainController -Filter * | Select-Object Name, Site, IPv4Address, IsGlobalCatalog, OperatingSystem
    } catch {
      Debug-Log ": Failed to get DCs: $($_.Exception.Message)" -Type "Error"
      @()
    }
  }

  if ($availableDCs.Count -eq 0) {
    Show-Modal "No DCs Found" "No Domain Controllers found for domain $($Script:CurrentDomain)"
    return
  }

  Debug-Log ": Found $($availableDCs.Count) DCs in $($Script:CurrentDomain)" -Type "Info"

  ## Create dialog
  $dlg = [Terminal.Gui.Dialog]::new("Change Domain Controller", 90, 28)
  $y = 1

  ## Header
  $lblHeader = [Terminal.Gui.Label]::new("Select Domain Controller for $($Script:CurrentDomain):")
  $lblHeader.X = 2
  $lblHeader.Y = $y
  $dlg.Add($lblHeader)
  $y += 2

  ## Current DC indicator
  $currentDCName = if ($Script:CurrentDC -is [string]) {
    $Script:CurrentDC
  } elseif ($Script:CurrentDC) {
    $Script:CurrentDC.Name
  } else {
    "None"
  }

  $lblCurrent = [Terminal.Gui.Label]::new("Current: $currentDCName")
  $lblCurrent.X = 2
  $lblCurrent.Y = $y
  $dlg.Add($lblCurrent)
  $y += 2

  ## DC List
  $lstDCs = [Terminal.Gui.ListView]::new()
  $lstDCs.X = 2
  $lstDCs.Y = $y
  $lstDCs.Width = [Terminal.Gui.Dim]::Fill(2)
  $lstDCs.Height = 12

  ## Build display items
  $displayItems = @()
  foreach ($dc in ($availableDCs | Sort-Object Name)) {
    $currentMarker = if ($dc.Name -eq $currentDCName) { "► " } else { "  " }
    $gcIcon = if ($dc.IsGlobalCatalog) { "🌐 " } else { "   " }

    ## Health status
    $healthIcon = if ($Script:DemoMode -and $dc.ReplicationHealth) {
      switch ($dc.ReplicationHealth) {
        'Healthy' { "✓" }
        { $_ -match 'Warning' } { "▲" }
        { $_ -match 'Critical|Error|Failed' } { "✗" }
        default { "?" }
      }
    } else {
      "✓"  # Assume healthy in production if we can query it
    }

    ## FSMO roles count
    $fsmoCount = if ($dc.FSMORoles) { $dc.FSMORoles.Count } else { 0 }
    $fsmoText = if ($fsmoCount -gt 0) { " [FSMO: $fsmoCount]" } else { "" }

    ## Site info
    $site = $dc.Site ?? "N/A"

    ## IP Address
    $ip = $dc.IPv4Address ?? "N/A"

    $displayItems += "$currentMarker$gcIcon$($dc.Name.PadRight(18)) | Site: $($site.PadRight(8)) | IP: $($ip.PadRight(15)) $fsmoText $healthIcon"
  }

  $lstDCs.SetSource($displayItems)

  ## Pre-select current DC
  $currentIndex = 0
  for ($i = 0; $i -lt $availableDCs.Count; $i++) {
    if ($availableDCs[$i].Name -eq $currentDCName) {
      $currentIndex = $i
      break
    }
  }
  $lstDCs.SelectedItem = $currentIndex
  $dlg.Add($lstDCs)
  $y += 14

  ## Details section
  $lblDetails = [Terminal.Gui.Label]::new("Details:")
  $lblDetails.X = 2
  $lblDetails.Y = $y
  $dlg.Add($lblDetails)
  $y += 1

  $lblDetailText = [Terminal.Gui.Label]::new("")
  $lblDetailText.X = 2
  $lblDetailText.Y = $y
  $lblDetailText.Width = 84
  $lblDetailText.Height = 3
  $dlg.Add($lblDetailText)

  ## Update details when selection changes
  $lstDCs.add_SelectedItemChanged({
    if ($lstDCs.SelectedItem -ge 0 -and $lstDCs.SelectedItem -lt $availableDCs.Count) {
      $selectedDC = $availableDCs[$lstDCs.SelectedItem]
      $detailText = "Name: $($selectedDC.Name)`n"
      $detailText += "Site: $($selectedDC.Site ?? 'N/A')  |  "
      $detailText += "IP: $($selectedDC.IPv4Address ?? 'N/A')  |  "
      $detailText += "GC: $(if ($selectedDC.IsGlobalCatalog) { 'Yes' } else { 'No' })`n"

      if ($Script:DemoMode -and $selectedDC.FSMORoles -and $selectedDC.FSMORoles.Count -gt 0) {
        $detailText += "FSMO Roles: $($selectedDC.FSMORoles -join ', ')"
      } elseif ($Script:DemoMode -and $selectedDC.Location) {
        $detailText += "Location: $($selectedDC.Location)"
      } elseif ($selectedDC.OperatingSystem) {
        $detailText += "OS: $($selectedDC.OperatingSystem)"
      }
      $lblDetailText.Text = [NStack.ustring]::Make($detailText)
    }
  }.GetNewClosure())

  ## Trigger initial details update
  if ($lstDCs.SelectedItem -ge 0) {
    $selectedDC = $availableDCs[$lstDCs.SelectedItem]
    $detailText = "Name: $($selectedDC.Name)`n"
    $detailText += "Site: $($selectedDC.Site ?? 'N/A')  |  "
    $detailText += "IP: $($selectedDC.IPv4Address ?? 'N/A')  |  "
    $detailText += "GC: $(if ($selectedDC.IsGlobalCatalog) { 'Yes' } else { 'No' })`n"

    if ($Script:DemoMode -and $selectedDC.FSMORoles -and $selectedDC.FSMORoles.Count -gt 0) {
      $detailText += "FSMO Roles: $($selectedDC.FSMORoles -join ', ')"
    } elseif ($Script:DemoMode -and $selectedDC.Location) {
      $detailText += "Location: $($selectedDC.Location)"
    } elseif ($selectedDC.OperatingSystem) {
      $detailText += "OS: $($selectedDC.OperatingSystem)"
    }

    $lblDetailText.Text = [NStack.ustring]::Make($detailText)
  }

  ## Connect button
  $btnConnect = [Terminal.Gui.Button]::new("Connect")
  $btnConnect.add_Clicked({
    if ($lstDCs.SelectedItem -ge 0 -and $lstDCs.SelectedItem -lt $availableDCs.Count) {
      $selectedDC = $availableDCs[$lstDCs.SelectedItem]
      $oldDC = $currentDCName

      ## Update CurrentDC - store the DC object
      $Script:CurrentDC = $selectedDC

      Debug-Log ": Changed DC from $oldDC to $($selectedDC.Name)" -Type "Success"

      ## Update InfoPanel
      Show-InfoPanel -UpdateOnly

      ## Refresh data from new DC (only in production mode)
      if (-not $Script:DemoMode) {
        Refresh-Data -Domain $Script:CurrentDomain -RebuildTree
      }

      Show-Modal "DC Changed" "Successfully connected to $($selectedDC.Name)"
      [Terminal.Gui.Application]::RequestStop()
    } else {
      Show-Modal "No Selection" "Please select a Domain Controller"
    }
  }.GetNewClosure())
  $dlg.AddButton($btnConnect)

  ## Cancel button
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
  $btnCancel.add_Clicked({
    Debug-Log ": Change DC cancelled" -Type "Info"
    [Terminal.Gui.Application]::RequestStop()
  }.GetNewClosure())
  $dlg.AddButton($btnCancel)

  ## Run dialog
  [Terminal.Gui.Application]::Run($dlg)
}

## -------------------------{ Tree Expand/Collapse }-------------------------
## Check tree exists first or it blows things up If we're initalising, it is created later
if ($null -ne $Script:tree) {
  $Script:tree.Add_KeyPress({
  param($senders, $keyArgs)
    if ($keyArgs.KeyEvent.Key -ne [Terminal.Gui.Key]::Enter) { return }
    $node = $Script:tree.SelectedNode

    if (-not $node) {
      Debug-Log ("Enter pressed but no selected node — ignoring safely.") -Type "Info"
      return
    }

    ## Toggle expand/collapse properly
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

  ## -------------------------{ AD Search Dialog }-------------------------
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
    ## New parameters for clipboard buttons
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
      if (-not $searchName) { $TxtOutput.Text="Please enter a name."; return }
        if ($Script:DemoMode) {
          switch ($objType) {
            "User"  { $objs = $Script:Users | Where-Object { $_.Name -like "*$searchName*" } | Select-Object @{Name='Name';Expression={$_.Name}}, @{Name='Type';Expression={"user"}} }
            "Group" { $matchedGroups = @(); foreach ($u in $Script:Users) { foreach ($g in $u.Groups) { if ($g -like "*$searchName*") { $matchedGroups += $g } } }
                      $objs = ($matchedGroups | Sort-Object -Unique) | ForEach-Object { [PSCustomObject]@{ Name=$_; Type="group" } }}
            "OU"    { $ouNames = ($Script:Users | Select-Object -ExpandProperty OU -Unique)
                      $objs = ($ouNames | Where-Object { $_ -like "*$searchName*" }) | ForEach-Object { [PSCustomObject]@{ Name=$_; Type="organizationalUnit" } } }
            "Computer" { $objs = @() }
            "Contact"  { $objs = @() }
          }
        $Script:lastSearchType = "Basic ($objType)"
        } else {
          $loading = Show-LoadingDialog -Message "Searching AD for $objType '$searchName'..."
          try {
            $filterStr = "Name -like '*$searchName*'"
            if ($ChkDisabledOnly.Checked -and $objType -eq "User") { $filterStr = "Name -like '*$searchName*' -and Enabled -eq $false" }
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

      ## Store results for export
      $Script:lastSearchResults = $objs

      if (-not $objs -or $objs.Count -eq 0) { $TxtOutput.Text = "No results found"; return }
        ## Display results
        $resultText = "Found $($objs.Count) object(s):`n`n"
        $resultText += ($objs | ForEach-Object { "$($_.Name) [$($_.Type)]" }) -join "`n"
        $TxtOutput.Text = $resultText

    } catch {
      $TxtOutput.Text = "Error: $($_.Exception.Message)"
  }
}

# ====================================================={ Clipboard Helper Functions }=====================================================

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

    ## Format results for clipboard
    $clipboardText = "# AD Search Results - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
    $clipboardText += "# Search Type: $($Script:lastSearchType)`n"
    $clipboardText += "# Total Results: $($Script:lastSearchResults.Count)`n"
    $clipboardText += "`n"

    ## Add results in multiple formats for flexibility

    ## Format 1: Simple list
    $clipboardText += "=== SIMPLE LIST ===`n"
    foreach ($obj in $Script:lastSearchResults) {
      $clipboardText += "$($obj.Name) [$($obj.Type)]`n"
    }

    $clipboardText += "`n=== CSV FORMAT ===`n"
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
    $clipboardText += "`n=== POWERSHELL NAMES ===`n"
    $clipboardText += "@(`n"
    $names = $Script:lastSearchResults | ForEach-Object { '   "' + $_.Name + '"' }
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
# ====================================================={ Example: How to Add Clipboard Buttons to Your UI }=====================================================

<# In your Search/Lookup tab creation, add these buttons:

# LDAP Filter section
$lbl = [Terminal.Gui.Label]::new("LDAP Filter:"); $lbl.X=2; $lbl.Y=$y; $advView.Add($lbl); $y+=1
$txtLdapFilter = [Terminal.Gui.TextView]::new(); $txtLdapFilter.X=2; $txtLdapFilter.Y=$y; $txtLdapFilter.Width=[Terminal.Gui.Dim]::Fill(2); $txtLdapFilter.Height=4
$advView.Add($txtLdapFilter); $y+=5

# Clipboard buttons for LDAP query
$btnCopyQuery = [Terminal.Gui.Button]::new("Copy Query"); $btnCopyQuery.X=2; $btnCopyQuery.Y=$y
$btnCopyQuery.add_Clicked({ Copy-LDAPQueryToClipboard -LdapFilter $txtLdapFilter }).GetNewClosure()
$advView.Add($btnCopyQuery)

$btnPasteQuery = [Terminal.Gui.Button]::new("Paste Query"); $btnPasteQuery.X=18; $btnPasteQuery.Y=$y
$btnPasteQuery.add_Clicked({ Paste-LDAPQueryFromClipboard -LdapFilter $txtLdapFilter }).GetNewClosure()
$advView.Add($btnPasteQuery)
$y+=2

# Search button
$btnSearch = [Terminal.Gui.Button]::new("Search"); $btnSearch.X=2; $btnSearch.Y=$y
$btnSearch.add_Clicked({
    Invoke-ADSearch -UserField $txtSearchName -DomainField $txtSearchDomain -ObjType $cmbSearchType -TabView $searchTabView -TxtOutput $txtSearchOutput -ChkDisabledOnly $chkDisabledOnly -LdapFilter $txtLdapFilter -AdvTab $advTab
}).GetNewClosure()
$advView.Add($btnSearch); $y+=3

# Results section
$lbl = [Terminal.Gui.Label]::new("Results:"); $lbl.X=2; $lbl.Y=$y; $advView.Add($lbl); $y+=1
$txtSearchOutput = [Terminal.Gui.TextView]::new(); $txtSearchOutput.X=2; $txtSearchOutput.Y=$y
$txtSearchOutput.Width=[Terminal.Gui.Dim]::Fill(2); $txtSearchOutput.Height=[Terminal.Gui.Dim]::Fill(4); $txtSearchOutput.ReadOnly=$true
$advView.Add($txtSearchOutput); $y+=[Terminal.Gui.Dim]::Fill(3)

# Copy results button
$btnCopyResults = [Terminal.Gui.Button]::new("Copy Results to Clipboard"); $btnCopyResults.X=2; $btnCopyResults.Y=[Terminal.Gui.Pos]::Bottom($txtSearchOutput)+1
$btnCopyResults.add_Clicked({ Copy-SearchResultsToClipboard -TxtOutput $txtSearchOutput }).GetNewClosure()
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
    ## ---------------------- Buttons ----------------------
    $btnClose = [Terminal.Gui.Button]::new("Close")
    $btnClose.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()

    ## ---------------------- Dialog ----------------------
    $dlg = [Terminal.Gui.Dialog]::new("Active Directory Search", 120, 40, $btnClose)

    ## ---------------------- TabView for Basic/Advanced ----------------------
    $searchTabView = [Terminal.Gui.TabView]::new()
    $searchTabView.X = 0
    $searchTabView.Y = 0
    $searchTabView.Width = [Terminal.Gui.Dim]::Fill()
    $searchTabView.Height = [Terminal.Gui.Dim]::Fill(1)

    ## ==================== Basic Search Tab ====================
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
    $cmbObjectType.SetSource(@("User", "Group", "Computer", "OU", "Contact"))
    $cmbObjectType.SelectedItem = 0
    $basicView.Add($cmbObjectType); $y+=2

    ## Disabled only checkbox (for users)
    $chkDisabledOnly = [Terminal.Gui.CheckBox]::new("Disabled accounts only"); $chkDisabledOnly.X=20; $chkDisabledOnly.Y=$y
    $basicView.Add($chkDisabledOnly); $y+=2

    ## Search button
    $btnSearch = [Terminal.Gui.Button]::new("Search"); $btnSearch.X=20; $btnSearch.Y=$y
    $basicView.Add($btnSearch); $y+=3

    ## Results
    $lbl = [Terminal.Gui.Label]::new("Results:"); $lbl.X=2; $lbl.Y=$y; $basicView.Add($lbl); $y+=1
    $txtResults = [Terminal.Gui.TextView]::new()
    $txtResults.X = 2
    $txtResults.Y = $y
    $txtResults.Width = [Terminal.Gui.Dim]::Fill(2)
    $txtResults.Height = [Terminal.Gui.Dim]::Fill(3)
    $txtResults.ReadOnly = $true
    $basicView.Add($txtResults)

    ## Copy Results button
    $btnCopyResults = [Terminal.Gui.Button]::new("Copy Results to Clipboard")
    $btnCopyResults.X = 2
    $btnCopyResults.Y = [Terminal.Gui.Pos]::Bottom($txtResults) + 1
    $btnCopyResults.add_Clicked({ Copy-SearchResultsToClipboard -TxtOutput $txtResults }).GetNewClosure()
    $basicView.Add($btnCopyResults)
    $basicTab.View = $basicView
    $searchTabView.AddTab($basicTab, $false)

    ## ==================== Advanced LDAP Tab ====================
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
    $btnCopyQuery.add_Clicked({ Copy-LDAPQueryToClipboard -LdapFilter $txtLdapFilter }).GetNewClosure()
    $advView.Add($btnCopyQuery)

    $btnPasteQuery = [Terminal.Gui.Button]::new("Paste Query"); $btnPasteQuery.X=18; $btnPasteQuery.Y=$y
    $btnPasteQuery.add_Clicked({ Paste-LDAPQueryFromClipboard -LdapFilter $txtLdapFilter }).GetNewClosure()
    $advView.Add($btnPasteQuery); $y+=2

    ## Search button
    $btnSearchAdv = [Terminal.Gui.Button]::new("Execute LDAP Query"); $btnSearchAdv.X=2; $btnSearchAdv.Y=$y
    $advView.Add($btnSearchAdv); $y+=3

    ## Results
    $lbl = [Terminal.Gui.Label]::new("Results:"); $lbl.X=2; $lbl.Y=$y; $advView.Add($lbl); $y+=1
    $txtResultsAdv = [Terminal.Gui.TextView]::new()
    $txtResultsAdv.X = 2
    $txtResultsAdv.Y = $y
    $txtResultsAdv.Width = [Terminal.Gui.Dim]::Fill(2)
    $txtResultsAdv.Height = [Terminal.Gui.Dim]::Fill(3)
    $txtResultsAdv.ReadOnly = $true
    $advView.Add($txtResultsAdv)

    ## Copy Results button
    $btnCopyResultsAdv = [Terminal.Gui.Button]::new("Copy Results to Clipboard")
    $btnCopyResultsAdv.X = 2
    $btnCopyResultsAdv.Y = [Terminal.Gui.Pos]::Bottom($txtResultsAdv) + 1
    $btnCopyResultsAdv.add_Clicked({ Copy-SearchResultsToClipboard -TxtOutput $txtResultsAdv }).GetNewClosure()
    $advView.Add($btnCopyResultsAdv)
    $advTab.View = $advView
    $searchTabView.AddTab($advTab, $false)

    ## ==================== Wire up Search Buttons ====================

    ## Basic Search
    $btnSearch.add_Clicked({
      Invoke-ADSearch -UserField $txtSearchName -DomainField $txtDomain -ObjType $cmbObjectType -TabView $searchTabView -TxtOutput $txtResults -ChkDisabledOnly $chkDisabledOnly  -LdapFilter $txtLdapFilter -AdvTab $advTab
    }).GetNewClosure()

    ## Advanced Search
    $btnSearchAdv.add_Clicked({
      Invoke-ADSearch -UserField $txtSearchName -DomainField $txtDomainAdv -ObjType $cmbObjectType -TabView $searchTabView -TxtOutput $txtResultsAdv -ChkDisabledOnly $chkDisabledOnly -LdapFilter $txtLdapFilter -AdvTab $advTab
    }).GetNewClosure()

    ## Add TabView to dialog
    $dlg.Add($searchTabView)
    Debug-Log ": AD Search Dialog created, running" -Type "Success"

    ## Run the dialog
    [Terminal.Gui.Application]::Run($dlg)

    Debug-Log ": AD Search Dialog closed" -Type "Info"

  } catch {
    Debug-Log ": Exception in Show-ADSearchDialog: $($_.Exception.Message)" -Type "Error"
    Show-Modal "Error" "Failed to open search dialog:`n$($_.Exception.Message)"
  }
}

function Show-DCPropertiesDialog {
  param($dc)

  ## Accept either DC object or DC name
  if ($dc -is [string]) {
    $dcName = $dc
    Debug-Log ": Looking for DC: $dcName" -Type "Info"

    if (-not $Script:DCs) {
      Debug-Log ": Script:DCs is null or not Initialised" -Type "Error"
      Show-Modal "Error" "Domain Controllers list is not loaded"
      return
    }

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

  ## ==================== General Tab ====================
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Domain Controller Information"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Name:" -FieldName 'lblName' -State $state -Value ($dc.Name ?? "Unknown") -FieldX 25 -IsTextField $false

      $hostname = if ($dc.HostName) { $dc.HostName } elseif ($dc.DNSHostName) { $dc.DNSHostName } else { $dc.Name }
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Hostname:" -FieldName 'lblHostname' -State $state -Value $hostname -FieldX 25 -IsTextField $false

      if ($dc.Site) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Site:" -FieldName 'lblSite' -State $state -Value $dc.Site -FieldX 25 -IsTextField $false
      }

      if ($dc.Domain) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Domain:" -FieldName 'lblDomain' -State $state -Value $dc.Domain -FieldX 25 -IsTextField $false
      }

      if ($dc.Forest) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Forest:" -FieldName 'lblForest' -State $state -Value $dc.Forest -FieldX 25 -IsTextField $false
      }

      if ($dc.Location) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Location:" -FieldName 'lblLocation' -State $state -Value $dc.Location -FieldX 25 -IsTextField $false
      }

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Network"

      if ($dc.IPv4Address -or $dc.IPAddress) {
        $ip = if ($dc.IPv4Address) { $dc.IPv4Address } else { $dc.IPAddress }
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "IPv4 Address:" -FieldName 'lblIPv4' -State $state -Value $ip -FieldX 25 -IsTextField $false
      }

      if ($dc.IPv6Address) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "IPv6 Address:" -FieldName 'lblIPv6' -State $state -Value $dc.IPv6Address -FieldX 25 -IsTextField $false
      }

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Operating System"

      if ($dc.OperatingSystem -or $dc.OS) {
        $os = if ($dc.OperatingSystem) { $dc.OperatingSystem } else { $dc.OS }
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "OS:" -FieldName 'lblOS' -State $state -Value $os -FieldX 25 -IsTextField $false
      }

      if ($dc.OperatingSystemVersion) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "OS Version:" -FieldName 'lblOSVer' -State $state -Value $dc.OperatingSystemVersion -FieldX 25 -IsTextField $false
      }

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Capabilities"

      $isGC = if ($null -ne $dc.IsGlobalCatalog) { $dc.IsGlobalCatalog.ToString() } else { "Unknown" }
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Global Catalog:" -FieldName 'lblGC' -State $state -Value $isGC -FieldX 25 -IsTextField $false

      if ($null -ne $dc.IsReadOnly) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Read-Only:" -FieldName 'lblRO' -State $state -Value $dc.IsReadOnly.ToString() -FieldX 25 -IsTextField $false
      }

      if ($null -ne $dc.Enabled) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Enabled:" -FieldName 'lblEnabled' -State $state -Value $dc.Enabled.ToString() -FieldX 25 -IsTextField $false
      }
    }
  }

  ## ==================== Roles Tab ====================
  $rolesTab = @{
    Name = "Roles"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "FSMO Roles"

      $fsmoRoles = @()
      if ($dc.PSObject.Properties['FSMORoles'] -and $dc.FSMORoles -and $dc.FSMORoles.Count -gt 0) {
        $fsmoRoles = $dc.FSMORoles
      } elseif ($dc.PSObject.Properties['OperationMasterRoles'] -and $dc.OperationMasterRoles -and $dc.OperationMasterRoles.Count -gt 0) {
        $fsmoRoles = $dc.OperationMasterRoles
      }

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

  ## ==================== Replication Tab ====================
  $replicationTab = @{
    Name = "Replication"
    Builder = {
      param($view, $dc, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Replication Status"

      if ($dc.ReplicationHealth) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Health:" -FieldName 'lblHealth' -State $state -Value $dc.ReplicationHealth -FieldX 25 -IsTextField $false
      }

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

  ## ==================== Services Tab ====================
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

  ## ==================== Disk Space Tab ====================
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

  ## ==================== RID Pool Tab ====================
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

        if ($ridSet.RIDAvailablePool) {
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Available Pool:" -FieldName 'lblRIDAvail' -State $state -Value $ridSet.RIDAvailablePool.ToString() -LabelX 6 -FieldX 30 -IsTextField $false
        }

        if ($ridSet.RIDAllocationPool) {
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Allocation Pool:" -FieldName 'lblRIDAlloc' -State $state -Value $ridSet.RIDAllocationPool.ToString() -LabelX 6 -FieldX 30 -IsTextField $false
        }

        if ($ridSet.RIDPreviousAllocationPool) {
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Previous Pool:" -FieldName 'lblRIDPrev' -State $state -Value $ridSet.RIDPreviousAllocationPool.ToString() -LabelX 6 -FieldX 30 -IsTextField $false
        }

        if ($ridSet.RIDUsedPool) {
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Used Pool:" -FieldName 'lblRIDUsed' -State $state -Value $ridSet.RIDUsedPool.ToString() -LabelX 6 -FieldX 30 -IsTextField $false
        }

        if ($ridSet.RIDNextRID) {
          Add-LabelAndField -View $view -Y ([ref]$y) -Label "Next RID:" -FieldName 'lblRIDNext' -State $state -Value $ridSet.RIDNextRID.ToString() -LabelX 6 -FieldX 30 -IsTextField $false
        }

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

  ## ==================== DFSR Tab ====================
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

  ## ==================== Create Dialog ====================
  ## Note: Search tab auto-added with DC-specific checkboxes
  $tabs = @($generalTab, $rolesTab, $replicationTab, $servicesTab, $diskTab, $ridPoolTab, $dfsrTab)

  Debug-Log ": All DC tabs added, running dialog" -Type "Success"

New-PropertiesDialog -Title "Domain Controller Properties - $($dc.Name)" -Width 110 -Height 35 -Tabs $tabs -Data $dc -IncludeSearchTab $true -SearchTabConfig @{ ObjectType='DomainController'; SearchTypes=@("Domain Controller","Computer","OU") }
  Debug-Log ": DC dialog closed normally" -Type "Info"
}

## ==================== SINGLE UNIFIED MOVE/DELETE FUNCTION ====================
function Invoke-ObjectOperation {
  <#
  .SYNOPSIS
  Unified function for moving or deleting AD objects

  .PARAMETER Objects
  Object or array of objects to operate on

  .PARAMETER Operation
  'Move' or 'Delete'

  .PARAMETER IsBulk
  If true, shows bulk UI (for multiple objects)

  .EXAMPLE
  ## Single move
  Invoke-ObjectOperation -Objects @($user) -Operation 'Move'

  ## Bulk move
  Invoke-ObjectOperation -Objects $users -Operation 'Move' -IsBulk

  ## Delete
  Invoke-ObjectOperation -Objects @($user) -Operation 'Delete'
  #>

  param(
    [Parameter(Mandatory=$true)]
    [object[]]$Objects,

    [Parameter(Mandatory=$true)]
    [ValidateSet('Move', 'Delete')]
    [string]$Operation,

    [switch]$IsBulk
  )

  ## Convert single object to array if needed
  if ($Objects -isnot [array]) {
    $Objects = @($Objects)
  }

  ## ==================== DELETE OPERATION ====================
  if ($Operation -eq 'Delete') {
    foreach ($obj in $Objects) {
      ## Auto-detect object type
      $objectType = if ($obj.PSObject.Properties.Match('SamAccountName')) { 'User' }
                    elseif ($obj.PSObject.Properties.Match('Members')) { 'Group' }
                    elseif ($obj.PSObject.Properties.Match('Site')) { 'Domain Controller' }
                    elseif ($obj.PSObject.Properties.Match('ComputerType')) { 'Computer' }
                    elseif ($obj.PSObject.Properties.Match('distinguishedName') -and
                            $obj.distinguishedName -match '^OU=') { 'OU' }
                    else { 'Object' }

      $name = $obj.Name
      Debug-Log ": Delete requested for $objectType '$name'" -Type "Warn"

      ## Double confirmation for safety
      $result = Show-Modal "DELETE CONFIRMATION" "⚠️  WARNING: You are about to DELETE:`n`n Type: $objectType`n Name: $name`n`n This action CANNOT be undone!`n`n Are you absolutely sure?" -YesNo

      if ($result -ne 0) {
        Debug-Log ": Delete cancelled by user" -Type "Info"
        return $false
      }

      try {
        if ($Script:DemoMode) {
          ## Demo mode - remove from arrays
          switch ($objectType) {
            'User' {
              $Script:Users = $Script:Users | Where-Object { $_.Name -ne $name }
              if ($Script:rawUsers) { $Script:rawUsers = $Script:rawUsers | Where-Object { $_.Name -ne $name } }
            }
            'Group' {
              if ($Script:Groups) { $Script:Groups = $Script:Groups | Where-Object { $_.Name -ne $name } }
              if ($Script:rawDemoGroups) { $Script:rawDemoGroups = $Script:rawDemoGroups | Where-Object { $_.Name -ne $name } }
            }
            'Computer' {
              if ($Script:Computers) { $Script:Computers = $Script:Computers | Where-Object { $_.Name -ne $name } }
            }
            'OU' {
              if ($Script:rawOUs) { $Script:rawOUs = $Script:rawOUs | Where-Object { $_.Name -ne $name } }
            }
          }

          Debug-Log ": Deleted $objectType '$name' (demo mode)" -Type "Success"
          Show-Modal "Success" "$objectType '$name' deleted successfully (demo mode)"

        } else {
          ## Production mode - use AD cmdlets
          switch ($objectType) {
            'User'     { Remove-ADUser -Identity $obj.SamAccountName -Confirm:$false -ErrorAction Stop }
            'Group'    { Remove-ADGroup -Identity $obj.Name -Confirm:$false -ErrorAction Stop }
            'Computer' { Remove-ADComputer -Identity $obj.SamAccountName -Confirm:$false -ErrorAction Stop }
            'OU' {
              ## Check if OU is protected from deletion
              $adOU = Get-ADOrganizationalUnit -Identity $obj.DistinguishedName -Properties ProtectedFromAccidentalDeletion -ErrorAction Stop
              if ($adOU.ProtectedFromAccidentalDeletion) {
                Set-ADOrganizationalUnit -Identity $obj.DistinguishedName -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
              }
              Remove-ADOrganizationalUnit -Identity $obj.DistinguishedName -Confirm:$false -ErrorAction Stop
            }
            'Domain Controller' {
              Show-Modal "Not Supported" "Domain Controllers cannot be deleted from this interface"
              return $false
            }
            default { Remove-ADObject -Identity $obj.DistinguishedName -Confirm:$false -ErrorAction Stop }
          }

          Debug-Log ": Deleted $objectType '$name' from AD" -Type "Success"
          Show-Modal "Success" "$objectType '$name' deleted successfully"
        }

        ## Refresh UI
        Refresh-Data -domain $Script:CurrentDomain
        Build-Tree -domain $Script:CurrentDomain
        if ($Script:FilterStatusLabel) { Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel }
        return $true
      } catch {
        Debug-Log ": Failed to delete $objectType '$name': $($_.Exception.Message)" -Type "Error"
        Show-Modal "Delete Failed" "Failed to delete $objectType '$name':`n`n$($_.Exception.Message)"
        return $false
      }
    }
  }

  ## ==================== MOVE OPERATION ====================
  if ($Operation -eq 'Move') {
    ## Validate objects are moveable (users or groups only)
    foreach ($obj in $Objects) {
      $isUser = $obj.PSObject.Properties.Match('SamAccountName').Count -gt 0
      $isGroup = $obj.PSObject.Properties.Match('Members').Count -gt 0

      if (-not ($isUser -or $isGroup)) {
        Show-Modal "Not Supported" "Object '$($obj.Name)' cannot be moved (unsupported type)"
        return
      }
    }

    ## Get current OU (from first object)
    $currentOU = if ($Objects[0].DistinguishedName) {
      ($Objects[0].DistinguishedName -replace '^CN=[^,]+,', '')
    } else {
      $Objects[0].OU ?? "N/A"
    }

    ## Get list of available OUs
    $ouList = if ($Script:DemoMode) {
      $Script:Users | Where-Object OU | Select-Object -ExpandProperty OU -Unique | Sort-Object
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
      $title = "Move Object - $($Objects[0].Name)"
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

    ## Move button
    $btnMove = [Terminal.Gui.Button]::new((if ($IsBulk) { "Move All" } else { "Move" }))
    $btnMove.add_Clicked({
      if ($lstOU.SelectedItem -lt 0) {
        Show-Modal "Error" "Please select a target OU"
        return
      }

      $targetOU = $ouList[$lstOU.SelectedItem]

      if ($targetOU -eq $currentOU) {
        Show-Modal "Error" "Object$(if ($IsBulk) { 's are' } else { ' is' }) already in that OU"
        return
      }

      ## Confirm move
      $confirmMsg = if ($IsBulk) {
        "Move $($Objects.Count) object(s) to:`n$targetOU?"
      } else {
        "Move '$($Objects[0].Name)' to:`n$targetOU?"
      }

      $confirm = Show-Modal $(if ($IsBulk) { "Confirm Bulk Move" } else { "Confirm Move" }) $confirmMsg -YesNo
      if ($confirm -ne 0) { return }

      ## Perform the move
      $successCount = 0
      $failCount = 0
      $errors = @()

      foreach ($obj in $Objects) {
        $name = $obj.Name

        try {
          if ($Script:DemoMode) {
            ## Demo mode - update OU property
            if ($obj.PSObject.Properties.Match('OU')) {
              $obj.OU = $targetOU
            }
            $successCount++
            Debug-Log ": Moved $name to $targetOU (demo mode)" -Type "Info"
          } else {
            ## Production mode
            Move-ADObject -Identity $obj.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
            $successCount++
            Debug-Log ": Moved $name to $targetOU in AD" -Type "Info"
          }
        } catch {
          $failCount++
          $errors += "${name}: $($_.Exception.Message)"
          Debug-Log ": Failed to move ${name}: $($_.Exception.Message)" -Type "Error"
        }
      }

      ## Show results
      if ($IsBulk -or $failCount -gt 0) {
        $msg = "Successfully moved $successCount object(s)"
        if ($failCount -gt 0) {
          $msg += "`n`nFailed: $failCount"
          if ($errors.Count -gt 0 -and $errors.Count -le 5) {
            $msg += "`n`nErrors:`n" + ($errors -join "`n")
          }
        }
        Show-Modal $(if ($failCount -eq 0) { "Success" } else { "Move Complete" }) $msg
      } else {
        Show-Modal "Success" "Object moved successfully$(if (-not $Script:DemoMode) { '' } else { ' (demo mode)' })"
      }

      ## Refresh UI
      Refresh-Data -domain $Script:CurrentDomain
      Build-Tree -domain $Script:CurrentDomain
      if ($Script:FilterStatusLabel) {
        Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
      }

      [Terminal.Gui.Application]::RequestStop()

    }.GetNewClosure())

    $dlg.AddButton($btnMove)

    ## Cancel button
    $btnCancel = [Terminal.Gui.Button]::new("Cancel")
    $btnCancel.add_Clicked({ [Terminal.Gui.Application]::RequestStop() }).GetNewClosure()
    $dlg.AddButton($btnCancel)

    [Terminal.Gui.Application]::Run($dlg)
  }
}

## ==================== BULK ACCOUNT STATUS MANAGEMENT ====================
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
    Debug-Log ": Bulk $Action cancelled by user" -Type "Info"
    return
  }

  Debug-Log ": Starting bulk $Action on $($validObjects.Count) objects" -Type "Info"
  if ($Reason) {
    Debug-Log ":   Reason: $Reason" -Type "Info"
  }

  ## Process objects
  $successCount = 0
  $failCount = 0
  $errors = @()

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
          Debug-Log ":   ${Action}d $objectType '$name' (demo mode)" -Type "Info"

      } else {
        ## Production mode - use AD cmdlets
        if ($Action -eq 'Enable') {
          if ($objectType -eq 'User') { Enable-ADAccount -Identity $obj.SamAccountName -ErrorAction Stop
          } else { Enable-ADAccount -Identity $obj.SamAccountName -ErrorAction Stop }
        } else {
          if ($objectType -eq 'User') { Disable-ADAccount -Identity $obj.SamAccountName -ErrorAction Stop
          } else { Disable-ADAccount -Identity $obj.SamAccountName -ErrorAction Stop }
        }
        $successCount++
        Debug-Log ":   ${Action}d $objectType '$name' in AD" -Type "Success"
        }

      } catch {
        $failCount++
        $errors += "${name}: $($_.Exception.Message)"
        Debug-Log ":   Failed to $Action $objectType '$name': $($_.Exception.Message)" -Type "Error"
      }
    }

    ## Show results
    $msg = "Successfully ${Action}d $successCount account(s)"
    if ($failCount -gt 0) {
      $msg += "`n`nFailed: $failCount"
      if ($errors.Count -gt 0 -and $errors.Count -le 5) {
        $msg += "`n`nErrors:`n" + ($errors -join "`n")
      }
    }

    Show-Modal $(if ($failCount -eq 0) { "Success" } else { "$Action Complete" }) $msg

    ## Refresh UI
    Refresh-Data -domain $Script:CurrentDomain
    Build-Tree -domain $Script:CurrentDomain
    if ($Script:FilterStatusLabel) { Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel }

  Debug-Log ": Bulk $Action completed - $successCount succeeded, $failCount failed" -Type "Success"
}

## --------------------{ Common tab builders - re-usable across dialogs }--------------------
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
  } elseif ($Data.ObjectClass -eq 'user') {
    'User'
  } elseif ($Data.ObjectClass -eq 'group') {
    'Group'
  } elseif ($Data.ObjectClass -eq 'computer') {
    'Computer'
  } else {
    'Object'
  }
  $searchTypes = $Config.SearchTypes ?? @("$objectType", "Group", "User", "Computer", "OU")
  return @{
    Name = "Search/Lookup"
    Builder = {
      param($view, $data, $state)
      $y = 1
      ## Domain field
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Domain:" -FieldName 'txtSearchDomain' -State $state -Value $Script:CurrentDomain -Width 30
      ## Name field
      $nameLabel = "$objectType Name:"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label $nameLabel -FieldName 'txtSearchName' -State $state -Value $data.Name -Width 30
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
      $lblFilter.X = 48; $lblFilter.Y = 1
      $view.Add($lblFilter)
      $state.txtSearchFilter = [Terminal.Gui.TextField]::new("")
      $state.txtSearchFilter.X = 62; $state.txtSearchFilter.Y = 1; $state.txtSearchFilter.Width = 20
      $view.Add($state.txtSearchFilter)
      ## Filter handler
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
      ## Results label
      $lblResult = [Terminal.Gui.Label]::new("Properties:")
      $lblResult.X = 2; $lblResult.Y = $y
      $view.Add($lblResult)
      $y += 1
      ## Results text view
      $state.txtSearchOutput = [Terminal.Gui.TextView]::new()
      $state.txtSearchOutput.X = 2; $state.txtSearchOutput.Y = $y
      $state.txtSearchOutput.Width = [Terminal.Gui.Dim]::Fill(2)
      $state.txtSearchOutput.Height = [Terminal.Gui.Dim]::Fill(4)
      $state.txtSearchOutput.ReadOnly = $true
      $state.txtSearchOutput.WordWrap = $false
      $view.Add($state.txtSearchOutput)
      ## Object-specific checkboxes
      script:Set-ObjectCheckboxes -View $view -State $state -Data $data -ObjectType $objectType -Mode 'Create'
      ## Search button
      $btnSearch = [Terminal.Gui.Button]::new("Search")
      $btnSearch.X = 48; $btnSearch.Y = 3
      $view.Add($btnSearch)
      ## Capture the function for use in closure
      $setCheckboxesFunc = ${function:Set-ObjectCheckboxes}
      & $setCheckboxesFunc -State $state -Data $data -ObjectType $objectType -Mode 'Update'

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
          ## Update checkboxes - call via captured function
          & $updateCheckboxesFunc -State $state -Data $data -ObjectType $objectType
        }
      }.GetNewClosure())
    }
  }
}

## --------------------{ Helper functions for UI building }--------------------
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

  $lbl = [Terminal.Gui.Label]::new($Text)
  $lbl.X = $X
  $lbl.Y = $Y.Value
  $View.Add($lbl)

  $Y.Value += $SpaceAfter
}

## --------------------{ Common tab builders - re-usable across dialogs }--------------------

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
  } elseif ($Data.ObjectClass -eq 'user') {
    'User'
  } elseif ($Data.ObjectClass -eq 'group') {
    'Group'
  } elseif ($Data.ObjectClass -eq 'computer') {
    'Computer'
  } else {
    'Object'
  }

  $searchTypes = $Config.SearchTypes ?? @("$objectType", "Group", "User", "Computer", "OU")

  return @{
    Name = "Search/Lookup"
    Builder = {
      param($view, $data, $state)
      $y = 1

      ## Domain field
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Domain:" -FieldName 'txtSearchDomain' -State $state -Value $Script:CurrentDomain -Width 30

      ## Name field
      $nameLabel = "$objectType Name:"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label $nameLabel -FieldName 'txtSearchName' -State $state -Value $data.Name -Width 30

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
      $lblFilter.X = 48; $lblFilter.Y = 1
      $view.Add($lblFilter)

      $state.txtSearchFilter = [Terminal.Gui.TextField]::new("")
      $state.txtSearchFilter.X = 62; $state.txtSearchFilter.Y = 1; $state.txtSearchFilter.Width = 20
      $view.Add($state.txtSearchFilter)

      ## Filter handler
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

      ## Results label
      $lblResult = [Terminal.Gui.Label]::new("Properties:")
      $lblResult.X = 2; $lblResult.Y = $y
      $view.Add($lblResult)
      $y += 1

      ## Results text view
      $state.txtSearchOutput = [Terminal.Gui.TextView]::new()
      $state.txtSearchOutput.X = 2; $state.txtSearchOutput.Y = $y
      $state.txtSearchOutput.Width = [Terminal.Gui.Dim]::Fill(2)
      $state.txtSearchOutput.Height = [Terminal.Gui.Dim]::Fill(4)
      $state.txtSearchOutput.ReadOnly = $true
      $state.txtSearchOutput.WordWrap = $false
      $view.Add($state.txtSearchOutput)

      ## Object-specific checkboxes
      script:Set-ObjectCheckboxes -View $view -State $state -Data $data -ObjectType $objectType -Mode 'Create'

      ## Search button
      $btnSearch = [Terminal.Gui.Button]::new("Search")
      $btnSearch.X = 48; $btnSearch.Y = 3
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

## --------------------{ Helper funcitons for UI building }--------------------
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

  $lbl = [Terminal.Gui.Label]::new($Text)
  $lbl.X = $X
  $lbl.Y = $Y.Value
  $View.Add($lbl)

  $Y.Value += $SpaceAfter
}

function Show-GroupPropertiesDialog {
  param($group)

  if (-not $group) {
    Debug-Log ": Group object is null" -Type "Warn"
    return
  }
  Debug-Log ": Show-GroupPropertiesDialog starting for: $($group.Name)" -Type "Info"

  ## ==================== General Tab ====================
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $group, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Group Information"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Group name:" -FieldName 'txtName' -State $state -Value ($group.Name ?? "") -IsTextField $false
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Description:" -FieldName 'txtDescription' -State $state -Value ($group.Description ?? "")

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Group Details"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Type:" -FieldName 'lblType' -State $state -Value ($group.Type ?? "Security") -IsTextField $false
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Scope:" -FieldName 'lblScope' -State $state -Value ($group.Scope ?? "Global") -IsTextField $false

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Contact Information"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Email:" -FieldName 'txtEmail' -State $state -Value ($group.Email ?? "")
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Managed by:" -FieldName 'txtManagedBy' -State $state -Value ($group.ManagedBy ?? "")

      ## Audit Log button
      $btnAuditLog = [Terminal.Gui.Button]::new("View Audit Log...")
      $btnAuditLog.X = 2; $btnAuditLog.Y = $y
      $btnAuditLog.add_Clicked({
        Show-AuditLogDialog -Object $group -ObjectType 'Group'
      }.GetNewClosure())
      $view.Add($btnAuditLog)
    }
  }

  ## ==================== Members Tab ====================
  $membersTab = @{
    Name = "Members"
    Builder = {
      param($view, $group, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Group Members:")
      $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      $state.lstMembers = [Terminal.Gui.ListView]::new()
      $state.lstMembers.X = 2; $state.lstMembers.Y = $y
      $state.lstMembers.Width = [Terminal.Gui.Dim]::Fill(2)
      $state.lstMembers.Height = 25

      $state.memberList = @()
      if ($Script:DemoMode) {
        $state.memberList = $Script:Users | Where-Object { $_.Groups -contains $group.Name } | Select-Object -ExpandProperty Name | Sort-Object
      } else {
        try {
          $members = Get-ADGroupMember -Identity $group.Name -ErrorAction Stop
          $state.memberList = $members | Select-Object -ExpandProperty Name | Sort-Object
        } catch {
          Debug-Log ": Failed to get group members: $($_.Exception.Message)" -Type "Warn"
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

  ## ==================== Report Tab ====================
  $reportTab = @{
    Name = "Report"
    Builder = {
      param($view, $group, $state)

      ## Get detailed member information
      $memberDetails = @()
      if ($Script:DemoMode) {
        $members = $Script:Users | Where-Object { $_.Groups -contains $group.Name }
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
          $members = Get-ADGroupMember -Identity $group.Name -ErrorAction Stop
          foreach ($member in $members) {
            if ($member.objectClass -eq 'user') {
              $userDetails = Get-ADUser -Identity $member.SamAccountName -Properties EmailAddress,Department,Title,Enabled -ErrorAction SilentlyContinue
              if ($userDetails) {
                $memberDetails += [PSCustomObject]@{
                  Name = $userDetails.Name
                  SamAccountName = $userDetails.SamAccountName
                  Email = $userDetails.EmailAddress
                  Department = $userDetails.Department
                  Title = $userDetails.Title
                  Enabled = $userDetails.Enabled
                  Status = if ($userDetails.Enabled) { "Enabled" } else { "Disabled" }
                }
              }
            }
          }
        } catch {
          Debug-Log ": Failed to get member details: $($_.Exception.Message)" -Type "Warn"
        }
      }

      $y = 1

      $lblHeader = [Terminal.Gui.Label]::new("Membership Report: $($group.Name)")
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
        $filename = "group_members_$($group.Name)_$timestamp.csv"
        try {
          $memberDetails | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
          Show-Modal "Export Complete" "Membership report exported to:`n`n$filename"
          Debug-Log ": Exported group membership report to $filename" -Type "Success"
        } catch {
          Show-Modal "Export Failed" "Failed to export report:`n`n$($_.Exception.Message)"
        }
      }.GetNewClosure())
      $view.Add($btnExportFull)

      $btnCompare = [Terminal.Gui.Button]::new("Compare with Group...")
      $btnCompare.X = 25; $btnCompare.Y = $y
      $btnCompare.add_Clicked({
        $compareDlg = [Terminal.Gui.Dialog]::new("Compare Groups", 60, 20)
        $lblInfo = [Terminal.Gui.Label]::new("Select group to compare with $($group.Name):")
        $lblInfo.X = 2; $lblInfo.Y = 1
        $compareDlg.Add($lblInfo)

        ## Get other groups - ensure it's an array
        $otherGroups = @()
        if ($Script:DemoMode) {
          $otherGroups = @($Script:Groups | Where-Object { $_.Name -ne $group.Name } | Sort-Object Name | Select-Object -ExpandProperty Name)
        } else {
          try {
            $otherGroups = @(Get-ADGroup -Filter * | Where-Object { $_.Name -ne $group.Name } | Sort-Object Name | Select-Object -ExpandProperty Name)
          } catch {
            $otherGroups = @()
          }
        }

        Debug-Log ": Found $($otherGroups.Count) other groups for comparison" -Type "Debug"

        $lstGroups = [Terminal.Gui.ListView]::new()
        $lstGroups.X = 2; $lstGroups.Y = 3; $lstGroups.Width = 54; $lstGroups.Height = 12

        if ($otherGroups.Count -gt 0) {
          $lstGroups.SetSource($otherGroups)
        } else {
          $lstGroups.SetSource(@("(No other groups found)"))
        }

        $compareDlg.Add($lstGroups)

        $btnSelect = [Terminal.Gui.Button]::new("Compare")
        $btnSelect.add_Clicked({
          Debug-Log ": Compare button clicked, otherGroups count: $($otherGroups.Count), selected: $($lstGroups.SelectedItem)" -Type "Debug"

          if ($otherGroups.Count -eq 0) {
            Show-Modal "No Groups" "No other groups available for comparison"
            return
          }

          if ($lstGroups.SelectedItem -lt 0) {
            Show-Modal "No Selection" "Please select a group to compare with"
            return
          }

          ## Get the selected group name
          $compareGroupName = $otherGroups[$lstGroups.SelectedItem]
          Debug-Log ": Comparing $($group.Name) with $compareGroupName" -Type "Debug"

          [Terminal.Gui.Application]::RequestStop()

          ## Get members of current group
          $group1Members = @($memberDetails | Select-Object -ExpandProperty SamAccountName)

          ## Get members of comparison group
          $group2Members = @()
          if ($Script:DemoMode) {
            $group2Members = @($Script:Users | Where-Object { $_.Groups -contains $compareGroupName } | Select-Object -ExpandProperty SamAccountName)
          } else {
            try {
              $group2Members = @(Get-ADGroupMember -Identity $compareGroupName -ErrorAction Stop | Select-Object -ExpandProperty SamAccountName)
            } catch {
              Debug-Log ": Failed to get members of $compareGroupName : $($_.Exception.Message)" -Type "Error"
              $group2Members = @()
            }
          }

          $inBoth = @($group1Members | Where-Object { $group2Members -contains $_ })
          $onlyInGroup1 = @($group1Members | Where-Object { $group2Members -notcontains $_ })
          $onlyInGroup2 = @($group2Members | Where-Object { $group1Members -notcontains $_ })

          $resultMsg = "Comparison: $($group.Name) vs $compareGroupName`n`n"
          $resultMsg += "In both groups: $($inBoth.Count)`n"
          $resultMsg += "Only in $($group.Name): $($onlyInGroup1.Count)`n"
          $resultMsg += "Only in ${compareGroupName}: $($onlyInGroup2.Count)"

          Show-Modal "Group Comparison" $resultMsg
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

  ## ==================== Apply Logic ====================
  $applyLogic = {
    param($group, $state)
    try {
      $changesMade = $false

      if ($state.txtDescription) {
        $newDescription = $state.txtDescription.Text.ToString().Trim()
        if ($newDescription -ne $group.Description) {
          if ($Script:DemoMode) {
            $group.Description = $newDescription
          } else {
            Set-UnifiedObject -ObjectType Group -Object $group -Properties @{ Description = $newDescription }
          }
          $changesMade = $true
        }
      }

      if ($state.txtEmail) {
        $newEmail = $state.txtEmail.Text.ToString().Trim()
        if ($newEmail -ne $group.Email) {
          if ($Script:DemoMode) {
            $group.Email = $newEmail
          } else {
            Set-UnifiedObject -ObjectType Group -Object $group -Properties @{ mail = $newEmail; Email = $newEmail }
          }
          $changesMade = $true
        }
      }

      if ($state.txtManagedBy) {
        $newManagedBy = $state.txtManagedBy.Text.ToString().Trim()
        if ($newManagedBy -ne $group.ManagedBy) {
          if ($Script:DemoMode) {
            $group.ManagedBy = $newManagedBy
          } else {
            Set-UnifiedObject -ObjectType Group -Object $group -Properties @{ ManagedBy = $newManagedBy }
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

  ## ==================== Create Dialog ====================
  $tabs = @($generalTab, $membersTab, $reportTab)
  New-PropertiesDialog -Title "Group Properties - $($group.Name)" -Width 100 -Height 40 -Tabs $tabs -Data $group -OnApply $applyLogic -IncludeSearchTab $true -SearchTabConfig @{ObjectType='Group'}
}

# ----------------------- Show Computer Properties ----------------------
function Show-ComputerPropertiesDialog {
  param([string]$computerName)

  Debug-Log ": Showing computer properties for: $computerName" -Type "Info"
  $computer = $Script:Computers | Where-Object { $_.Name -eq $computerName } | Select-Object -First 1

  if (-not $computer) {
    Debug-Log ": Computer NOT found in Script:Computers for name: $computerName" -Type "Info"
    Show-Modal "Not Found" "Computer '$computerName' not found"
    return
  }

  Debug-Log ": Computer found: $($computer.Name)" -Type "Info"

  ## ==================== General Tab ====================
  $generalTab = @{
    Name = "General"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Computer Information"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Name:" -FieldName 'txtName' -State $state -Value ($computer.Name ?? "") -FieldX 25
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "DNS Host Name:" -FieldName 'txtDNS' -State $state -Value ($computer.DNSHostName ?? "") -FieldX 25 -IsTextField $false
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Domain:" -FieldName 'txtDomain' -State $state -Value ($computer.Domain ?? "") -FieldX 25 -IsTextField $false
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Description:" -FieldName 'txtDescription' -State $state -Value ($computer.Description ?? "") -FieldX 25 -Width 65

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Operating System"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "OS:" -FieldName 'txtOS' -State $state -Value ($computer.OperatingSystem ?? "") -FieldX 25 -IsTextField $false
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Version:" -FieldName 'txtOSVer' -State $state -Value ($computer.OperatingSystemVersion ?? "") -FieldX 25 -IsTextField $false

      if ($computer.OperatingSystemServicePack) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Service Pack:" -FieldName 'txtSP' -State $state -Value $computer.OperatingSystemServicePack -FieldX 25 -IsTextField $false
      }

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Network"
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "IPv4 Address:" -FieldName 'txtIPv4' -State $state -Value ($computer.IPv4Address ?? "") -FieldX 25 -IsTextField $false

      if ($computer.IPv6Address) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "IPv6 Address:" -FieldName 'txtIPv6' -State $state -Value $computer.IPv6Address -FieldX 25 -IsTextField $false
      }

      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Location:" -FieldName 'txtLocation' -State $state -Value ($computer.Location ?? "") -FieldX 25 -Width 65
    }
  }

  ## ==================== Account Tab ====================
  $accountTab = @{
    Name = "Account"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Account Status"

      $state.chkEnabled = [Terminal.Gui.CheckBox]::new("Computer Account Enabled")
      $state.chkEnabled.X=4; $state.chkEnabled.Y=$y; $state.chkEnabled.Checked=($computer.Enabled??$true)
      $view.Add($state.chkEnabled); $y+=1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Password Settings"

      $state.chkPasswordExpired = [Terminal.Gui.CheckBox]::new("Password Expired")
      $state.chkPasswordExpired.X=4; $state.chkPasswordExpired.Y=$y; $state.chkPasswordExpired.Checked=($computer.PasswordExpired??$false); $state.chkPasswordExpired.Enabled=$false
      $view.Add($state.chkPasswordExpired); $y+=1

      $state.chkPasswordNeverExpires = [Terminal.Gui.CheckBox]::new("Password never expires")
      $state.chkPasswordNeverExpires.X=4; $state.chkPasswordNeverExpires.Y=$y; $state.chkPasswordNeverExpires.Checked=($computer.PasswordNeverExpires??$false); $state.chkPasswordNeverExpires.Enabled=$false
      $view.Add($state.chkPasswordNeverExpires); $y+=1

      $state.chkCannotChangePassword = [Terminal.Gui.CheckBox]::new("Cannot change password")
      $state.chkCannotChangePassword.X=4; $state.chkCannotChangePassword.Y=$y; $state.chkCannotChangePassword.Checked=($computer.CannotChangePassword??$false); $state.chkCannotChangePassword.Enabled=$false
      $view.Add($state.chkCannotChangePassword); $y+=1

      $state.chkPasswordNotRequired = [Terminal.Gui.CheckBox]::new("Password not required")
      $state.chkPasswordNotRequired.X=4; $state.chkPasswordNotRequired.Y=$y; $state.chkPasswordNotRequired.Checked=($computer.PasswordNotRequired??$false); $state.chkPasswordNotRequired.Enabled=$false
      $view.Add($state.chkPasswordNotRequired); $y+=2

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Account Details" -SpaceBefore 0
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "SAM Account:" -FieldName 'txtSAM' -State $state -Value ($computer.SamAccountName ?? "") -FieldX 25 -IsTextField $false

      if ($computer.PasswordLastSet) {
        $lbl = [Terminal.Gui.Label]::new("Password last set: " + $computer.PasswordLastSet.ToString('yyyy-MM-dd HH:mm'))
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($computer.LastLogonDate) {
        $lbl = [Terminal.Gui.Label]::new("Last logon: " + $computer.LastLogonDate.ToString('yyyy-MM-dd HH:mm'))
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($computer.logonCount) {
        $lbl = [Terminal.Gui.Label]::new("Logon count: " + $computer.logonCount)
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($computer.AccountExpirationDate) {
        $lbl = [Terminal.Gui.Label]::new("Account expires: " + $computer.AccountExpirationDate.ToString('yyyy-MM-dd'))
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      $btnAuditLog = [Terminal.Gui.Button]::new("View Audit Log...")
      $btnAuditLog.X = 2; $btnAuditLog.Y = $y
      $btnAuditLog.add_Clicked({ Show-AuditLogDialog -Object $computer -ObjectType 'Computer' }.GetNewClosure())
      $view.Add($btnAuditLog)
    }
  }

  ## ==================== Security Tab ====================
  $securityTab = @{
    Name = "Security"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Delegation Settings"

      $state.chkTrustedForDelegation = [Terminal.Gui.CheckBox]::new("Trusted for delegation")
      $state.chkTrustedForDelegation.X=4; $state.chkTrustedForDelegation.Y=$y; $state.chkTrustedForDelegation.Checked=($computer.TrustedForDelegation??$false); $state.chkTrustedForDelegation.Enabled=$false
      $view.Add($state.chkTrustedForDelegation); $y+=1

      $state.chkTrustedToAuth = [Terminal.Gui.CheckBox]::new("Trusted to authenticate for delegation")
      $state.chkTrustedToAuth.X=4; $state.chkTrustedToAuth.Y=$y; $state.chkTrustedToAuth.Checked=($computer.TrustedToAuthForDelegation??$false); $state.chkTrustedToAuth.Enabled=$false
      $view.Add($state.chkTrustedToAuth); $y+=1

      $state.chkAccountNotDelegated = [Terminal.Gui.CheckBox]::new("Account not delegated")
      $state.chkAccountNotDelegated.X=4; $state.chkAccountNotDelegated.Y=$y; $state.chkAccountNotDelegated.Checked=($computer.AccountNotDelegated??$false); $state.chkAccountNotDelegated.Enabled=$false
      $view.Add($state.chkAccountNotDelegated); $y+=1

      $state.chkNoPreAuth = [Terminal.Gui.CheckBox]::new("Does not require Kerberos preauthentication")
      $state.chkNoPreAuth.X=4; $state.chkNoPreAuth.Y=$y; $state.chkNoPreAuth.Checked=($computer.DoesNotRequirePreAuth??$false); $state.chkNoPreAuth.Enabled=$false
      $view.Add($state.chkNoPreAuth); $y+=1

      $state.chkUseDES = [Terminal.Gui.CheckBox]::new("Use DES encryption types")
      $state.chkUseDES.X=4; $state.chkUseDES.Y=$y; $state.chkUseDES.Checked=($computer.UseDESKeyOnly??$false); $state.chkUseDES.Enabled=$false
      $view.Add($state.chkUseDES); $y+=2

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Kerberos Encryption" -SpaceBefore 0

      if ($computer.KerberosEncryptionType) {
        $kerbTypes = $computer.KerberosEncryptionType -join ', '
        $lbl = [Terminal.Gui.Label]::new("Encryption types: " + $kerbTypes)
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($computer.'msDS-SupportedEncryptionTypes') {
        $lbl = [Terminal.Gui.Label]::new("Supported encryption: " + $computer.'msDS-SupportedEncryptionTypes')
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Identifiers"

      $sid = $computer.SID ?? $computer.objectSid
      if ($sid) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "SID:" -FieldName 'txtSID' -State $state -Value $sid.ToString() -FieldX 25 -IsTextField $false
      }

      if ($computer.ObjectGUID) {
        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Object GUID:" -FieldName 'txtGUID' -State $state -Value $computer.ObjectGUID.ToString() -FieldX 25 -IsTextField $false
      }
    }
  }

  ## ==================== Member Of Tab ====================
  $memberOfTab = @{
    Name = "Member Of"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Group Memberships:")
      $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      $state.lstGroups = [Terminal.Gui.ListView]::new()
      $state.lstGroups.X = 2; $state.lstGroups.Y = $y
      $state.lstGroups.Width = [Terminal.Gui.Dim]::Fill(2)
      $state.lstGroups.Height = 20

      $state.groupList = @()
      if ($computer.MemberOf) {
        $state.groupList = $computer.MemberOf | ForEach-Object { if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ } }
      }

      if ($state.groupList.Count -gt 0) {
        $state.lstGroups.SetSource($state.groupList)
      } else {
        $state.lstGroups.SetSource(@("(No group memberships)"))
      }

      $view.Add($state.lstGroups)

      $pseudoUser = [PSCustomObject]@{
        Name = $computer.Name
        SamAccountName = $computer.SamAccountName
        MemberOf = $computer.MemberOf
        Groups = $state.groupList
      }

      $btnAdd = [Terminal.Gui.Button]::new("Add to Group...")
      $btnAdd.X = 2; $btnAdd.Y = 23
      $btnAdd.add_Clicked({
        Show-EditGroupMembershipDialog -User $pseudoUser -OnUpdate {
          $refreshedGroups = @()
          if ($computer.MemberOf) {
            $refreshedGroups = $computer.MemberOf | ForEach-Object {
              if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
            }
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
            Show-Modal "Info" "No group selected"
            return
          }

          $confirmDlg = Show-Modal "Confirm Removal" "Remove $($computer.Name) from group '$selectedGroup'?" -YesNo

          if ($confirmDlg -eq 0) {
            try {
              if ($Script:DemoMode) {
                $computer.MemberOf = $computer.MemberOf | Where-Object { $_ -notmatch "CN=$selectedGroup," }
              } else {
                Remove-ADGroupMember -Identity $selectedGroup -Members $computer.SamAccountName -Confirm:$false
              }

              $updatedGroups = @()
              if ($computer.MemberOf) {
                $updatedGroups = $computer.MemberOf | ForEach-Object {
                  if ($_ -match '^CN=([^,]+)') { $matches[1] } else { $_ }
                }
              }

              if ($updatedGroups.Count -gt 0) {
                $state.lstGroups.SetSource($updatedGroups)
              } else {
                $state.lstGroups.SetSource(@("(No group memberships)"))
              }

              $state.groupList = $updatedGroups
              Show-Modal "Success" "Successfully removed $($computer.Name) from group '$selectedGroup'"

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

  ## ==================== SPNs Tab ====================
  $spnTab = @{
    Name = "SPNs"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      $lbl = [Terminal.Gui.Label]::new("Service Principal Names:")
      $lbl.X=2; $lbl.Y=$y; $view.Add($lbl); $y+=2

      $state.lstSPNs = [Terminal.Gui.ListView]::new()
      $state.lstSPNs.X = 2; $state.lstSPNs.Y = $y
      $state.lstSPNs.Width = [Terminal.Gui.Dim]::Fill(2)
      $state.lstSPNs.Height = [Terminal.Gui.Dim]::Fill(2)

      $spnList = @()
      if ($computer.ServicePrincipalNames) {
        $spnList = $computer.ServicePrincipalNames
      } elseif ($computer.servicePrincipalName) {
        $spnList = $computer.servicePrincipalName
      }

      if ($spnList.Count -gt 0) {
        $state.lstSPNs.SetSource($spnList)
      } else {
        $state.lstSPNs.SetSource(@("(No SPNs configured)"))
      }

      $view.Add($state.lstSPNs)
    }
  }

  ## ==================== LAPS Tab ====================
  $lapsTab = @{
    Name = "LAPS"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Local Administrator Password Solution"

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
            } else { "Unknown" }
          }
        } catch {}
      } elseif ($computer.'ms-Mcs-AdmPwd') {
        $lapsEnabled = $true
        $lapsPassword = $computer.'ms-Mcs-AdmPwd'
        $lapsExpiry = if ($computer.'ms-Mcs-AdmPwdExpirationTime') {
          [DateTime]::FromFileTime($computer.'ms-Mcs-AdmPwdExpirationTime').ToString('yyyy-MM-dd HH:mm:ss')
        } else { "Unknown" }
      }

      if ($lapsEnabled) {
        $lbl = [Terminal.Gui.Label]::new("LAPS Status: ✓ Enabled")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=2

        $lbl = [Terminal.Gui.Label]::new("Administrator Password:")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1

        $state.txtPassword = [Terminal.Gui.TextField]::new($lapsPassword)
        $state.txtPassword.X=4; $state.txtPassword.Y=$y; $state.txtPassword.Width=70; $state.txtPassword.ReadOnly=$true
        $view.Add($state.txtPassword); $y+=2

        Add-LabelAndField -View $view -Y ([ref]$y) -Label "Password Expires:" -FieldName 'txtExpiry' -State $state -Value $lapsExpiry -FieldX 25 -IsTextField $false
        $y += 1

        $lbl = [Terminal.Gui.Label]::new("⚠ This password provides full local administrator access")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1

        $lbl = [Terminal.Gui.Label]::new("  Handle with appropriate security controls")
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

  ## ==================== BitLocker Tab ====================
  $bitlockerTab = @{
    Name = "BitLocker"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "BitLocker Recovery Information"

      $computerDN = $computer.DistinguishedName
      $recoveryKeys = @()

      if ($Script:rawBitLockerRecovery) {
        $recoveryKeys = $Script:rawBitLockerRecovery | Where-Object {
          $_.DN -match [regex]::Escape($computerDN)
        }
      }

      if ($recoveryKeys.Count -gt 0) {
        $lbl = [Terminal.Gui.Label]::new("Found $($recoveryKeys.Count) recovery key(s) for this computer")
        $lbl.X = 4; $lbl.Y = $y; $view.Add($lbl); $y += 2

        $state.lstRecoveryKeys = [Terminal.Gui.ListView]::new()
        $state.lstRecoveryKeys.X = 4; $state.lstRecoveryKeys.Y = $y
        $state.lstRecoveryKeys.Width = [Terminal.Gui.Dim]::Fill(2)
        $state.lstRecoveryKeys.Height = 8

        $keyList = $recoveryKeys | ForEach-Object {
          $created = if ($_.Created) { $_.Created.ToString('yyyy-MM-dd HH:mm') } else { 'Unknown' }
          "$($_.Name) - Created: $created"
        }
        $state.lstRecoveryKeys.SetSource($keyList)
        $view.Add($state.lstRecoveryKeys)
        $y += 9

        $btnView = [Terminal.Gui.Button]::new("View Selected Key...")
        $btnView.X = 4; $btnView.Y = $y
        $btnView.add_Clicked({
          $selectedIndex = $state.lstRecoveryKeys.SelectedItem
          if ($selectedIndex -ge 0 -and $selectedIndex -lt $recoveryKeys.Count) {
            $key = $recoveryKeys[$selectedIndex]
            $details = @"
BitLocker Recovery Key Details
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Computer: $($computer.Name)
Key Name: $($key.Name)

Recovery Password:
$($key.RecoveryPassword)

Recovery GUID: $($key.RecoveryGuid)
Volume GUID:   $($key.VolumeGuid)

Created: $($key.Created)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⚠ SECURITY WARNING ⚠

This recovery password provides full access to
encrypted data on this volume. Handle with extreme
care and store securely.

Do not share via email or unsecured channels.
"@
            Show-Modal "BitLocker Recovery Key" $details
          } else {
            Show-Modal "Info" "Please select a recovery key to view"
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
              if ($IsWindows) {
                Set-Clipboard -Value $key.RecoveryPassword
              } elseif ($IsLinux) {
                $key.RecoveryPassword | xclip -selection clipboard
              } elseif ($IsMacOS) {
                $key.RecoveryPassword | pbcopy
              }
              Show-Modal "Success" "Recovery password copied to clipboard"
            } catch {
              Show-Modal "Error" "Failed to copy to clipboard:`n$($_.Exception.Message)"
            }
          } else {
            Show-Modal "Info" "Please select a recovery key to copy"
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

  ## ==================== Advanced Tab ====================
  $advancedTab = @{
    Name = "Advanced"
    Builder = {
      param($view, $computer, $state)
      $y = 1

      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Distinguished Name:" -FieldName 'txtDN' -State $state -Value ($computer.DistinguishedName ?? "") -LabelX 2 -FieldX 2 -Width 90 -ReadOnly $true
      $y += 1
      Add-LabelAndField -View $view -Y ([ref]$y) -Label "Canonical Name:" -FieldName 'txtCN' -State $state -Value ($computer.CanonicalName ?? "") -LabelX 2 -FieldX 2 -Width 90 -ReadOnly $true

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Timestamps"

      $created = if ($computer.Created -or $computer.whenCreated) {
        ($computer.Created ?? $computer.whenCreated).ToString('yyyy-MM-dd HH:mm:ss')
      } else { "Unknown" }
      $lbl = [Terminal.Gui.Label]::new("Created: " + $created)
      $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1

      $modified = if ($computer.Modified -or $computer.whenChanged) {
        ($computer.Modified ?? $computer.whenChanged).ToString('yyyy-MM-dd HH:mm:ss')
      } else { "Unknown" }
      $lbl = [Terminal.Gui.Label]::new("Modified: " + $modified)
      $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1

      if ($computer.LastBootUpTime) {
        $lbl = [Terminal.Gui.Label]::new("Last Boot: " + $computer.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss'))
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Update Sequence Numbers"

      if ($computer.uSNCreated) {
        $lbl = [Terminal.Gui.Label]::new("USN Created: " + $computer.uSNCreated)
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($computer.uSNChanged) {
        $lbl = [Terminal.Gui.Label]::new("USN Changed: " + $computer.uSNChanged)
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      Add-SectionHeader -View $view -Y ([ref]$y) -Text "Additional Properties"

      if ($computer.isCriticalSystemObject) {
        $lbl = [Terminal.Gui.Label]::new("✓ Critical System Object")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($computer.ProtectedFromAccidentalDeletion) {
        $lbl = [Terminal.Gui.Label]::new("✓ Protected from Accidental Deletion")
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }

      if ($computer.PrimaryGroup) {
        $lbl = [Terminal.Gui.Label]::new("Primary Group: " + $computer.PrimaryGroup)
        $lbl.X=4; $lbl.Y=$y; $view.Add($lbl); $y+=1
      }
    }
  }

  ## ==================== Apply Logic ====================
  $applyLogic = {
    param($computer, $state)
    try {
      $changesMade = $false

      if ($state.txtDescription) {
        $newDescription = $state.txtDescription.Text.ToString().Trim()
        if ($newDescription -ne $computer.Description) {
          if ($Script:DemoMode) {
            $computer.Description = $newDescription
          } else {
            Set-UnifiedObject -ObjectType Computer -Object $computer -Properties @{ Description = $newDescription }
            $computer.Description = $newDescription
          }
          $changesMade = $true
        }
      }

      if ($state.txtLocation) {
        $newLocation = $state.txtLocation.Text.ToString().Trim()
        if ($newLocation -ne $computer.Location) {
          if ($Script:DemoMode) {
            $computer.Location = $newLocation
          } else {
            Set-UnifiedObject -ObjectType Computer -Object $computer -Properties @{ Location = $newLocation }
            $computer.Location = $newLocation
          }
          $changesMade = $true
        }
      }

      if ($changesMade) {
        Show-Modal "Success" "Computer changes applied successfully"
      } else {
        Show-Modal "Info" "No changes to apply"
      }
    } catch {
      Show-Modal "Error" "Failed to apply computer changes:`n$($_.Exception.Message)"
    }
  }

  ## ==================== Create Dialog ====================
  ## Note: Search tab auto-added!
  $tabs = @($generalTab, $accountTab, $securityTab, $memberOfTab, $spnTab, $lapsTab, $bitlockerTab, $advancedTab)

  New-PropertiesDialog -Title "Computer Properties - $($computer.Name)" -Width 100 -Height 40 -Tabs $tabs -Data $computer -OnApply $applyLogic -IncludeSearchTab $true -SearchTabConfig @{ObjectType='Computer'}
}

function Show-GPODetailsDialog {
  param($GPO)

  $gpoName = if ($GPO.DisplayName) { $GPO.DisplayName } else { $GPO.Name }
  Debug-Log ": Showing details for GPO: $gpoName" -Type "Info"
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
      if ($IsWindows) {
        Set-Clipboard -Value $details
      } elseif ($IsLinux) {
        $details | xclip -selection clipboard
      } elseif ($IsMacOS) {
        $details | pbcopy
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

## --------------------{ List dialog helpers - For GPO, DNS, Trusts, etc. }--------------------
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
  $dialog.Title = $Title
  $dialog.Width = $Width
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

## --------------------{ Cleaned up GPO list dialog }--------------------
function Show-GPOListDialog {
  param([string]$Domain)

  if (-not $Domain) { $Domain = $Script:CurrentDomain }
  Debug-Log ": Opening GPO list for domain: $Domain" -Type "Info"
  ## Get GPOs
  $gpoResult = Test-GPOHealth -Domain $Domain
  $gpos = $gpoResult.GPOs

  if (-not $gpos -or $gpos.Count -eq 0) {
    Show-Modal "No GPOs" "No Group Policy Objects found for domain: $Domain"
    return
  }

  Debug-Log ": Found $($gpos.Count) GPOs" -Type "Info"
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
      $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
      $filename = "gpo_list_${Domain}_$timestamp.csv"

      $exportData = $items | ForEach-Object {
        [PSCustomObject]@{
          Name = if ($_.DisplayName) { $_.DisplayName } else { $_.Name }
          Description = $_.Description
          Path = $_.GPCFileSysPath
          Version = $_.VersionNumber
          Created = $_.Created
          Modified = $_.Modified
          DN = $_.DN
        }
      }

      $exportData | Export-Csv -Path $filename -NoTypeInformation -Encoding UTF8
      Show-Modal "Success" "Exported $($items.Count) GPOs to:`n$filename"
      Debug-Log ": Exported GPO list to $filename" -Type "Success"
    } catch {
      Show-Modal "Error" "Failed to export GPO list:`n$($_.Exception.Message)"
      Debug-Log ": Export failed: $($_.Exception.Message)" -Type "Error"
    }
  }

  ## Refresh handler
  $onRefresh = { Show-GPOListDialog -Domain $Domain }
  ## Show dialog
  New-ListDialog -Title "Group Policy Objects - $Domain" -Items $gpos -FormatItem $formatGPO -OnView $onView -OnExport $onExport -OnRefresh $onRefresh -FilterHelp "(Filter by GPO name or description)"
}

## -------------------------[ Context Menu Handler }-------------------------
function Show-ContextMenu {
  param(
    [array]$menuItems,
    [object]$obj,
    [string]$objType
  )

  Debug-Log ": Showing context menu for $($obj.Name) (Type: $objType)" -Type "Info"

  ## Create dialog
  $contextDialog = [Terminal.Gui.Dialog]::new("Actions", 30, ($menuItems.Count + 4))
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
  Debug-Log ": Menu item selected: $selected" -Type "Info"
  [Terminal.Gui.Application]::RequestStop()

  if ($selected -ne "---") {
    switch ($selected) {
      "Properties" {
        switch ($objType) {
      'user' {
        Debug-Log ": Showing user properties for $($obj.Name)" -Type "Info"
        Show-UserPropertiesDialog -user $obj
      }
      'group' {
        Debug-Log ": Showing group properties for $($obj.Name)" -Type "Info"
        Show-GroupPropertiesDialog -group $obj
      }
      'computer' {
        Debug-Log ": Showing computer properties for $($obj.Name)" -Type "Info"
        Show-ComputerPropertiesDialog -computerName $obj.Name
      }
      'dc' {
        Debug-Log ": Showing DC properties for $($obj.Name)" -Type "Info"
        Show-DCPropertiesDialog -dc $obj
      }
      'ou' {
        Debug-Log ": Showing OU properties for $($obj.Name)" -Type "Info"
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
    if ($Script:selectionPanel) { Update-SelectionPanel -panel $Script:selectionPanel }
  }
}

## Additional debug helper - call this to verify your demo data loaded correctly
## This one not being called isn't criticla, it just needs to exist for when we do need it called
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

  ## Call this after loading demo data to verify:
  ## Test-DemoData
}

## -------------------------{ Helper Functions }-------------------------
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
        Debug-Log (": Password reset for $userName (demo mode)") -Type "Info"
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
      Refresh-Data -Domain $Script:CurrentDomain -RebuildTree -ShowModal -ShowLoadingDialog
    } catch {
      $errMsg = $_.Exception.Message
      Show-Modal "Error" "Failed to $action account:`n$errMsg"
    }
  }
}

## DSA-TUI Batch Operations Module v1.0
## Select multiple objects and perform bulk actions

## -------------------------{ Global Selection State }-------------------------
$Script:SelectedObjects = @()
$Script:SelectionMode = $false

## -------------------------{ Toggle Selection Mode }-------------------------
function Toggle-SelectionMode {
  $Script:SelectionMode = -not $Script:SelectionMode

  if ($Script:SelectionMode) {
    Debug-Log (": Selection mode ENABLED") -Type "Info"
    Show-Modal "Selection Mode" "Selection mode enabled!`n`nClick objects to select/deselect them.`nPress Ctrl+A to select all.`nPress Ctrl+D to deselect all."
  } else {
    Debug-Log (": Selection mode DISABLED") -Type "Info"
    $Script:SelectedObjects = @()
    Build-Tree -domain $Script:CurrentDomain
    Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel
  }
}

## -------------------------{ Enhanced Tree with Selection Support }-------------------------
## TODO: Never called - confirm if it's still needed or is dead code for removal
function Handle-TreeClick {
  param($mouseArgs)

  if (-not $Script:tree.SelectedObject) { return }
  $selName = $Script:tree.SelectedObject.Text
  ## Check if in selection mode
  if ($Script:SelectionMode) {
    ## Toggle selection
    if ($Script:SelectedObjects -contains $selName) {
      ## Deselect
      $Script:SelectedObjects = $Script:SelectedObjects | Where-Object { $_ -ne $selName }
      Debug-Log (": Deselected $selName") -Type "Info"
    } else {
      ## Select
      $Script:SelectedObjects += $selName
      Debug-Log (": Selected $selName") -Type "Info"
    }

    ## Update visual indicator (mark selected items)
    Update-SelectionPanel -panel $selectionPanel
    $mouseArgs.Handled = $true
  }
}

## Add keyboard shortcuts for selection TODO: This is never called and needs to be
function Add-SelectionKeyBindings {
  param($view)

  $view.add_KeyPress({ param($senders, $keyArgs)
    ## Ctrl+A = Select All
    if ($keyArgs.KeyEvent.Key -eq ([Terminal.Gui.Key]::A -bor [Terminal.Gui.Key]::CtrlMask)) {
      Manage-Selection -Action 'SelectAll'
      $keyArgs.Handled = $true
    }

    ## Ctrl+D = Deselect All
    if ($keyArgs.KeyEvent.Key -eq ([Terminal.Gui.Key]::D -bor [Terminal.Gui.Key]::CtrlMask)) {
      Manage-Selection -Action 'DeselectAll'
      $keyArgs.Handled = $true
    }

    ## Ctrl+S = Toggle Selection Mode
    if ($keyArgs.KeyEvent.Key -eq ([Terminal.Gui.Key]::S -bor [Terminal.Gui.Key]::CtrlMask)) {
      Toggle-SelectionMode
      $keyArgs.Handled = $true
    }
  })
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

    Debug-Log ": Selected all users ($($Script:SelectedObjects.Count))" -Type "Info"

    if ($Script:selectionPanel) { Update-SelectionPanel -panel $Script:selectionPanel }
    Show-Modal "Selected All" "Selected $($Script:SelectedObjects.Count) users"
  }

  if ($Action -eq 'DeselectAll') {
    $Script:SelectedObjects = @()
    Debug-Log ": Deselected all objects" -Type "Info"

    if ($Script:selectionPanel) {
      Update-SelectionPanel -panel $Script:selectionPanel
    }
  }
}

## -------------------------[ Bulk Disable/Enable }-------------------------
function Invoke-BulkDisableEnable {
  param([bool]$disable)

  if ($Script:SelectedObjects.Count -eq 0) {
    Show-Modal "No Selection" "No objects selected. Select objects first."
    return
  }

  $action = if ($disable) { "disable" } else { "enable" }
  $result = Show-Modal "Confirm Bulk Action" "Are you sure you want to $action $($Script:SelectedObjects.Count) user account(s)?" -YesNo

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

  ## Show results
  $msg = "Successfully $action`d $successCount account(s)"
  if ($failCount -gt 0) {
    $msg += "`n`nFailed: $failCount"
    if ($errors.Count -gt 0 -and $errors.Count -le 5) { $msg += "`n`nErrors:`n" + ($errors -join "`n") }
  }

  Show-Modal "Bulk Action Complete" $msg

  ## Refresh tree
  Initialise-DataSource -Domain $Script:CurrentDomain
  Show-InfoPanel -UpdateOnly
  Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel

  ## Clear selection
  $Script:SelectedObjects = @()
  $Script:SelectionMode = $false
  Update-SelectionPanel -panel $selectionPanel
  }
}

## -------------------------{ Bulk Add to Group }-------------------------
function Invoke-BulkAddToGroup {
  if ($Script:SelectedObjects.Count -eq 0) {
    Show-Modal "No Selection" "No objects selected. Select objects first."
    return
  }

  $dlg = [Terminal.Gui.Dialog]::new("Bulk Add to Group", 70, 18)
  $lblInfo = [Terminal.Gui.Label]::new("Add $($Script:SelectedObjects.Count) user(s) to group:")
  $lblInfo.X=2; $lblInfo.Y=1; $dlg.Add($lblInfo)

  ## Get list of groups
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
  $confirm = Show-Modal "Confirm Bulk Add" "Add $($Script:SelectedObjects.Count) user(s) to group:`n$targetGroup?" -YesNo

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

## Add keyboard shortcuts for selection
## Never callled - confirm if it needs to be, or is dead code for removal
function Add-SelectionKeyBindings {
  param($view)

  $view.add_KeyPress({ param($senders, $keyArgs)
  ## Ctrl+A = Select All
  if ($keyArgs.KeyEvent.Key -eq ([Terminal.Gui.Key]::A -bor [Terminal.Gui.Key]::CtrlMask)) {
    Manage-Selection -Action 'SelectAll'
    $keyArgs.Handled = $true
  }

  ## Ctrl+D = Deselect All
  if ($keyArgs.KeyEvent.Key -eq ([Terminal.Gui.Key]::D -bor [Terminal.Gui.Key]::CtrlMask)) {
    Manage-Selection -Action 'DeselectAll'
    $keyArgs.Handled = $true
  }

  ## Ctrl+S = Toggle Selection Mode
  if ($keyArgs.KeyEvent.Key -eq ([Terminal.Gui.Key]::S -bor [Terminal.Gui.Key]::CtrlMask)) {
    Toggle-SelectionMode
    $keyArgs.Handled = $true
    }
  })
}

## -------------------------{ Bulk Disable/Enable }-------------------------
function Invoke-BulkDisableEnable {
  param([bool]$disable)

  if ($Script:SelectedObjects.Count -eq 0) {
    Show-Modal "No Selection" "No objects selected. Select objects first."
    return
  }

  $action = if ($disable) { "disable" } else { "enable" }
  $result = Show-Modal "Confirm Bulk Action" "Are you sure you want to $action $($Script:SelectedObjects.Count) user account(s)?" -YesNo

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
        if ($disable) { Disable-ADAccount -Identity $cleanName -ErrorAction Stop
        } else { Enable-ADAccount -Identity $cleanName -ErrorAction Stop
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

  ## Show results
  $msg = "Successfully $action`d $successCount account(s)"
  if ($failCount -gt 0) {
    $msg += "`n`nFailed: $failCount"
    if ($errors.Count -gt 0 -and $errors.Count -le 5) {
      $msg += "`n`nErrors:`n" + ($errors -join "`n")
    }
  }

  Show-Modal "Bulk Action Complete" "$msg"
  ## Refresh tree
  Build-Tree -domain $Script:CurrentDomain
  Manage-FilterStatusLabel -Action 'Update' -Label $Script:FilterStatusLabel

  ## Clear selection
  $Script:SelectedObjects = @()
  $Script:SelectionMode = $false
  Update-SelectionPanel -panel $selectionPanel
  }
}

## -------------------------{ Bulk Add to Group }-------------------------
function Invoke-BulkAddToGroup {
  if ($Script:SelectedObjects.Count -eq 0) {
    Show-Modal "No Selection" "No objects selected. Select objects first."
    return
  }

  $dlg = [Terminal.Gui.Dialog]::new("Bulk Add to Group", 70, 18)
  $lblInfo = [Terminal.Gui.Label]::new("Add $($Script:SelectedObjects.Count) user(s) to group:")
  $lblInfo.X=2; $lblInfo.Y=1; $dlg.Add($lblInfo)

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

## -------------------------{ Selection Panel }-------------------------
function Create-SelectionPanel {
  param(
    [int]$panelWidth  = 30,
    [int]$panelHeight = 10
  )

  ## Create frame view
  $selPanel = [Terminal.Gui.FrameView]::new("Selected Objects")
  $selPanel.Width  = $panelWidth
  $selPanel.Height = $panelHeight
  $selPanel.X = [Terminal.Gui.Pos]::AnchorEnd($panelWidth)  # Right-align
  $selPanel.Y = 1  # Just below top margin

  ## Count label
  $lblCount = [Terminal.Gui.Label]::new("0 objects selected")
  $lblCount.X = 1
  $lblCount.Y = 0
  $selPanel.Add($lblCount)

  ## ListView
  $lstSelected = [Terminal.Gui.ListView]::new(@())
  $lstSelected.X = 1
  $lstSelected.Y = 1
  $lstSelected.Width  = [Terminal.Gui.Dim]::Fill(2)  # margin on both sides
  $lstSelected.Height = [Terminal.Gui.Dim]::Fill(6)  # leaves space for buttons
  $selPanel.Add($lstSelected)

  ## Store references in Tag
  $selPanel | Add-Member -MemberType NoteProperty -Name Tag -Value @{
    CountLabel = $lblCount
    ListView   = $lstSelected
  } -Force

  ## ----------------- Batch action buttons -----------------
  ## TODO: You can merge this and enable all and check the items status couldn't you....?
  $yPos = [Terminal.Gui.Pos]::Bottom($lstSelected) + 1

  $btnBulkDisable = [Terminal.Gui.Button]::new("Disable")
  $btnBulkDisable.X = 2
  $btnBulkDisable.Y = $yPos
  $btnBulkDisable.add_Clicked({ Invoke-BulkDisableEnable -disable $true }).GetNewClosure()
  $selPanel.Add($btnBulkDisable)

  #$yPos = [Terminal.Gui.Pos]::Bottom($btnBulkDisable) + 1
  $btnBulkEnable = [Terminal.Gui.Button]::new("Enable")
  $btnBulkEnable.X = 15
  $btnBulkEnable.Y = $yPos
  $btnBulkEnable.add_Clicked({ Invoke-BulkDisableEnable -disable $false }).GetNewClosure()
  $selPanel.Add($btnBulkEnable)

  ##$yPos = [Terminal.Gui.Pos]::Bottom($btnBulkEnable) + 1
  $btnBulkMove = [Terminal.Gui.Button]::new("Move")
  $btnBulkMove.X = 27
  $btnBulkMove.Y = $yPos
  $btnBulkMove.add_Clicked({ Invoke-BulkMove }).GetNewClosure()
  $selPanel.Add($btnBulkMove)
  return $selPanel
}

## -------------------------{ Update Selection Panel }-------------------------
function Update-SelectionPanel {
  param($panel)

  if (-not $panel -or -not $panel.Tag) { return }
  $lblCount = $panel.Tag.CountLabel
  $lstSelected = $panel.Tag.ListView
  $count = $Script:SelectedObjects.Count
  $lblCount.Text = "$count object(s) selected"
  $displayNames = $Script:SelectedObjects | ForEach-Object { $name = $_ -replace '^\(.\)\s*', '' -replace '^[○⊗]\s*', '' ; $name }
  $lstSelected.SetSource($displayNames)
  $panel.SetNeedsDisplay()
}

## -------------------------{ Add / Remove group members aka edit }-------------------------
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

    ## Remove from groups
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
  $btnCancel = [Terminal.Gui.Button]::new("Cancel")
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
----------------------------------------
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

  Debug-Log ": Showing OU properties dialog" -Type "Info"

  if (-not $ou) {
    Debug-Log ": OU object is null" -Type "Warn"
    return
  }

  $ouName = $ou.Name ?? $ou.Text ?? "Unknown"
  Debug-Log ": OU name resolved to: $ouName" -Type "Info"

  ## If not found, create a basic one
  if (-not $ou.Name) {
    $ou = @{
      Name = $ouName
      Path = ""
      Description = ""
    }
  }

  ## ==================== General Tab ====================
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

  ## ==================== Statistics Tab ====================
  $statisticsTab = @{
    Name = "Statistics"
    Builder = {
      param($view, $ou, $state)

      ## Calculate statistics
      $ouName = $ou.Name
      $usersInOU     = $Script:Users | Where-Object { $_.OU -contains $ouName }
      $enabledUsers  = ($usersInOU | Where-Object { $_.Enabled -eq $true }).Count
      $disabledUsers = ($usersInOU | Where-Object { $_.Disabled -eq $true }).Count
      $lockedUsers   = ($usersInOU | Where-Object { $_.LockedOut -eq $true }).Count

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
          Debug-Log ": Exported OU statistics to $filename" -Type "Success"
        } catch {
          Show-Modal "Export Failed" "Failed to export statistics:`n`n$($_.Exception.Message)"
          Debug-Log ": Failed to export OU statistics: $($_.Exception.Message)" -Type "Error"
        }
      }.GetNewClosure())
      $view.Add($btnExport)
    }
  }

  ## ==================== Apply Logic ====================
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

  ## ==================== Create Dialog ====================
  ## Note: OUs don't get a search tab
  $tabs = @($generalTab, $statisticsTab)
  Debug-Log ": Show-OUPropertiesDialog running" -Type "Info"
  New-PropertiesDialog -Title "OU Properties - $ouName" -Width 80 -Height 28 -Tabs $tabs -Data $ou -OnApply $applyLogic -IncludeSearchTab $false
  Debug-Log ": Show-OUPropertiesDialog completed" -Type "Info"
}

## --------------------{ Program Launch Begins Here }--------------------
Get-Command Debug-Log, script:Set-ObjectCheckboxes -All | Format-Table Name, CommandType, Source

## ===== STEP 1: Environment & Logging =====

## Capture Set-ObjectCheckboxes for use in closures (Needed here, outside of helper functions)
$Script:SetObjectCheckboxes_Func = ${function:Set-ObjectCheckboxes}

## Don't let users do stupid stuff
if ($DemoMode -and $PSBoundParameters.ContainsKey('Domain')) {
  Debug-Log "Invalid startup: -Domain cannot be used with -DemoMode" -Type "Error"
  return
}

## Echo basic info for debugging
Debug-Log "DemoMode: $DemoMode" -Type "info"
Debug-Log "Logging: $Logging" -Type "Info"
Debug-Log "LogFile: $LogFile" -Type "Info"

## Initialise logging if requested
if ($Logging -or $LogFile) {
  Debug-Log "Logging condition TRUE" "Debug"

  if (-not $LogFile) {
    $LogFile = Join-Path $PSScriptRoot "dsa_tui_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Debug-Log "Auto-generated LogFile: $LogFile" -Type "Debug"
  }

  if (-not [System.IO.Path]::IsPathRooted($LogFile)) { $LogFile = Join-Path (Get-Location).Path $LogFile }

  $Script:Logging = $true
  $Script:LogFile = $LogFile
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

## Set globals
$Script:DemoMode  = $DemoMode
$Script:ThemeMode = $Theme

## CSV import tracking - CRITICAL FLAGS
$Script:CSVDataLoaded = $false
$Script:CSVDataPath   = $null

## ===== STEP 2: Module Checks & Terminal.Gui =====
Debug-Log "Performing pre-flight module checks..." -Type "Info"

## Required module
$requiredOK = Test-RequiredModule -Name "Microsoft.PowerShell.ConsoleGuiTools"
if (-not $requiredOK) {
  Debug-Log "Missing required module Microsoft.PowerShell.ConsoleGuiTools. Exiting." -Type "Error"
  exit
}

## Check all modules ONCE at startup
$Script:hasConsoleTools    = Test-RequiredModule -Name 'Microsoft.PowerShell.ConsoleGuiTools'
$Script:HasPSWriteColor    = Test-RequiredModule -Name "PSWriteColor" -Optional
$Script:HasTerminalIcons   = Test-RequiredModule -Name "Terminal-Icons" -Optional
$Script:hasNerdFonts       = Test-RequiredModule -Name 'NerdFonts' -Optional
$Script:HasActiveDirectory = Test-RequiredModule -Name "ActiveDirectory" -Optional
$Script:HasGroupPolicy     = Test-RequiredModule -Name "GroupPolicy" -Optional
$Script:HasDNSServer       = Test-RequiredModule -Name "DnsServer" -Optional
$Script:HasDFSR            = Test-RequiredModule -Name "DFSR" -Optional

## Lack of AD module defaults to demo mode
if (-not $Script:HasActiveDirectory) {
  Debug-Log "ActiveDirectory module missing. Falling back to DEMO mode..." -Type "Warn"
  $Script:DemoMode = $true
}
Debug-Log "Module availability check complete" -Type "Info"
$Script:UseIcons = $false
if ($Script:HasTerminalIcons) { try { Write-Host '' -NoNewline; $Script:UseIcons = $true } catch {} }

## Ensure Terminal.Gui.dll is loaded
Debug-Log "Checking Terminal.Gui assembly..." -Type "Debug"
if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Terminal.Gui' })) {
  if ($Script:hasConsoleTools) {
    $mod = Get-Module -Name 'Microsoft.PowerShell.ConsoleGuiTools'
    $dll = Join-Path $mod.ModuleBase 'Terminal.Gui.dll'
    if (Test-Path $dll) {
      Add-Type -Path $dll -ErrorAction Stop
      Debug-Log "Loaded Terminal.Gui from $dll" -Type "Debug"
    } else {
      Debug-Log "Terminal.Gui.dll not found in $($mod.ModuleBase). Install Microsoft.PowerShell.ConsoleGuiTools." -Type "Error"
      return
    }
  } else {
    Debug-Log "Microsoft.PowerShell.ConsoleGuiTools module not found." -Type "Error"
    return
  }
} else {
  Debug-Log "Terminal.Gui assembly already loaded." -Type "Info"
}

## ===== STEP 3: Initialise Terminal.Gui UI =====
Initialise-DirectoryEmoji
$windowTitle = "$($Script:ProjectName) $($Script:DirectoryEmoji) Active Directory $BuildVersion Codename: $($Script:FruitName)"

$uiComponents = Initialise-UIFramework -Theme $Theme -Title $windowTitle
$top = $uiComponents.Top
$win = $uiComponents.Window
$Script:themeData = $uiComponents.Theme

Debug-Log "UI Framework ready - window visible to user" -Type "Success"
Get-Theme -Dump $Script:themeData

## ===== STEP 4: Forest/Domain Initialization =====
Debug-Log "Initializing forest/domain globals..." -Type "Info"

## Tab & layout placeholders
$Script:LayoutInProgress = $false
$Script:TabRows = @()
$Script:AllTabs = @()
$Script:ActiveTab = $null
$Script:TabRowHeight = 1

if ($Script:DemoMode) {
  Debug-Log "DemoMode enabled: creating demo forest structure..." -Type "Info"
  $Script:ForestName = "jukebox.example"
  $Script:RootDomain = "example.com"
  $Script:Domains = @('example.com', 'example.net', 'example.org')
  $Script:CurrentDomain = $Script:RootDomain

  ## Handle CSV import if requested
  if ($ImportDemoData) {
    if (-not $DemoDataCsv) {
      Debug-Log "ImportDemoData specified but no CSV file was provided (-DemoDataCsv)." -Type "Error"
      throw "Demo data import aborted: CSV file not specified."
    }

    if (-not (Test-Path -LiteralPath $DemoDataCsv)) {
      Debug-Log "Demo data CSV not found: $DemoDataCsv" -Type "Error"
      throw "Demo data import aborted: CSV file: $DemoDataCsv does not exist."
    }

    Debug-Log "Importing demo data from CSV: $DemoDataCsv" -Type "Info"
    Initialise-DataSource -CSVPath $DemoDataCsv

    ## Set flags immediately after import
    $Script:CSVDataLoaded = $true
    $Script:CSVDataPath = $DemoDataCsv
    Show-InfoPanel -UpdateOnly
  }

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
    $Script:RootDomain = $forest.RootDomain
    $Script:Domains = $forest.Domains
    $Script:Sites = $forest.Sites | ForEach-Object { $_.Name }
    $Script:CurrentDomain = if ($Domain) { $Domain } else { $Script:RootDomain }
  } catch {
    Debug-Log ("Failed to query AD domain/forest: $_") -Type "Error"
    Debug-Log ("Falling back to minimal domain info.") -Type "Warn"
    $Script:ForestName = "DOMAIN"
    $Script:RootDomain = if ($Domain) { $Domain } else { $env:USERDNSDOMAIN }
    $Script:Domains = @($Script:RootDomain)
    $Script:Sites = @()
    $Script:CurrentDomain = $Script:RootDomain
  }
}

$Script:Domain = $Script:CurrentDomain  ## Compatibility

## Initialise object arrays ONLY if CSV wasn't loaded
if (-not $Script:CSVDataLoaded) {
  Debug-Log "Initializing empty object arrays..." -Type "Debug"
  $Script:CurrentDC       = $null
  $Script:Users           = @()
  $Script:Groups          = @()
  $Script:DCs             = @()
  $Script:Computers       = @()
  $Script:ADObjects       = @()
  $Script:SelectedObjects = @()
  $Script:SelectionMode   = $false
} else {
  Debug-Log "Skipping array initialization - CSV data already loaded" -Type "Info"
}

## Global Search filters
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

## ===== STEP 5: Create Status Bar =====
Debug-Log "Creating status bar..." -Type "Info"
$statusBar = Set-StatusBar -Initialise -ThemeData $Script:themeData
$top.Add($statusBar)

## ===== STEP 6: Load Domain Data =====
## Only load if CSV wasn't already loaded
if (-not $Script:CSVDataLoaded) {
  Debug-Log "Loading domain data for $($Script:CurrentDomain)..." -Type "Info"
  Set-StatusBar "Loading domain data for $($Script:CurrentDomain)..." -Spinner
  Initialise-DataSource -Domain $Script:CurrentDomain
  Set-StatusBar "Ready" -Final
} else {
  Debug-Log "Skipping data load - CSV data already loaded" -Type "Info"
  Set-StatusBar "Ready" -Final
}

## Set Current DC after data is loaded
if ($Script:DCs -and $Script:DCs.Count -gt 0) {
  ## Prefer Global Catalog DC
  $Script:CurrentDC = $Script:DCs | Where-Object { $_.IsGlobalCatalog } | Select-Object -First 1
  if (-not $Script:CurrentDC) { $Script:CurrentDC = $Script:DCs | Select-Object -First 1 }
  Debug-Log ": Set current DC to: $($Script:CurrentDC.Name)" -Type "Info"
} else {
  Debug-Log ": No DCs available to set as CurrentDC" -Type "Warn"
  $Script:CurrentDC = $null
}

Debug-Log "POST-LOAD: Users=$($Script:Users.Count), DCs=$($Script:DCs.Count), Computers=$($Script:Computers.Count), Group=$($Script:Groups.Count), Objects=$($Script:ADObjects.Count)" -Type "Info"
Debug-Log "Forest/Domain initialization complete: CurrentDomain=$($Script:CurrentDomain)" -Type "Info"

## ===== STEP 7: Build UI Components =====

Debug-Log ": Creating main menu..." -Type "Info"
$menu = Build-MainMenu
$top.Add($menu)

Debug-Log ": Creating filter panel..." -Type "Info"
$filterPanel = Create-FilterPanel
if (-not ($filterPanel -is [Terminal.Gui.View])) { $filterPanel = [Terminal.Gui.FrameView]::new("Filters") }
$filterPanel.Width = 40
$filterPanel.Height = 26
$filterPanel.X = [Terminal.Gui.Pos]::AnchorEnd(40)
$filterPanel.Y = 0
$win.Add($filterPanel)

Debug-Log ": Creating selection panel..." -Type "Info"
$selectedObjectsPanel = Create-SelectionPanel
if (-not ($selectedObjectsPanel -is [Terminal.Gui.View])) { $selectedObjectsPanel = [Terminal.Gui.FrameView]::new("Selected Objects") }
$selectedObjectsPanel.Width = 40
$selectedObjectsPanel.Height = 5
$selectedObjectsPanel.X = [Terminal.Gui.Pos]::AnchorEnd(40)
$selectedObjectsPanel.Y = 26
$win.Add($selectedObjectsPanel)

Debug-Log ": Creating Info panel..." -Type "Info"
## First run - DO NOT update!
$InfoPanel = Show-InfoPanel
if (-not ($InfoPanel -is [Terminal.Gui.View])) { $InfoPanel = [Terminal.Gui.FrameView]::new("Active Directory Info") }
$InfoPanel.Width = 40
$InfoPanel.Height = 12
$InfoPanel.X = [Terminal.Gui.Pos]::AnchorEnd(40)
$InfoPanel.Y = 31
$win.Add($InfoPanel)
## NOW update it with the current DC
Show-InfoPanel -UpdateOnly
Debug-Log ": InfoPanel created and updated with DC: $($Script:CurrentDC.Name ?? 'None')" -Type "Debug"


Debug-Log ": Initializing TreeView..." -Type "Info"
Set-StatusBar "Building tree view..." -Spinner

$treeFrame = [Terminal.Gui.FrameView]::new("Active Directory Objects")
$treeFrame.X = 0
$treeFrame.Y = 0
$treeFrame.Width = [Terminal.Gui.Dim]::Fill(42)
$treeFrame.Height = [Terminal.Gui.Dim]::Fill()

$Script:tree = [Terminal.Gui.TreeView]::new()
$Script:tree.X = 0
$Script:tree.Y = 0
$Script:tree.Width = [Terminal.Gui.Dim]::Fill()
$Script:tree.Height = [Terminal.Gui.Dim]::Fill()

## Right-click context menu handler
$Script:tree.add_MouseClick({
  param($senders, $arguments)
  if ($arguments.MouseEvent.Flags -band [Terminal.Gui.MouseFlags]::Button3Clicked) {
    Debug-Log ": Right-click detected on tree" -Type "Info"
    $selectedNode = $Script:tree.SelectedObject
    if ($null -eq $selectedNode) {
      Debug-Log ": No node selected" -Type "Info"
      return
    }
    $tag = $selectedNode.Tag
    if (-not $tag -or -not $tag.Object) {
      Debug-Log ": Container/OU selected (no context menu)" -Type "Info"
      return
    }
    $obj = $tag.Object
    $objType = $tag.Type
    Debug-Log ": Right-click on object: $($obj.Name), Type: $objType" -Type "Info"
    $menuItems = Build-ContextMenuItems -ObjectType $objType -Object $obj
    Show-ContextMenu -menuItems $menuItems -obj $obj -objType $objType
  }
})

## Build and populate tree
$rootNode = Build-Tree -domain $Script:CurrentDomain
if ($null -ne $rootNode) {
  $Script:tree.ClearObjects()
  $Script:tree.AddObject($rootNode)
  Debug-Log ": Root node added to TreeView" -Type "Success"
} else {
  Debug-Log ": FATAL - Build-Tree returned null root node" -Type "Error"
}

$treeFrame.Add($Script:tree)
$win.Add($treeFrame)
Debug-Log ": TreeView created and added to window successfully" -Type "Success"
Set-StatusBar "Ready" -Final

## Global Key Handlers
$top.add_KeyPress({
  param($e)
  $handled = $false
  switch ($e.KeyEvent.Key) {
    ([Terminal.Gui.Key]::F1)  { Show-Modal "Shortcuts" "F1-Help | F2-Password | F3-New | F5-Refresh | F6-Themes | F7-Search | F10-Quit" ; $handled = $true }
    ([Terminal.Gui.Key]::F2)  { Generate-RandomPassword ; $handled = $true }
    ([Terminal.Gui.Key]::F3)  { Show-NewObjectWizard ; $handled = $true }
    ([Terminal.Gui.Key]::F5)  { Refresh-Data -domain $Script:CurrentDomain -RebuildTree ; $handled = $true }
    ([Terminal.Gui.Key]::F6)  { Show-ThemeSelector ; $handled = $true }
    ([Terminal.Gui.Key]::F7)  { Show-ADSearchDialog ; $handled = $true }
    ([Terminal.Gui.Key]::F10) { [Terminal.Gui.Application]::RequestStop() ; $handled = $true }
  }
  $e.Handled = $handled
})

## Debug view tree dump
if ($DebugMode -or $Logging) {
  Debug-Log "=== FULL VIEW TREE DUMP ===" -Type "Info"
  Debug-DumpViewTree -View $top
  Debug-Log "=== END VIEW TREE DUMP ===" -Type "Info"
}

## ===== STEP 8: Run Application =====
Debug-Log ": Starting Terminal.Gui main loop..." -Type "Success"
[Terminal.Gui.Application]::Run($top)

## ===== Cleanup =====
Debug-Log ": Application stopped, cleaning up..." -Type "Info"
Set-StatusBar "Shutting down"
[Terminal.Gui.Application]::Shutdown()
Debug-Log "Application shut down cleanly" -Type "Success"
Debug-Log "End of line..." -Type "Info"

## Gracefully close logs
if ($Script:LogStream) {
  $Script:LogStream.Close()
  $Script:LogStream.Dispose()
}
